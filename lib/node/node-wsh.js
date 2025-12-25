import PATH from 'node:path';
import { Syncline } from "@arcticnotes/syncline";

const COMMAND = 'cscript.exe';
const ARGS = [ '//E:jscript', '//NoLogo'];
const SCRIPT_FILE = PATH.join( PATH.dirname( import.meta.dirname), 'wsh', 'host.js');
const REF_WSCRIPT = '0'; // must match host.js
const REF_GETOBJECT = '1'; // must match host.js
const PROXY = Symbol();
const TRACE_REF = 1;

export class WindowsScriptingHost {

	static async connect( options = {}) {
		const command = options.command || COMMAND;
		const args = options.args || ARGS;
		const scriptFile = options.scriptFile || SCRIPT_FILE;
		const trace = options.trace || 0;
		return new WindowsScriptingHost( await Syncline.spawn( command, [ ...args, scriptFile], { trace}), options);
	}

	#syncline;
	#proxies;
	#WScript;
	#GetObject;

	constructor( syncline, options) {
		this.#syncline = syncline;
		this.#syncline.on( 'stderr', line => console.log( 'wsh:', line));
		this.#syncline.on( 'stdout', line => console.log( 'wsh:', line));
		this.#proxies = new Proxies( syncline, options);
		this.#WScript = this.#proxies.getOrCreateObject( REF_WSCRIPT);
		this.#GetObject = this.#proxies.getOrCreateFunction( REF_GETOBJECT);
	}

	get WScript() {
		return this.#WScript;
	}

	get GetObject() {
		return this.#GetObject;
	}

	async disconnect() {
		await this.#syncline.close();
	}
}

class Proxies {

	#syncline;
	#trace;
	#finalizer = new FinalizationRegistry( this.#finalized.bind( this));
	#ref2proxy = new Map();
	#proxy2ref = new Map();

	#objectHandler = {

		proxies: this,
	
		get( target, prop, receiver) {
			if( prop === Symbol.toPrimitive)
				return () => `ref#${ this.proxies.#proxy2ref.get( target[ PROXY])}`;
			const encodedTarget = this.proxies.#encode( target[ PROXY]);
			const encodedProp = this.proxies.#encode( prop);
			const output = JSON.parse( this.proxies.#syncline.exchange( JSON.stringify( [ 'get', encodedTarget, encodedProp])));
			switch( output[ 0]) {
				case 'result': return this.proxies.#decode( output[ 1]);
				case 'error': throw new Error( output[ 1]);
				default: throw new Error( `unknown status: ${ output[ 0]}`);
			}
		}
	};

	#functionHandler = {
	};

	constructor( syncline, options) {
		this.#syncline = syncline;
		this.#trace = options.trace;
	}

	getOrCreateObject( ref) {
		return this.#getOrCreate( ref, this.#objectHandler);
	}

	getOrCreateFunction( ref) {
		return this.#getOrCreate( ref, this.#functionHandler);
	}

	#getOrCreate( ref, handler) {
		const existingProxy = this.#ref2proxy.get( ref);
		if( existingProxy)
			return existingProxy;

		const target = handler === this.#objectHandler? new RemoteObject(): function() {};
		const newProxy = new Proxy( target, handler);
		target[ PROXY] = newProxy;
		this.#ref2proxy.set( ref, newProxy);
		this.#proxy2ref.set( newProxy, ref);
		this.#finalizer.register( newProxy, ref);
		return newProxy;
	}

	#finalized( ref) {
		const output = this.#syncline.exchange( JSON.stringify( [ 'unref', ref]));
		if( this.#trace >= TRACE_REF)
			switch( output[ 0]) {
				case 'error':
					console.log( `failed to unref: ${ ref}`);
					break;
				case 'done':
					console.log( `unreferenced: ${ ref}`);
					break;
				default:
					console.log( `unknown response: ${ output[ 0]}`);
			}
	}

	#encode( decoded) {
		switch( typeof decoded) {
			case 'boolean':
			case 'number':
			case 'string':
				return decoded;
			case 'undefined':
				return { type: 'undefined'};
			case 'object':
				if( decoded === null)
					return decoded;
				if( decoded instanceof Array) {
					const encoded = [];
					for( const item of decoded)
						encoded.push( this.#encode( item));
					return encoded;
				}
				if( decoded instanceof RemoteObject) {
					const objref = this.#proxy2ref.get( decoded);
					if( objref === undefined)
						throw new Error( `remote object reference not found: ${ decoded}`);
					return { type: 'objref', value: objref};
				}
				const encoded = { type: 'object', value: {}};
				for( const [ name, value] of Object.entries( decoded))
					encoded.value[ name] = this.#encode( value);
				return encoded;
			case 'function':
				const funref = this.#proxy2ref.get( decoded);
				if( funref === undefined)
					throw new Error( `functions from node are disallowed: ${ decoded}`);
				return { type: 'funref', value: funref};
			case 'bigint':
			case 'symbol':
			default:
				throw new Error( `unsupported value: ${ decoded}`);
		}
	}

	#decode( encoded) {
		switch( typeof encoded) {
			case 'boolean':
			case 'number':
			case 'string':
				return encoded;
			case 'object':
				if( encoded === null)
					return encoded;
				if( encoded instanceof Array) {
					const decoded = [];
					for( const item of encoded)
						decoded.push( this.#decode( item));
					return decoded;
				}
				switch( encoded.type) {
					case 'undefined':
						return undefined;
					case 'object':
						const decoded = {};
						for( const [ name, value] of Object.entries( encoded.value))
							decoded[ name] = this.#decode( value);
						return decoded;
					case 'objref':
						return this.getOrCreateObject( encoded.value);
					case 'funref':
						return this.getOrCreateFunction( encoded.value);
					default:
						throw new Error( `illegal value: ${ encoded}`);
				}
			case 'undefined':
			case 'function':
			case 'bigint':
			case 'symbol':
			default:
				throw new Error( `illegal value: ${ encoded}`);
		}
	}
}

class RemoteObject {
}
