import EventEmitter from 'node:events';
import PATH from 'node:path';
import { Syncline} from "@arcticnotes/syncline";

const COMMAND = 'cscript.exe';
const ARGS = [ '//E:jscript', '//NoLogo'];
const SCRIPT_FILE = PATH.join( PATH.dirname( import.meta.dirname), 'wsh', 'host.js');
const TRACE_REF = 1;

export class WindowsScriptingHost extends EventEmitter {

	static async connect( options = {}) {
		const command = options.command || COMMAND;
		const args = options.args || ARGS;
		const scriptFile = options.scriptFile || SCRIPT_FILE;
		const trace = options.trace || 0;
		return new WindowsScriptingHost( await Syncline.spawn( command, [ ...args, scriptFile], { trace}), options);
	}

	#syncline;
	#trace;
	#finalizer = new FinalizationRegistry( this.#finalized.bind( this));
	#ref2proxy = new Map();
	#proxy2ref = new WeakMap();

	#objectHandler = {

		wsh: this,
	
		get( target, prop) {
			if( prop === Symbol.toPrimitive)
				return () => `ref#${ target.ref}`;
			const encodedTarget = this.wsh.#encode( target.proxy);
			const encodedProp = this.wsh.#encode( prop);
			const output = JSON.parse( this.wsh.#syncline.exchange( JSON.stringify( [ 'get', encodedTarget, encodedProp])));
			switch( output[ 0]) {
				case 'value': return this.wsh.#decode( output[ 1]);
				case 'error': throw new Error( output[ 1]);
				default: throw new Error( `unknown status: ${ output[ 0]}`);
			}
		},

		set( target, prop, value) {
			const encodedTarget = this.wsh.#encode( target.proxy);
			const encodedProp = this.wsh.#encode( prop);
			const encodedValue = this.wsh.#encode( value);
			const output = JSON.parse( this.wsh.#syncline.exchange( JSON.stringify( [ 'set', encodedTarget, encodedProp, encodedValue])));
			switch( output[ 0]) {
				case 'set': return;
				case 'error': throw new Error( output[ 1]);
				default: throw new Error( `unknown status: ${ output[ 0]}`);
			}
		},
	};

	#functionHandler = {

		wsh: this,

		get( target, prop) {
			if( prop === Symbol.toPrimitive)
				return () => `ref#${ target.ref}`;
			return undefined;
		},

		apply( target, thisArg, argumentList) {
			const encodedTarget = this.wsh.#encode( target.proxy);
			const encodedThisArg = this.wsh.#encode( thisArg);
			const encodedArgumentList = this.wsh.#encode( argumentList);
			const output = JSON.parse( this.wsh.#syncline.exchange( JSON.stringify( [ 'apply', encodedTarget, encodedThisArg, encodedArgumentList])));
			switch( output[ 0]) {
				case 'value': return this.wsh.#decode( output[ 1]);
				case 'error': throw new Error( output[ 1]);
				default: throw new Error( `unknown status: ${ output[ 0]}`);
			}
		},

		construct( target, argumentList) {
			const encodedTarget = this.wsh.#encode( target.proxy);
			const encodedArgumentList = this.wsh.#encode( argumentList);
			const output = JSON.parse( this.wsh.#syncline.exchange( JSON.stringify( [ 'construct', encodedTarget, encodedArgumentList])));
			switch( output[ 0]) {
				case 'value': return this.wsh.#decode( output[ 1]);
				case 'error': throw new Error( output[ 1]);
				default: throw new Error( `unknown status: ${ output[ 0]}`);
			}
		},
	};

	#WScript;
	#GetObject;
	#Enumerator;

	constructor( syncline, options) {
		super();
		this.#syncline = syncline;
		this.#syncline.on( 'stderr', line => console.log( 'wsh:', line));
		this.#syncline.on( 'stdout', line => console.log( 'wsh:', line));
		this.#trace = options.trace;
		this.#WScript = this.#getOrCreateObject( 0);
		this.#GetObject = this.#getOrCreateFunction( 1);
		this.#Enumerator = this.#getOrCreateFunction( 2);
	}

	get remoteObjects() {
		const proxies = this;
		return {
			get count() {
				return proxies.#ref2proxy.size;
			},
		};
	}

	get WScript() {
		return this.#WScript;
	}

	get GetObject() {
		return this.#GetObject;
	}

	get Enumerator() {
		return this.#Enumerator;
	}

	async disconnect() {
		await this.#syncline.close();
	}

	#getOrCreateObject( ref) {
		return this.#getOrCreate( ref, this.#objectHandler);
	}

	#getOrCreateFunction( ref) {
		return this.#getOrCreate( ref, this.#functionHandler);
	}

	#getOrCreate( ref, handler) {
		const existingWeakRef = this.#ref2proxy.get( ref);
		const existingProxy = existingWeakRef && existingWeakRef.deref();
		if( existingProxy)
			return existingProxy;

		const target = handler === this.#objectHandler? new RemoteObject( ref): function() {};
		const newProxy = new Proxy( target, handler);
		target.proxy = newProxy;
		this.#ref2proxy.set( ref, new WeakRef( newProxy)); // may be overwriting a dead WeakRef
		this.#proxy2ref.set( newProxy, ref);
		this.#finalizer.register( newProxy, ref);
		this.emit( 'ref', ref, newProxy);
		return newProxy;
	}

	#finalized( ref) {
		if( this.#ref2proxy.get( ref).deref() === undefined) // otherwise, it's overwritten by a refreshed proxy
			this.#ref2proxy.delete( ref);
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
		this.emit( 'unref', ref);
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
					if( objref === undefined) // not because garbage-collected, because clearly `decoded` is still alive
						throw new Error( `remote object reference not found: ${ decoded}`);
					return { type: 'objref', value: objref};
				}
				const encoded = { type: 'object', value: {}};
				for( const [ name, value] of Object.entries( decoded))
					encoded.value[ name] = this.#encode( value);
				return encoded;
			case 'function':
				const funref = this.#proxy2ref.get( decoded);
				if( funref === undefined) // not because garbage-collected, because clearly `decoded` is still alive
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
						return this.#getOrCreateObject( encoded.value);
					case 'funref':
						return this.#getOrCreateFunction( encoded.value);
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

	#ref;
	#proxy;

	constructor( ref) {
		this.#ref = ref;
	}

	get ref() {
		return this.#ref;
	}

	get proxy() {
		return this.#proxy;
	}

	set proxy( proxy) {
		this.#proxy = proxy;
	}
}
