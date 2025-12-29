import EventEmitter from 'node:events';
import PATH from 'node:path';
import { Syncline} from "@arcticnotes/syncline";

const COMMAND = 'cscript.exe';
const ARGS = [ '//NoLogo'];
const SCRIPT_FILE = PATH.join( PATH.dirname( import.meta.dirname), 'wsh', 'host.wsf');

export class WindowsScriptingHost extends EventEmitter {

	static async connect( options = {}) {
		const command = options.command || COMMAND;
		const args = options.args || ARGS;
		const scriptFile = options.scriptFile || SCRIPT_FILE;
		const trace = options.trace || 0;
		return new WindowsScriptingHost( await Syncline.spawn( command, [ ...args, scriptFile], { trace}));
	}

	#syncline;
	#closed = false;
	#finalizer = new FinalizationRegistry( this.#finalized.bind( this));
	#ref2proxy = new Map(); // Map< string, WeakRef< Proxy | <custom-mapped>>>
	#proxy2ref = new WeakMap(); // Map< Proxy | <custom-mapped>, string>

	constructor( syncline) {
		super();
		this.#syncline = syncline;
		this.#syncline.on( 'stderr', line => console.log( 'wsh-stderr:', line));
		this.#syncline.on( 'stdout', line => console.log( 'wsh-stdout:', line));
	}

	get remoteObjects() {
		const proxies = this;
		return {
			get count() {
				return proxies.#ref2proxy.size;
			},
		};
	}

	global( name, mapper = undefined) {
		const output = JSON.parse( this.#syncline.exchange( JSON.stringify( [ 'global', name])));
		switch( output[ 0]) {
			case 'value': return this.#decode( output[ 1], mapper);
			case 'error': throw new Error( output[ 1]);
			default: throw new Error( `unknown status: ${ output[ 0]}`); // bug
		}
	}

	async disconnect() {
		this.#closed = true;
		await this.#syncline.close();
	}

	#decode( encoded, mapper) {
		switch( typeof encoded) {
			case 'boolean':
			case 'number':
			case 'string':
				return encoded;
			case 'object':
				if( encoded === null)
					return encoded;
				if( Array.isArray( encoded)) {
					const decoded = [];
					for( const item of encoded)
						decoded.push( this.#decode( item, mapper));
					return decoded;
				}
				switch( encoded.type) {
					case 'undefined':
						return undefined;
					case 'object':
						const decoded = {};
						for( const [ name, value] of Object.entries( encoded.value))
							decoded[ name] = this.#decode( value, mapper);
						return decoded;
					case 'ref':
						return this.#getOrCreate( encoded.value, mapper);
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

	#encode( decoded) {
		switch( typeof decoded) {
			case 'boolean':
			case 'number':
			case 'string':
				return decoded;
			case 'undefined':
				return { type: 'undefined'};
			case 'object': {
				if( decoded === null)
					return decoded;
				const ref = this.#proxy2ref.get( decoded);
				if( ref !== undefined)
					return { type: 'ref', value: ref};
				if( Array.isArray( decoded)) {
					const encoded = [];
					for( const item of decoded)
						encoded.push( this.#encode( item));
					return encoded;
				}
				const encoded = { type: 'object', value: {}};
				for( const [ name, value] of Object.entries( decoded))
					encoded.value[ name] = this.#encode( value);
				return encoded;
			}
			case 'function': {
				const ref = this.#proxy2ref.get( decoded);
				if( ref !== undefined)
					return { type: 'ref', value: ref};
				throw new Error( `functions from node cannot be sent: ${ decoded}`);
			}
			case 'bigint':
			case 'symbol':
			default:
				throw new Error( `unsupported value: ${ decoded}`);
		}
	}

	#getOrCreate( ref, mapper) {
		const existingWeakRef = this.#ref2proxy.get( ref);
		const existingProxy = existingWeakRef && existingWeakRef.deref();
		if( existingProxy)
			return existingProxy;

		const newProxy = new RemoteObject( this.#syncline, this.#encode.bind( this), this.#decode.bind( this), ref, mapper).proxy;
		if( this.#proxy2ref.has( newProxy)) // sanity check, doesn't catch all misbehaving mappers
			throw new Error( `mapper must return new objects: ${ newProxy}`);
		this.#ref2proxy.set( ref, new WeakRef( newProxy)); // may be overwriting a dead WeakRef
		this.#proxy2ref.set( newProxy, ref);
		this.#finalizer.register( newProxy, ref);
		this.emit( 'ref', ref, newProxy);
		return newProxy;
	}

	#finalized( ref) {
		if( this.#ref2proxy.get( ref).deref() === undefined) // otherwise, it's overwritten by a refreshed proxy
			this.#ref2proxy.delete( ref);
		if( !this.#closed)
			this.#syncline.exchange( JSON.stringify( [ 'unref', ref]));
		this.emit( 'unref', ref);
	}
}

class RemoteObject extends Function {

	static #handler = {
		get: ( ticket, prop) => ticket.get( prop, undefined),
		set: ( ticket, prop, value) => ticket.set( prop, value),
		apply: ( ticket, thisArg, argumentsList) => ticket.apply( thisArg, argumentsList, undefined),
		construct: ( ticket, argumentsList) => ticket.construct( argumentsList, undefined),
	};

	#syncline;
	#encode;
	#decode;
	#ref;
	#proxy;

	constructor( syncline, encode, decode, ref, mapper) {
		super();
		this.#syncline = syncline;
		this.#encode = encode;
		this.#decode = decode;
		this.#ref = ref;
		this.#proxy = mapper? mapper( this): this.newProxy();
	}

	get proxy() {
		return this.#proxy;
	}

	newProxy() {
		return new Proxy( this, RemoteObject.#handler);
	}

	get( prop, mapper) {
		if( prop === Symbol.toPrimitive)
			return () => `ref#${ this.#ref}`; 
		const encodedTarget = this.#encode( this.#proxy);
		const encodedProp = this.#encode( prop);
		const output = JSON.parse( this.#syncline.exchange( JSON.stringify( [ 'get', encodedTarget, encodedProp])));
		switch( output[ 0]) {
			case 'value': return this.#decode( output[ 1], mapper);
			case 'error': throw new Error( output[ 1]);
			default: throw new Error( `unknown status: ${ output[ 0]}`);
		}
	}

	set( prop, value) {
		const encodedTarget = this.#encode( this.#proxy);
		const encodedProp = this.#encode( prop);
		const encodedValue = this.#encode( value);
		const output = JSON.parse( this.#syncline.exchange( JSON.stringify( [ 'set', encodedTarget, encodedProp, encodedValue])));
		switch( output[ 0]) {
			case 'set': return true;
			case 'error': throw new Error( output[ 1]);
			default: throw new Error( `unknown status: ${ output[ 0]}`);
		}
	}

	apply( thisArg, argumentsList, mapper) {
		const encodedTarget = this.#encode( this.#proxy);
		const encodedThisArg = this.#encode( thisArg);
		const encodedArgumentList = this.#encode( [ ...argumentsList]); // argumentsList may not be instanceof Array
		const output = JSON.parse( this.#syncline.exchange( JSON.stringify( [ 'apply', encodedTarget, encodedThisArg, encodedArgumentList])));
		switch( output[ 0]) {
			case 'value': return this.#decode( output[ 1], mapper);
			case 'error': throw new Error( output[ 1]);
			default: throw new Error( `unknown status: ${ output[ 0]}`);
		}
	}

	construct( argumentsList, mapper) {
		const encodedTarget = this.#encode( this.#proxy);
		const encodedArgumentList = this.#encode( [ ...argumentsList]); // argumentsList may not be instanceof Array
		const output = JSON.parse( this.#syncline.exchange( JSON.stringify( [ 'construct', encodedTarget, encodedArgumentList])));
		switch( output[ 0]) {
			case 'value': return this.#decode( output[ 1], mapper);
			case 'error': throw new Error( output[ 1]);
			default: throw new Error( `unknown status: ${ output[ 0]}`);
		}
	}
}
