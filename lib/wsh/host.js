// This file is in JScript, not in JavaScript. It is executed in Windows Scripting Host (WSH).

var FSO = new ActiveXObject( 'Scripting.FileSystemObject');
var OBJECT_TOSTRING = Object.toString();
var GLOBAL = this;
var REFERENCES = {}
var nextRefId = 0;

eval( FSO.OpenTextFile( FSO.BuildPath( FSO.GetParentFolderName( WScript.ScriptFullName), 'json2.js')).ReadAll());

function decode( encoded) {
	var decoded;
	var item;
	var i;
	switch( typeof encoded) {
		case 'boolean':
		case 'number':
		case 'string':
			return encoded;
		case 'object':
			if( encoded === null)
				return encoded;
			if( encoded instanceof Array) {
				decoded = [];
				for( i = 0; i < encoded.length; i++)
					decoded.push( decode( encoded[ i]));
				return decoded;
			}
			switch( encoded.type) {
				case 'undefined':
					return undefined;
				case 'object':
					decoded = {};
					for( i in encoded.value)
						decoded[ i] = decode( encoded.value[ i]);
					return decoded;
				case 'objref':
					item = REFERENCES[ encoded.value];
					if( item === undefined)
						throw new Error( 'reference not found: ' + encoded.value);
					if( item.type !== 'obj')
						throw new Error( 'reference type mismatch: ' + encoded.value);
					return item.value;
				case 'funref':
					item = REFERENCES[ encoded.value];
					if( item === undefined)
						throw new Error( 'reference not found: ' + encoded.value);
					if( item.type === 'potential-method')
						throw new Error( 'potentially a method, cannot be evaluated standalone: ' + encoded.value);
					if( item.type !== 'fun')
						throw new Error( 'reference type mismatch: ' + encoded.value);
					return item.value;
				default:
					throw new Error( 'unknown object type: ' + encoded.type);
			}
		case 'undefined':
		case 'symbol':
		case 'bigint':
		case 'function':
		default:
			throw new Error( 'illegal data type: ' + typeof encoded);
	}
}

function encode( decoded) {
	var encoded;
	var i;
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
				encoded = [];
				for( i = 0; i < decoded.length; i++)
					encoded.push( encode( decoded[ i]));
				return encoded;
			}
			if( decoded.constructor && decoded.constructor.toString() === OBJECT_TOSTRING) {
				encoded = { type: 'object', value: {}};
				for( i in decoded)
					encoded.value[ i] = encode( decoded[ i]);
				return encoded;
			}
			for( i in REFERENCES)
				if( REFERENCES[ i].value === decoded)
					return { type: 'objref', value: i};
			i = '' + nextRefId++;
			REFERENCES[ i] = { type: 'obj', value: decoded};
			return { type: 'objref', value: i};
		case 'function':
			for( i in REFERENCES)
				if( REFERENCES[ i].value === decoded)
					return { type: 'funref', value: i};
			i = '' + nextRefId++;
			REFERENCES[ i] = { type: 'fun', value: decoded};
			return { type: 'funref', value: i};
		case 'symbol':
		case 'bigint':
		default:
			throw new Error( 'unsupported data type: ' + typeof decoded);
	}
}

function encodePotentialMethod( target, prop) {
	var i;
	var item;
	for( i in REFERENCES) {
		item = REFERENCES[ i];
		if( item.type === 'potential-method' && item.target === target && item.prop === prop)
			return { type: 'funref', value: i};
	}
	i = '' + nextRefId++;
	REFERENCES[ i] = { type: 'potential-method', target: target, prop: prop};
	return { type: 'funref', value: i};
}

function decodePotentialMethod( encoded) {
	var item;
	if( typeof encoded === 'object' && encoded.type === 'funref') {
		item = REFERENCES[ encoded.value];
		if( item.type === 'potential-method')
			return { 'potential-method': true, target: item.target, prop: item.prop};
	}
	return { 'potential-method': false, value: decode( encoded)}
}

( function() {
	var input;
	var output;
	var target;
	var prop;
	var value;
	var thisArg;
	var args;
	while( !WScript.StdIn.AtEndOfLine)
		try {
			input = JSON.parse( WScript.StdIn.ReadLine());
			switch( input[ 0]) {
				case 'global': // [ 'global', name] => [ 'value', value]
					output = [ 'value', encode( GLOBAL[ input[ 1]])];
					break;
				case 'unref': // [ 'unref', ref] => [ 'done']
					if( REFERENCES[ input[ 1]] === undefined)
						throw new Error( 'unknown ref: ' + input[ 1]);
					delete REFERENCES[ input[ i]];
					output = [ 'done'];
					break;
				case 'get': // [ 'get', target, prop] => [ 'value', value] | [ 'potential-method']
					target = decode( input[ 1]);
					prop = decode( input[ 2]);
					try {
						output = [ 'value', encode( target[ prop])];
					} catch( error) {
						// could be a method
						if( typeof target[ prop] !== 'unknown')
							throw error;
						output = [ 'value', encodePotentialMethod( target, prop)];
					}
					break;
				case 'set': // [ 'set', target, prop, value] => [ 'set']
					target = decode( input[ 1]);
					prop = decode( input[ 2]);
					value = decode( input[ 3]);
					target[ prop] = value;
					output = [ 'set'];
					break;
				case 'apply': // [ 'apply', target, thisArg, argumentList] => [ 'value', value]
					target = decodePotentialMethod( input[ 1]);
					thisArg = decode( input[ 2]);
					args = decode( input[ 3]);
					if( target[ 'potential-method']) {
						if( thisArg === undefined)
							throw new Error( 'potentially a method, use with a "this"');
						if( thisArg !== target.target)
							throw new Error( 'potentially a method, "this" has changed');
						switch( args.length) {
							case 0: output = [ 'value', encode( target.target[ target.prop]())]; break;
							case 1: output = [ 'value', encode( target.target[ target.prop]( args[ 0]))]; break;
							case 2: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1]))]; break;
							case 3: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2]))]; break;
							case 4: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3]))]; break;
							case 5: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4]))]; break;
							case 6: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5]))]; break;
							case 7: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6]))]; break;
							case 8: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6], args[ 7]))]; break;
							case 9: output = [ 'value', encode( target.target[ target.prop]( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6], args[ 7], args[ 8]))]; break;
							default: throw new Error( 'too many arguments');
						}
					} else
						output = [ 'value', encode( target.value.apply( thisArg, args))];
					break;
				case 'construct': // [ 'construct', target, argumentList] => [ 'value', value]
					target = decode( input[ 1]);
					args = decode( input[ 2]);
					switch( args.length) {
						case 0: output = [ 'value', encode( new target())]; break;
						case 1: output = [ 'value', encode( new target( args[ 0]))]; break;
						case 2: output = [ 'value', encode( new target( args[ 0], args[ 1]))]; break;
						case 3: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2]))]; break;
						case 4: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3]))]; break;
						case 5: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4]))]; break;
						case 6: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5]))]; break;
						case 7: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6]))]; break;
						case 8: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6], args[ 7]))]; break;
						case 9: output = [ 'value', encode( new target( args[ 0], args[ 1], args[ 2], args[ 3], args[ 4], args[ 5], args[ 6], args[ 7], args[ 8]))]; break;
						default: throw new Error( 'too many arguments');
					}
					break;
				default:
					throw new Error( 'unknown command: ' + input[ 0]);
			}
		} catch( error) {
			output = [ 'error', '' + error.message];
		} finally {
			WScript.StdOut.WriteLine( JSON.stringify( output));
		}
})();
