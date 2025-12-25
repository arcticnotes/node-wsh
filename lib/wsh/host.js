// This file is in JScript, not in JavaScript. It is executed in Windows Scripting Host (WSH).

var FSO = new ActiveXObject( 'Scripting.FileSystemObject');
var OBJECT_TOSTRING = Object.toString();
var REFERENCES = {
	'0': WScript, // must match node-wsh.js
	'1': GetObject // must match node-wsh.js
};
var nextRefId = 2;

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
				case 'funref':
					item = REFERENCES[ encoded.value];
					if( item === undefined)
						throw new Error( 'reference not found: ' + encoded.value);
					return item;
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
				if( REFERENCES[ i] === decoded)
					return { type: 'objref', value: i};
			i = '' + nextRefId++;
			REFERENCES[ i] = decoded;
			return { type: 'objref', value: i};
		case 'function':
			for( i in REFERENCES)
				if( REFERENCES[ i] === decoded)
					return { type: 'funref', value: i};
			i = '' + nextRefId++;
			REFERENCES[ i] = decoded;
			return { type: 'funref', value: i};
		case 'symbol':
		case 'bigint':
		default:
			throw new Error( 'unsupported data type: ' + typeof decoded);
	}
}

( function() {
	var input;
	var output;
	while( !WScript.StdIn.AtEndOfLine)
		try {
			input = JSON.parse( WScript.StdIn.ReadLine());
			switch( input[ 0]) {
				case 'unref': // [ 'unref', ref] => [ 'done']
					if( REFERENCES[ input[ 1]] === undefined)
						throw new Error( 'unknown ref: ' + input[ 1]);
					delete REFERENCES[ input[ i]];
					output = [ 'done'];
					break;
				case 'get': // [ 'get', target, prop] => [ 'result', value]
					output = [ 'result', encode( decode( input[ 1])[ decode( input[ 2])])];
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
