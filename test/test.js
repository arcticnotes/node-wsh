import ASSERT from 'node:assert/strict';
import TEST from 'node:test';
import { WindowsScriptingHost} from '@arcticnotes/node-wsh';

TEST( 'smoke-test', async() => {
	const wsh = await WindowsScriptingHost.connect();
	try {
		const { WScript, GetObject, Enumerator} = wsh;
		console.log( WScript.Version);
		ASSERT.equal( typeof WScript.Version, 'string');
		const procs = GetObject( "winmgmts:\\\\.\\root\\cimv2").ExecQuery( 'SELECT ProcessId, Name FROM Win32_Process');
		for( const enumerator = new Enumerator( procs); !enumerator.atEnd(); enumerator.moveNext()) {
			console.log( `${ enumerator.item().ProcessId}: ${ enumerator.item().Name}`);
			ASSERT.equal( typeof enumerator.item().ProcessId, 'number');
			ASSERT.equal( typeof enumerator.item().Name, 'string');
		}
	} finally {
		await wsh.disconnect();
	}
});
