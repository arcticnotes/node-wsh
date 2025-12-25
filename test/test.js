import ASSERT from 'node:assert/strict';
import TEST from 'node:test';
import { WindowsScriptingHost} from '@arcticnotes/node-wsh';

TEST( 'smoke-test', async() => {
	const wsh = await WindowsScriptingHost.connect();
	try {
		const WScript = wsh.WScript;
		const GetObject = wsh.GetObject;
		console.log( WScript.Version);
		ASSERT.equal( typeof WScript.Version, 'string');
		const cimv2 = GetObject( "winmgmts:\\\\.\\root\\cimv2");
		console.log( 'cimv2=', cimv2);
		const procs = cimv2.ExecQuery( 'SELECT ProcessId, Name FROM Win32_Process');
		console.log( 'procs=', procs);
	} finally {
		await wsh.disconnect();
	}
});
