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
		console.log( GetObject( "winmgmts:\\\\.\\root\\cimv2"));
	} finally {
		await wsh.disconnect();
	}
});
