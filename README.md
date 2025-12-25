
# Node-WSH Bridge

This is a Node.js libary that runs Windows Scripting Host (WSH) as a child process and exposes the resources from the
WSH world to the Node.js world through serialized communication over the standard input and output channels. Obviously,
it only works on Windows.

## Usage

To install:

```console
$ npm install @arcticnotes/node-wsh
```

In JavaScript code:

```javascript
import {WindowsScriptingHost} from '@arcticnotes/node-wsh';

const WSH = await WindowsScriptingHost.connect();
const WScript = WSH.global( 'WScript');
console.log(WScript.Version);
await WSH.disconnect();
```

## Dependencies

*  Node.js release 20.x or higher is required.
*  Windows Scripting Host is required. Version compatibility is unclear. Please report issues.
   *  Version 5.812 is known to work.
*  Dependency libraries included by `package.json`:
   *  `@arcticnotes/syncline`. See https://github.com/arcticnotes/syncline.

## License

This library is shared under the [MIT License](https://opensource.org/license/mit). See the `LICENSE` file for details.

This library includes `json2.js`, an implementation of JSON, donated by Douglas Crockford (the designer of JSON) to the
public domain. The license aforementioned does not apply to this file.
