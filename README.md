
# Node-WSH Bridge

This is a Node.js package that runs Windows Scripting Host (WSH) as a sub-process and exposes the resources from the
WSH world to the Node.js world through serialized communication over the standard input and output channels. Obviously,
it only works on Windows.

```
+-------------+
|  node.exe   |
+-------------+
       |
       v
+-------------+
| cscript.exe |
+-------------+
```
