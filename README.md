# CompuMaster.MsAccess
Open MS Access applications and databases with COM interop, but without the need of COM interop assemblies

## Diagnostic logs

AccessApplication exposes DiagnosticLogFiles for its automation process. By default, it removes confirmed logs after a successful quit and trims trailing NUL padding after an abnormal exit. Set PreserveDiagnosticLogFiles to True to retain the files byte for byte.

For external automation, use the Access CurrentData and CurrentProject COM properties to enumerate the open database's objects. The VBE COM property returns the editor object.
