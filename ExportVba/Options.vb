Public Class Options

    <CommandLine.Option("d"c, "database", HelpText:="Database path, or if not provided then the database of current directory if exactly 1 database found", Required:=False)>
    Public Property Database As String

    <CommandLine.Option("f"c, "target-folder", Required:=True)>
    Public Property TargetPath As String

    <CommandLine.Option("c"c, "clear", [Default]:=False, Required:=False)>
    Public Property ClearTargetPath As Boolean

    <CommandLine.Option("e"c, "file-extension", [Default]:=".bas", Required:=False)>
    Public Property TargetFilesExtension As String

End Class
