Imports System

Module Program

    Sub Main(args As String())
        Console.WriteLine("Current directory: " & System.Environment.CurrentDirectory)
        Console.WriteLine()

        Dim ParserResult = CommandLine.Parser.Default.ParseArguments(Of Options)(args)
        If ParserResult.Value Is Nothing Then
            ShowHelpText()
            Return
        End If
        Dim LaunchOptions As Options = ParserResult.Value

        If LaunchOptions.Database = Nothing Then
            Dim PotentialDBs As String() = System.IO.Directory.GetFiles(System.Environment.CurrentDirectory, "*.accdb")
            If PotentialDBs.Count = 1 Then
                LaunchOptions.Database = PotentialDBs(0)
                System.Console.WriteLine("Found database at " & LaunchOptions.Database)
            Else
                ShowHelpText()
                Return
            End If
        End If
        LaunchOptions.TargetPath = System.IO.Path.Combine(System.Environment.CurrentDirectory, LaunchOptions.TargetPath)
        If System.IO.Directory.Exists(LaunchOptions.TargetPath) = False Then
            System.IO.Directory.CreateDirectory(LaunchOptions.TargetPath)
        End If
        If LaunchOptions.ClearTargetPath Then
            ClearFilesAndEmptyDirs(LaunchOptions.TargetPath, "*" & LaunchOptions.TargetFilesExtension, False)
        End If
        Console.WriteLine("Export directory: " & LaunchOptions.TargetPath)
        Console.WriteLine()
        ExportVBA.ExportModules(LaunchOptions.Database,
                                System.IO.Path.Combine(System.Environment.CurrentDirectory, LaunchOptions.TargetPath),
                                LaunchOptions.TargetFilesExtension,
                                Sub(fullPath As String, filename As String)
                                    Console.WriteLine("Exported: " & filename)
                                End Sub)

    End Sub

    Public Sub ShowHelpText()
        Console.WriteLine("Invalid arguments")
    End Sub

    Public Sub ClearFilesAndEmptyDirs(folderName As String, searchPattern As String, removeThisFolderIfEmpty As Boolean)
        For Each SubDir In System.IO.Directory.GetDirectories(folderName)
            ClearFilesAndEmptyDirs(SubDir, searchPattern, True)
        Next
        For Each Item In System.IO.Directory.GetFiles(folderName, searchPattern)
            System.IO.File.Delete(Item)
        Next
        If System.IO.Directory.GetDirectories(folderName).Length = 0 AndAlso System.IO.Directory.GetFiles(folderName).Length = 0 Then
            System.IO.Directory.Delete(folderName)
        End If
    End Sub

End Module
