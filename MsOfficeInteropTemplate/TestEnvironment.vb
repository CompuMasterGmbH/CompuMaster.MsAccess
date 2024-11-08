﻿Imports NUnit.Framework

Public NotInheritable Class TestEnvironment

#If CI_CD = True Then
    Public Shared Property ConsoleOutputSuppressed As Boolean = True
#Else
    Public Shared Property ConsoleOutputSuppressed As Boolean = False
#End If

    Public Shared Function DirectoryOfTestAssembly() As String
        Return System.IO.Path.GetDirectoryName(GetType(TestEnvironment).Assembly.Location)
    End Function

    Public Shared Function FullPathOfExistingTestFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        If System.IO.File.Exists(Result) = False Then
            Throw New System.IO.FileNotFoundException(Result)
        End If
        Return Result
    End Function

    Public Shared Function FullPathOfDynTestFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Paths.Insert(1, "temp")
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        Dim ParentDir As String = System.IO.Path.GetDirectoryName(Result)
        If System.IO.Directory.Exists(ParentDir) = False Then
            System.IO.Directory.CreateDirectory(ParentDir)
        End If
        If System.IO.File.Exists(Result) Then
            System.IO.File.Delete(Result)
        End If
        Return Result
    End Function

    Public Shared Function FullPathOfDynTestFile_KeepExistingFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Paths.Insert(1, "temp")
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        Dim ParentDir As String = System.IO.Path.GetDirectoryName(Result)
        If System.IO.Directory.Exists(ParentDir) = False Then
            System.IO.Directory.CreateDirectory(ParentDir)
        End If
        Return Result
    End Function

    Public Class TestFiles

        Public Shared ReadOnly Property TestFileNorthwindDatabase As System.IO.FileInfo
            Get
                Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "NorthwindDatabase.accdb"))
            End Get
        End Property

    End Class
End Class