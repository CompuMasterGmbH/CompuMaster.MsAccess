'Public Class DevelopmentTools

'    Dim arrConvertFiles() As String
'    Dim currentIndexOfConvertFiles As Integer

'    Private Application As Global.CompuMaster.MsAccess.AccessApplication

'    Public Sub ExportAllComponents()
'        Dim maxCount As Integer
'        Dim currentVBProject As VBProject
'        Dim currentVBComponent As VBComponent
'        Dim componentFileName As String

'        maxCount = Application.VBE.ActiveVBProject.VBComponents.count
'        arrConvertFilesMaxCount = maxCount
'        ReDim arrConvertFiles(maxCount)
'        currentVBProject = Application.VBE.ActiveVBProject

'        For i = 1 To maxCount
'            'Debug.Print "Exporting " & i & " of " & maxCount & "."
'            currentIndexOfConvertFiles = i
'            currentVBComponent = currentVBProject.VBComponents(i)
'            componentFileName = GetFileNameForComponent(currentVBComponent)
'            ExportVBComponent(currentVBComponent, GetFolderName(currentVBComponent), componentFileName, True, RedirectToolsOutputToDebug)
'        Next

'        Debug.Print("Finished export of all components.")

'        ConvertAllTextFilesToUTF8 arrConvertFiles, maxCount, RedirectToolsOutputToDebug
'End Sub

'    Public Function ExportVBComponent(VBComp As VBComponent,
'                FolderName As String,
'                Optional FileName As String,
'                Optional OverwriteExisting As Boolean = True,
'                Optional RedirectConsoleToDebug As Boolean = False) As Boolean

'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ' This function exports the code module of a VBComponent to a text
'        ' file. If FileName is missing, the code will be exported to
'        ' a file with the same name as the VBComponent followed by the
'        ' appropriate extension.
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'        Dim Extension As String
'        Dim FName As String
'        Extension = GetFileExtension(VBComp:=VBComp)
'        If Trim(FileName) = vbNullString Then
'            FName = VBComp.Name & Extension
'        Else
'            FName = FileName
'            If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
'                FName = FName & Extension
'            End If
'        End If

'        If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
'            FName = FolderName & FName
'        Else
'            FName = FolderName & "\" & FName
'        End If

'        If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
'            If OverwriteExisting = True Then
'                Kill(FName)
'            Else
'                ExportVBComponent = False
'                Exit Function
'            End If
'        End If

'        VBComp.Export(FileName:=FName)
'        arrConvertFiles(currentIndexOfConvertFiles) = FName
'        ExportVBComponent = True

'    End Function

'    Private Function GetFolderName(VBComp As VBComponent) As String
'        Dim rootPath As String
'        Dim subPath As String

'        rootPath = CurrentProject.path & "\Sources\"

'        If VBComp.Name Like "Report_*" Then
'            subPath = "Reports\"
'        Else
'            Select Case VBComp.type
'                Case vbext_ct_ClassModule
'                    subPath = "Classmodules\"
'                Case vbext_ct_Document
'                    subPath = "Forms\"
'                Case vbext_ct_MSForm
'                    subPath = "Forms\"
'                Case vbext_ct_StdModule
'                    subPath = "Modules\"
'                Case Else
'                    subPath = "\"
'            End Select
'        End If

'        GetFolderName = rootPath & subPath
'    End Function

'    Private Function GetFileNameForComponent(VBComp As VBComponent) As String
'        Dim Name As String
'        Name = VBComp.Name

'        Name = Replace(Name, " / ", "_")
'        'Name = Replace(Name, " ", "_")

'        GetFileNameForComponent = Name
'    End Function

'    Private Function GetFileExtension(VBComp As VBComponent) As String
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ' This returns the appropriate file extension based on the Type of
'        ' the VBComponent.
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Select Case VBComp.type
'            Case vbext_ct_ClassModule
'                GetFileExtension = ".cls"
'            Case vbext_ct_Document
'                GetFileExtension = ".cls"
'            Case vbext_ct_MSForm
'                GetFileExtension = ".frm"
'            Case vbext_ct_StdModule
'                GetFileExtension = ".bas"
'            Case Else
'                GetFileExtension = ".bas"
'        End Select
'    End Function

'End Class
