Imports CompuMaster.ComInterop
Imports System.IO

Public Class AccessApplication
    Inherits ComObjectBase

    Public Sub New()
        MyBase.New(Nothing, CreateObject("Access.Application"))
    End Sub

    Public Property UserControl As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("UserControl")
        End Get
        Set(value As Boolean)
            InvokePropertySet("UserControl", value)
        End Set
    End Property

    Public Property Visible As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Visible")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Visible", value)
        End Set
    End Property

    Private _DBEngine As AccessDBEngine
    Public Function DBEngine() As AccessDBEngine
        If _DBEngine Is Nothing OrElse _DBEngine.IsClosedComObject Then
            _DBEngine = New AccessDBEngine(Me, Me)
        End If
        Return _DBEngine
    End Function

    Public ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Private _CurrentDb As AccessDatabase
    Public ReadOnly Property CurrentDb() As AccessDatabase
        Get
            Return _CurrentDb
        End Get
    End Property
    Friend Sub SetCurrentDb(value As AccessDatabase)
        _CurrentDb = value
        _CurrentProject = value
    End Sub

    Private _CurrentProject As AccessDatabase
    Public ReadOnly Property CurrentProject() As AccessDatabase
        Get
            Return _CurrentProject
        End Get
    End Property

    Private _Modules As AccessModuleCollection
    Public ReadOnly Property Modules() As AccessModuleCollection
        Get
            If _Modules Is Nothing Then
                _Modules = New AccessModuleCollection(Me, Me.InvokePropertyGet(Of Object)("Modules"))
            End If
            Return _Modules
        End Get
    End Property

    Private _CodeData As CodeDataCollection
    Public ReadOnly Property CodeData() As CodeDataCollection
        Get
            If _CodeData Is Nothing Then
                _CodeData = New CodeDataCollection(Me, Me.InvokePropertyGet(Of Object)("CodeData"))
            End If
            Return _CodeData
        End Get
    End Property

    Private _CodeProject As AccessCodeProject
    Public ReadOnly Property CodeProject() As AccessCodeProject
        Get
            If _CodeProject Is Nothing Then
                _CodeProject = New AccessCodeProject(Me, Me.InvokePropertyGet(Of Object)("CodeProject"))
            End If
            Return _CodeProject
        End Get
    End Property

    'Private _CodeProject As CodeProjectCollection
    'Public ReadOnly Property CodeProject() As CodeProjectCollection
    '    Get
    '        If _CodeProject Is Nothing Then
    '            _CodeProject = New CodeProjectCollection(Me, Me.InvokePropertyGet(Of Object)("CodeProject"))
    '        End If
    '        Return _CodeProject
    '    End Get
    'End Property

    Private _VBE As VBECollection
    Public ReadOnly Property VBE() As VBECollection
        Get
            If _VBE Is Nothing Then
                _VBE = New VBECollection(Me, Me.InvokePropertyGet(Of Object)("VBE"))
            End If
            Return _VBE
        End Get
    End Property

    Private _VBComponent As VBComponentCollection
    Public ReadOnly Property VBComponent() As VBComponentCollection
        Get
            If _VBComponent Is Nothing Then
                _VBComponent = New VBComponentCollection(Me, Me.InvokePropertyGet(Of Object)("VBComponent"))
            End If
            Return _VBComponent
        End Get
    End Property


    Private _VBProject As VBProjectCollection
    Public ReadOnly Property VBProject() As VBProjectCollection
        Get
            If _VBProject Is Nothing Then
                _VBProject = New VBProjectCollection(Me, Me.InvokePropertyGet(Of Object)("VBProject"))
            End If
            Return _VBProject
        End Get
    End Property

    Public Function Run(Of T)(vbaMethod As String) As T
        Me.InvokeFunction(Of T)("Run", vbaMethod)
    End Function

    Public Sub Close()
        Me.Quit()
    End Sub

    Public Sub Quit()
        If Not IsDisposedComObject Then
            MyBase.CloseAndDisposeChildrenAndComObject()
        End If
    End Sub

    Protected Overrides Sub OnDisposeChildren()
        If Me._DBEngine IsNot Nothing Then Me._DBEngine.Dispose()
        If Me._CurrentDb IsNot Nothing Then Me._CurrentDb.Dispose()
        If Me._Modules IsNot Nothing Then Me._Modules.Dispose()
        If Me._CodeData IsNot Nothing Then Me._CodeData.Dispose()
        If Me._CodeProject IsNot Nothing Then Me._CodeProject.Dispose()
        If Me._VBE IsNot Nothing Then Me._VBE.Dispose()
        If Me._VBComponent IsNot Nothing Then Me._VBComponent.Dispose()
        If Me._VBProject IsNot Nothing Then Me._VBProject.Dispose()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Quit")
    End Sub

    Protected Overrides Sub OnClosed()
        GC.Collect(2, GCCollectionMode.Forced, True)
    End Sub

End Class
