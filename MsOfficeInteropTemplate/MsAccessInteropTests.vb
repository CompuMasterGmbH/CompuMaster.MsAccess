Imports CompuMaster.ComInterop
Imports Microsoft.Office.Interop.Access.Dao
Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace CompuMaster.Test.MsAccess

    <NonParallelizable>
    Public Class MsAccessInteropTests

        Public Sub CompileTests_MsAccessInteropApp()

            Dim App As Microsoft.Office.Interop.Access.Application = Nothing
            Dim DbEngine As Microsoft.Office.Interop.Access.Dao._DBEngine = App.DBEngine
            Dim Db = DbEngine.OpenDatabase("")

            App.CurrentDb()
            'ComObject.VBE
            'ComObject.DBEngine
            'ComObject.CodeData.AllFunctions.Item
            'ComObject.CodeProject.AllForms
            Dim dummy As Integer = App.Modules.Count

        End Sub

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
        End Sub

        Private _AppWithNordwind As ComRootObject(Of Microsoft.Office.Interop.Access.Application)
        Private ReadOnly Property AppWithNordwind As Microsoft.Office.Interop.Access.Application
            Get
                Return _AppWithNordwind.ComObjectStronglyTyped
            End Get
        End Property

        Private _DbEngine As ComChildObject(Of Microsoft.Office.Interop.Access.Dao._DBEngine, ComRootObject(Of Microsoft.Office.Interop.Access.Application))
        Private ReadOnly Property DbEngine As Microsoft.Office.Interop.Access.Dao._DBEngine
            Get
                Return _DbEngine.ComObjectStronglyTyped
            End Get
        End Property

        Private _NordwindDb As ComChildObject(Of Microsoft.Office.Interop.Access.Dao.Database, ComRootObject(Of Microsoft.Office.Interop.Access.Application))
        Private ReadOnly Property NordwindDb As Microsoft.Office.Interop.Access.Dao.Database
            Get
                Return _NordwindDb.ComObjectStronglyTyped
            End Get
        End Property

        Protected Sub OpenAccessAppAndDatabase(databasePath As String)
            _AppWithNordwind = New ComRootObject(Of Microsoft.Office.Interop.Access.Application)(
                New Microsoft.Office.Interop.Access.Application(),
                Nothing,
                Sub(instance)
                    instance.ComObjectStronglyTyped.Quit()
                End Sub,
                Nothing)
            AppWithNordwind.OpenCurrentDatabase(databasePath, False)
            '_NordwindDb = AppWithNordwind.CurrentProject
            _DbEngine = New ComChildObject(Of Microsoft.Office.Interop.Access.Dao._DBEngine, ComRootObject(Of Microsoft.Office.Interop.Access.Application))(_AppWithNordwind, AppWithNordwind.DBEngine)
            _NordwindDb = New ComChildObject(Of Microsoft.Office.Interop.Access.Dao.Database, ComRootObject(Of Microsoft.Office.Interop.Access.Application))(_AppWithNordwind, DbEngine.OpenDatabase(databasePath))
            ClassicAssert.IsNotNull(_NordwindDb.ComObjectStronglyTyped.Name)
        End Sub

        <Test>
        Public Sub AccessBasicObjectAccess()
            ClassicAssert.IsNotNull(AppWithNordwind)
        End Sub

        <Test>
        Public Sub CurrentDb()
            ClassicAssert.IsNotNull(_NordwindDb.ComObjectStronglyTyped)
            ClassicAssert.IsNull(AppWithNordwind.CurrentDb)
        End Sub

        <Test>
        Public Sub CurrentProject()
            ClassicAssert.IsNotNull(_NordwindDb.ComObjectStronglyTyped)
            ClassicAssert.IsNotNull(AppWithNordwind.CurrentProject)
        End Sub

        <Test>
        Public Sub Run()
            Console.WriteLine(_AppWithNordwind.ComObjectStronglyTyped.Visible.ToString)
            Console.WriteLine(_AppWithNordwind.InvokeFunction(Of Boolean)("Run", "HasSourceCode").ToString)
            Console.WriteLine(_AppWithNordwind.ComObjectStronglyTyped.Run("HasSourceCode").ToString)
            _AppWithNordwind.ComObjectStronglyTyped.Run("DoKbTestWithParameter", "Hello from VB .NET Client")
        End Sub

        <Test>
        Public Sub Modules()
            ClassicAssert.IsNotNull(AppWithNordwind.Modules)
            For Each Item As Microsoft.Office.Interop.Access.Module In AppWithNordwind.Modules
                ClassicAssert.IsNotNull(Item)
                Console.WriteLine("Found Module: " & Item.Name)
            Next
            ClassicAssert.IsNotNull(AppWithNordwind.Modules.Count)
        End Sub

        <Test>
        Public Sub CodeData()
            ClassicAssert.IsNotNull(AppWithNordwind.CodeData)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllTables.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllQueries.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllViews.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllStoredProcedures.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllFunctions.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeData.AllDatabaseDiagrams.Count)
        End Sub

        <Test>
        Public Sub CodeProject()
            ClassicAssert.IsNotNull(AppWithNordwind.CodeProject)
            Dim M = AppWithNordwind.CodeProject.AllModules
            ClassicAssert.NotZero(M.Count)
            Dim M0 = M(0)
            ClassicAssert.IsNotNull(M0.Name)

            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllModules.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllMacros.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllForms.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllReports.Count)
            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllDataAccessPages.Count)
        End Sub

        <Test>
        Public Sub VBE()
            ClassicAssert.IsNotNull(AppWithNordwind.VBE)
            ClassicAssert.AreEqual(1, AppWithNordwind.VBE.VBProjects.Count)
            ClassicAssert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0))
            ClassicAssert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).Name)
            ClassicAssert.NotZero(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Count)
            ClassicAssert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Item(0))
            ClassicAssert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Item(0).Type)
        End Sub

        'Public Sub CodeData()
        '    Dim App As Global.CompuMaster.MsAccess.AccessApplication = OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
        '    ClassicAssert.IsNotNull(App.Modules)
        '    ClassicAssert.IsNotNull(App.Modules.Count)
        '    ComObject.VBE
        '    ComObject.DBEngine
        '    ComObject.CodeData.AllFunctions.Item
        '    ComObject.CodeProject.AllForms
        '    ComObject.Modules
        'End Sub

    End Class

End Namespace
