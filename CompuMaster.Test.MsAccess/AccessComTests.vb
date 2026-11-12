Imports CompuMaster.ComInterop
Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace CompuMaster.Test.MsAccess

    <NonParallelizable, Apartment(Threading.ApartmentState.STA)>
    Public Class AccessComTests

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            If Not ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Access.Application") Then Assert.Ignore("Microsoft Access COM automation is unavailable.")
            AppWithNordwind = OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
        End Sub

        <OneTimeTearDown>
        Public Sub CloseAccess()
            AppWithNordwind?.Dispose()
        End Sub

        Private AppWithNordwind As Global.CompuMaster.MsAccess.AccessApplication

        Protected Function OpenAccessAppAndDatabase(databasePath As String) As Global.CompuMaster.MsAccess.AccessApplication
            Dim App As New Global.CompuMaster.MsAccess.AccessApplication
            Dim DbEngine = App.DBEngine
            Dim Db = DbEngine.OpenDatabase(databasePath)
            App.InvokeMethod("OpenCurrentDatabase", databasePath, False)
            Return App
        End Function

        <Test>
        Public Sub AccessBasicObjectAccess()
            ClassicAssert.IsNotNull(AppWithNordwind)
        End Sub

        <Test>
        Public Sub CurrentDb()
            ClassicAssert.IsNotNull(AppWithNordwind.CurrentDb)
        End Sub

        <Test>
        Public Sub CurrentProject()
            ClassicAssert.IsNotNull(AppWithNordwind.CurrentProject)
        End Sub

        <Test>
        Public Sub Modules()
            ClassicAssert.IsNotNull(AppWithNordwind.Modules)
            ClassicAssert.IsNotNull(AppWithNordwind.Modules.Count)
        End Sub

        <Test>
        Public Sub CurrentDataQueries()
            Dim data = AppWithNordwind.CreateWrapperAndRegisterComChildForDisposal(AppWithNordwind.InvokePropertyGet(Of Object)("CurrentData"), Nothing, Nothing, Nothing)
            Dim queries = data.CreateWrapperAndRegisterComChildForDisposal(data.InvokePropertyGet(Of Object)("AllQueries"), Nothing, Nothing, Nothing)
            Assert.That(queries.InvokePropertyGet(Of Integer)("Count"), [Is].GreaterThan(0))
        End Sub

        <Test>
        Public Sub CurrentProjectModules()
            Dim project = AppWithNordwind.CreateWrapperAndRegisterComChildForDisposal(AppWithNordwind.InvokePropertyGet(Of Object)("CurrentProject"), Nothing, Nothing, Nothing)
            Dim modules = project.CreateWrapperAndRegisterComChildForDisposal(project.InvokePropertyGet(Of Object)("AllModules"), Nothing, Nothing, Nothing)
            Assert.That(modules.InvokePropertyGet(Of Integer)("Count"), [Is].GreaterThan(0))
        End Sub

        <Test>
        Public Sub VBE()
            ClassicAssert.IsNotNull(AppWithNordwind.VBE)
            Assert.That(AppWithNordwind.VBE.InvokePropertyGet(Of String)("Version"), [Is].Not.Empty)
        End Sub

        <Test, Explicit("Requires trusted VBA in the test database and displays Microsoft Access.")>
        Public Sub Run()
            AppWithNordwind.Visible = True
            Console.WriteLine(AppWithNordwind.Run(Of Boolean)("HasSourceCode").ToString)
            Assert.Pass()
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