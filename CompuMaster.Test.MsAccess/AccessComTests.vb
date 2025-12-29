Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace CompuMaster.Test.MsAccess

    <NonParallelizable>
    Public Class AccessComTests

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            AppWithNordwind = OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
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
        Public Sub CodeData()
            ClassicAssert.IsNotNull(AppWithNordwind.CodeData)
            'Assert.IsNotNull(AppWithNordwind.CodeData.AllQueries)
            ClassicAssert.IsNotNull(AppWithNordwind.CodeData.Count)
            If False Then
                'Compilation test only
                AppWithNordwind.CodeData.Item(0).Parent.Parent.Quit()
                AppWithNordwind.CodeData.Parent.InvokeMethod("Quit")
            End If
        End Sub

        <Test>
        Public Sub CodeProject()
            ClassicAssert.IsNotNull(AppWithNordwind.CodeProject)
            'Console.WriteLine(AppWithNordwind.CodeProject.GetPublicMembersInfo)
            ClassicAssert.IsNotNull(AppWithNordwind.CodeProject.AllModules)
            ClassicAssert.NotZero(AppWithNordwind.CodeProject.AllModules.Count)
        End Sub

        <Test>
        Public Sub VBE()
            ClassicAssert.IsNotNull(AppWithNordwind.VBE)
            ClassicAssert.IsNotNull(AppWithNordwind.VBE.Count)
        End Sub

        <Test>
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