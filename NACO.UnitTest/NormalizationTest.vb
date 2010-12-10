Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NACO.Data

'''<summary>
'''This is a test class for NormalizationTest and is intended
'''to contain all NormalizationTest Unit Tests
'''</summary>
<TestClass()> _
Public Class NormalizationTest

    Private Shared rootPath As String
    Private testContextInstance As TestContext

    '''<summary>
    '''Gets or sets the test context which provides
    '''information about and functionality for the current test run.
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "Additional test attributes"
    '
    'You can use the following additional attributes as you write your tests:
    '
    'Use ClassInitialize to run code before running the first test in the class
    <ClassInitialize()> _
    Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
        rootPath = "C:\Documents and Settings\AMG.USP_AD\My Documents\Visual Studio 2008\Projects\NACONormalization\NACONormalization\"

    End Sub
    '
    'Use ClassCleanup to run code after all tests in a class have run
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    'Use TestInitialize to run code before running each test
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    'Use TestCleanup to run code after each test has run
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    '''<summary>
    '''A test for normalize
    '''</summary>
    <TestMethod()> _
    Public Sub NormalizeSimplifiedTest()
        Dim ptr As String = String.Empty

        Dim ScriptFileName As String = RootPath + "NACO.script"
        Dim CheckFileName As String = RootPath + "NACOsimplified.check"
        Dim LogFileName As String = RootPath + "NACOsimplified.log"

        Dim script As IO.StreamReader = New IO.StreamReader(ScriptFileName, System.Text.Encoding.UTF8)
        Dim check As IO.StreamReader = New IO.StreamReader(CheckFileName, System.Text.Encoding.UTF8)
        Dim log As IO.StreamWriter = IO.File.CreateText(LogFileName)

        Dim chk As String = String.Empty
        Dim i As Integer = 0
        Dim norm As String = String.Empty
        Dim hasFailed As Boolean = False

        Try
            ptr = script.ReadLine
            Do While Not ptr Is Nothing
                i += 1
                chk = check.ReadLine
                norm = NACO.Data.Normalization.NormalizeSimplified(ptr)

                log.WriteLine("")
                log.WriteLine("Line " + i.ToString)
                If (Not chk.Equals(norm)) Then
                    log.WriteLine("     input: " + ptr)
                    log.WriteLine("  expected: " + chk)
                    log.WriteLine("       got: " + norm)
                    hasFailed = True
                Else
                    log.WriteLine("    SUCCESS")
                End If
                ptr = script.ReadLine
            Loop
        Catch ex As Exception
            log.WriteLine("")
            log.WriteLine(ex.StackTrace)
            log.WriteLine("record " + ptr)
            log.WriteLine(ptr)
            log.WriteLine("c=" + chk)
            log.WriteLine("norm=" + norm)
            hasFailed = True
        End Try

        script.Close()
        check.Close()
        log.Close()

        If hasFailed Then
            Assert.Fail("NormalizeSimplifiedTest failed. Check log " + LogFileName)
        End If
    End Sub


    '''<summary>
    '''A test for normalize
    '''</summary>
    <TestMethod()> _
    Public Sub NormalizeStandardTest()
        Dim ptr As String = String.Empty
        Dim leaveFirstComma As Boolean = True
        Dim musicalFlat As Char = "F"
        Dim output_Subfield_Separator As Char = Normalization.DEFAULT_SUBFIELD_SEPARATOR

        Dim ScriptFileName As String = rootPath + "NACO.script"
        Dim CheckFileName As String = rootPath + "NACOstandard.check"
        Dim LogFileName As String = rootPath + "NACOstandard.log"

        Dim script As IO.StreamReader = New IO.StreamReader(ScriptFileName, System.Text.Encoding.UTF8)
        Dim check As IO.StreamReader = New IO.StreamReader(CheckFileName, System.Text.Encoding.UTF8)
        Dim log As IO.StreamWriter = IO.File.CreateText(LogFileName)

        Dim chk As String = String.Empty
        Dim i As Integer = 0
        Dim isSimplified As Boolean = False
        Dim norm As String = String.Empty
        Dim hasFailed As Boolean = False

        Try
            ptr = script.ReadLine
            Do While Not ptr Is Nothing
                i += 1
                chk = check.ReadLine
                norm = NACO.Data.Normalization.NormalizeStandard(ptr, leaveFirstComma, musicalFlat, output_Subfield_Separator)

                log.WriteLine("")
                log.WriteLine("Line " + i.ToString)
                If (Not chk.Equals(norm)) Then
                    log.WriteLine("     input: " + ptr)
                    log.WriteLine("  expected: " + chk)
                    log.WriteLine("       got: " + norm)
                    hasFailed = True
                Else
                    log.WriteLine("    SUCCESS")
                End If
                ptr = script.ReadLine
            Loop
        Catch ex As Exception
            log.WriteLine("")
            log.WriteLine(ex.StackTrace)
            log.WriteLine("record " + ptr)
            log.WriteLine(ptr)
            log.WriteLine("c=" + chk)
            log.WriteLine("norm=" + norm)
            hasFailed = True
        End Try

        script.Close()
        check.Close()
        log.Close()

        If hasFailed Then
            Assert.Fail("NormalizeStandardTest failed. Check log " + LogFileName)
        End If
    End Sub
End Class
