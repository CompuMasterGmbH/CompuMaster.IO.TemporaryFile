Option Explicit On
Option Strict On

Imports NUnit.Framework

Namespace CompuMaster.Tests.IO


    <TestFixture()> Public Class TemporaryFile

        <SetUp> Public Sub ResetTestEnvironment()
            CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Clear()
            CompuMaster.IO.TemporaryFile.IsUnitTestMode = True
        End Sub

        <Test()> Sub AddAndCleanupTempFile()
            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)

            Using NewTempFile1 As New CompuMaster.IO.TemporaryFile(".txt")
                NewTempFile1.CreateFile()
                Assert.AreEqual(True, NewTempFile1.FileExists)
                Assert.AreEqual(1, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
                Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            End Using

            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
        End Sub

        <Test()> Sub AddAndCleanupTempFiles()
            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)

            Dim NewTempFiles As List(Of String) = AddAndCleanupTempFilesHelper()

            'Dispose NewTempFile3 should lead to decreased temp files count ONLY after app exit, not on object dispose
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Assert.AreEqual(1, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFiles(2), CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            Assert.AreEqual(False, System.IO.File.Exists(NewTempFiles(0)))
            Assert.AreEqual(False, System.IO.File.Exists(NewTempFiles(1)))
            Assert.AreEqual(True, System.IO.File.Exists(NewTempFiles(2)))

            'Dispose NewTempFile3 should lead to decreased temp files count ONLY after app exit, now simulated
            CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit()
            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(False, System.IO.File.Exists(NewTempFiles(0)))
            Assert.AreEqual(False, System.IO.File.Exists(NewTempFiles(1)))
            Assert.AreEqual(False, System.IO.File.Exists(NewTempFiles(2)))
        End Sub

        Private Function AddAndCleanupTempFilesHelper() As List(Of String)
            Dim Result As New List(Of String)

            Dim NewTempFile1 As New CompuMaster.IO.TemporaryFile(".txt")
            Result.Add(NewTempFile1.FilePath)
            NewTempFile1.CreateFile()
            Assert.AreEqual(True, NewTempFile1.FileExists)
            Assert.AreEqual(1, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))

            Dim NewTempFile2 As New CompuMaster.IO.TemporaryFile(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.OnDispose, ".txt")
            Result.Add(NewTempFile2.FilePath)
            NewTempFile2.CreateFile()
            Assert.AreEqual(True, NewTempFile2.FileExists)
            Assert.AreEqual(2, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            Assert.AreEqual(NewTempFile2.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(1))

            Dim NewTempFile3 As New CompuMaster.IO.TemporaryFile(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.OnApplicationExit, ".txt")
            Result.Add(NewTempFile3.FilePath)
            NewTempFile3.CreateFile()
            Assert.AreEqual(True, NewTempFile3.FileExists)
            Assert.AreEqual(3, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            Assert.AreEqual(NewTempFile2.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(1))
            Assert.AreEqual(NewTempFile3.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(2))

            Return Result
        End Function

        <Test()> Sub CleanupOnApplicationExit()
            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)

            Dim NewTempFile1 As New CompuMaster.IO.TemporaryFile(".txt")
            NewTempFile1.CreateFile()
            Assert.AreEqual(True, NewTempFile1.FileExists)
            Assert.AreEqual(1, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))

            Dim NewTempFile2 As New CompuMaster.IO.TemporaryFile(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.OnDispose, ".txt")
            NewTempFile2.CreateFile()
            Assert.AreEqual(True, NewTempFile2.FileExists)
            Assert.AreEqual(2, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            Assert.AreEqual(NewTempFile2.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(1))

            Dim NewTempFile3 As New CompuMaster.IO.TemporaryFile(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.OnApplicationExit, ".txt")
            NewTempFile3.CreateFile()
            Assert.AreEqual(True, NewTempFile3.FileExists)
            Assert.AreEqual(3, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(NewTempFile1.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(0))
            Assert.AreEqual(NewTempFile2.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(1))
            Assert.AreEqual(NewTempFile3.FilePath, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit(2))

            'Dispose NewTempFile3 should lead to decreased temp files count ONLY after app exit, now simulated
            CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit()
            Assert.AreEqual(0, CompuMaster.IO.TemporaryFile.FilesToRemoveOnAppExit.Count)
            Assert.AreEqual(False, NewTempFile1.FileExists)
            Assert.AreEqual(False, NewTempFile2.FileExists)
            Assert.AreEqual(False, NewTempFile3.FileExists)
            Assert.AreEqual(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.None, NewTempFile1.CleanupTrigger)
            Assert.AreEqual(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.None, NewTempFile2.CleanupTrigger)
            Assert.AreEqual(CompuMaster.IO.TemporaryFile.TempFileCleanupEvent.None, NewTempFile3.CleanupTrigger)
        End Sub

        <Test>
        Sub PathTooLongExceptionOnCreateFile()
            Dim RootDir As String

            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
                    RootDir = "C:\"
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    RootDir = "/tmp/"
                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

            Dim LongPath As String = RootDir & New String("a"c, 300)
#If NETFRAMEWORK Then
            LongPath &="-netframework"
#Else
            LongPath &= "-net(core)"
#End If
            Dim TempFile As New CompuMaster.IO.TemporaryFile(LongPath, ".txt")
            Dim Ex As Exception = Nothing
            Try
                TempFile.CreateFile()
            Catch E As Exception
                Ex = E
            End Try

            Try
                Select Case System.Environment.OSVersion.Platform
                    Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                        ' Windows (NT-based and older versions)
#If NETFRAMEWORK Then
                        Assert.IsNotNull(Ex)
                        Assert.IsTrue(TypeOf Ex Is System.IO.PathTooLongException)
#Else
                        Assert.IsNotNull(Ex)
                        Assert.IsTrue(TypeOf Ex Is System.IO.IOException)
#End If
                    Case PlatformID.Unix
                        ' Unix-based systems (Linux, macOS, etc.)
                        Assert.IsNull(Ex)
                        Assert.IsTrue(TempFile.Exists)
                    Case Else
                        Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
                End Select
            Finally
                If TempFile IsNot Nothing Then TempFile.CleanUp()
            End Try
        End Sub

        <Test>
        Sub TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult()
            Dim RootDir As String

            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
                    RootDir = "C:\"
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    RootDir = "/tmp/"
                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

            Dim LongPath As String = RootDir & New String("a"c, 300)
#If NETFRAMEWORK Then
            LongPath &="-netframework"
#Else
            LongPath &= "-net(core)"
#End If
            Dim TempFile As New CompuMaster.IO.TemporaryFile(LongPath, ".txt")
            Dim TestResult As CompuMaster.IO.TemporaryFile.TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult = Nothing
            TestResult = TempFile.TestForWritablePathAndPathTooLongExceptionOnCurrentPlatform(False)

            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
#If NETFRAMEWORK Then
                    Assert.IsNotNull(TestResult)
                    Assert.IsNotNull(TestResult.FoundException)
                    Assert.IsTrue(TypeOf TestResult.FoundException Is System.IO.PathTooLongException)
#Else
                    Assert.IsNotNull(TestResult)
                    Assert.IsNotNull(TestResult.FoundException)
                    Assert.IsTrue(TypeOf TestResult.FoundException Is System.IO.IOException)
#End If
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    Assert.IsNotNull(TestResult)
                    Assert.IsFalse(TempFile.Exists) 'Should be false, because file was created + deleted in the test method
                    'ACCEPT ALL RESULTS AS IT IS - DON'T TEST FOR: Assert.IsNull(TestResult.FoundException)
                    'ACCEPT ALL RESULTS AS IT IS - DON'T TEST FOR: Assert.IsTrue(TestResult.FileWritable)

                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

            If TempFile IsNot Nothing Then TempFile.CleanUp()
        End Sub

        <Test>
        Sub PathTooLongExceptionOnCreateFileInfo()
            Dim RootDir As String

            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
                    RootDir = "C:\"
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    RootDir = "/tmp/"
                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

            Dim LongPath As String = RootDir & New String("a"c, 300)
#If NETFRAMEWORK Then
            LongPath &="-netframework"
#Else
            LongPath &= "-net(core)"
#End If
            Dim Ex As Exception = Nothing
            Dim TempFile As System.IO.FileInfo = Nothing
            Try
                TempFile = New System.IO.FileInfo(LongPath & ".txt")
            Catch E As Exception
                Ex = E
            End Try
            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
#If NETFRAMEWORK Then
                    Assert.IsNotNull(Ex)
                    Assert.IsTrue(TypeOf Ex Is System.IO.PathTooLongException)
#Else
                    Assert.IsFalse(TempFile.Exists)
#End If
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    Assert.IsFalse(TempFile.Exists)
                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

        End Sub

        <Test>
        Sub IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi()
            Dim LongPath As String
            Dim CharCount As Integer
            Dim RootDir As String

            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                    ' Windows (NT-based and older versions)
                    RootDir = "C:\"
                Case PlatformID.Unix
                    ' Unix-based systems (Linux, macOS, etc.)
                    RootDir = "/"
                Case Else
                    Throw New NotImplementedException("Platform not covered in unit test: " & System.Environment.OSVersion.Platform.ToString)
            End Select

            CharCount = 300
            LongPath = RootDir & New String("a"c, CharCount)
            Assert.IsTrue(CompuMaster.IO.TemporaryFile.IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(LongPath), "Test against filename in C:\ with " & CharCount & " chars")

            CharCount = 260
            LongPath = RootDir & New String("a"c, CharCount)
            Assert.IsTrue(CompuMaster.IO.TemporaryFile.IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(LongPath), "Test against filename in C:\ with " & CharCount & " chars")

            CharCount = 256
            LongPath = RootDir & New String("a"c, CharCount)
            Assert.IsTrue(CompuMaster.IO.TemporaryFile.IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(LongPath), "Test against filename in C:\ with " & CharCount & " chars")

            CharCount = 255
            LongPath = RootDir & New String("a"c, CharCount)
            Assert.IsFalse(CompuMaster.IO.TemporaryFile.IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(LongPath), "Test against filename in C:\ with " & CharCount & " chars")

            CharCount = 250
            LongPath = RootDir & New String("a"c, CharCount)
            Assert.IsFalse(CompuMaster.IO.TemporaryFile.IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(LongPath), "Test against filename in C:\ with " & CharCount & " chars")
        End Sub
    End Class

End Namespace