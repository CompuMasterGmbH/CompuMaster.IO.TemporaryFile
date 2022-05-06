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

    End Class

End Namespace