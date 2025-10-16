Option Explicit On
Option Strict On

Namespace CompuMaster.IO

    ''' <summary>
    ''' Represents a temporary file which shall be deleted automatically on application end
    ''' </summary>
    ''' <remarks>
    ''' Requires implementation of following pattern in MainForm / on application exit
    ''' <code language="C#">
    ''' private void MainForm_FormClosing(Object sender, FormClosingEventArgs e)
    ''' {
    '''     TemporaryFile.TryToRemoveAllDisposedButLockedTempFiles();
    ''' }
    ''' </code>
    ''' <code language="vb">
    ''' Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '''     TemporaryFile.TryToRemoveAllDisposedButLockedTempFiles()
    ''' End Sub
    ''' </code>
    ''' </remarks>
    Public Class TemporaryFile
        Implements IDisposable

        ''' <summary>
        ''' Create a new instance of a temp file with auto-cleanup-feature with auto-removal on dispose
        ''' </summary>
        ''' <param name="extension">A file extension like .docx</param>
        ''' <remarks></remarks>
        Public Sub New(extension As String)
            Me.New(TempFileCleanupEvent.OnDispose, Nothing, Nothing, extension)
        End Sub

        ''' <summary>
        ''' Create a temporary file in a random sub directory of user's temp directory) with auto-removal on dispose
        ''' </summary>
        ''' <param name="fileNameWithoutExtension"></param>
        ''' <param name="extension"></param>
        Public Sub New(fileNameWithoutExtension As String, extension As String)
            Me.New(TempFileCleanupEvent.OnDispose, Nothing, fileNameWithoutExtension, extension)
        End Sub

        ''' <summary>
        ''' Create a new instance of a temp file with auto-cleanup-feature
        ''' </summary>
        ''' <param name="extension">A file extension like .docx</param>
        ''' <remarks></remarks>
        Public Sub New(cleanupTrigger As TempFileCleanupEvent, extension As String)
            Me.New(cleanupTrigger, Nothing, Nothing, extension)
        End Sub

        ''' <summary>
        ''' Create a temporary file in a random sub directory of user's temp directory)
        ''' </summary>
        ''' <param name="cleanupTrigger"></param>
        ''' <param name="fileNameWithoutExtension"></param>
        ''' <param name="extension"></param>
        Public Sub New(cleanupTrigger As TempFileCleanupEvent, fileNameWithoutExtension As String, extension As String)
            Me.New(cleanupTrigger, Nothing, fileNameWithoutExtension, extension)
        End Sub

        ''' <summary>
        ''' Create a temporary file in a custom directory (relative directory names are relative to the user's temp directory) with auto-removal on dispose
        ''' </summary>
        ''' <param name="subDirectory"></param>
        ''' <param name="fileNameWithoutExtension"></param>
        ''' <param name="extension"></param>
        Public Sub New(subDirectory As String, fileNameWithoutExtension As String, extension As String)
            Me.New(TempFileCleanupEvent.OnDispose, subDirectory, fileNameWithoutExtension, extension)
        End Sub

        ''' <summary>
        ''' Create a temporary file in a custom directory (relative directory names are relative to the user's temp directory)
        ''' </summary>
        ''' <param name="subDirectory"></param>
        ''' <param name="fileNameWithoutExtension"></param>
        ''' <param name="extension"></param>
        Public Sub New(cleanupTrigger As TempFileCleanupEvent, subDirectory As String, fileNameWithoutExtension As String, extension As String)
            If fileNameWithoutExtension = Nothing Then fileNameWithoutExtension = System.IO.Path.GetRandomFileName
            If extension = Nothing Then Throw New ArgumentNullException(NameOf(extension))
            If extension.StartsWith(".") = False Then Throw New ArgumentException("extension must start with a dot character")
            Dim TempDir As String
            If subDirectory <> Nothing Then
                TempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), subDirectory)
            Else
                TempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath())
            End If
            If System.IO.Directory.Exists(TempDir) = False Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If
            Me.FilePath = System.IO.Path.Combine(TempDir, fileNameWithoutExtension & extension)
            Me._CleanupTrigger = cleanupTrigger
            If IsUnitTestMode Then WriteToLog("Reserved temp file location " & Me.FilePath)
        End Sub

        ''' <summary>
        ''' Preferred cleanup trigger
        ''' </summary>
        Public Enum TempFileCleanupEvent As Byte
            ''' <summary>
            ''' Remove the file on application exit
            ''' </summary>
            OnApplicationExit = 0
            ''' <summary>
            ''' Remove the file when this instance is getting disposed or (in case of locked files) on application exit
            ''' </summary>
            OnDispose = 1
            ''' <summary>
            ''' No automatic file removal
            ''' </summary>
            None = 255
        End Enum

        ''' <summary>
        ''' An absolute file path to the temporary file
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property FilePath As String

        ''' <summary>
        ''' The applying cleanup trigger
        ''' </summary>
        ''' <returns></returns>
        Public Overridable ReadOnly Property CleanupTrigger As TempFileCleanupEvent
            Get
                Return _CleanupTrigger
            End Get
        End Property

        ''' <summary>
        ''' The configured cleanup trigger
        ''' </summary>
        Private __CleanupTrigger As TempFileCleanupEvent

        ''' <summary>
        ''' The applying cleanup trigger based on file existance
        ''' </summary>
        ''' <returns></returns>
        Private Property _CleanupTrigger As TempFileCleanupEvent
            Get
                If Me.FileExists Then
                    Return __CleanupTrigger
                Else
                    Return TempFileCleanupEvent.None
                End If
            End Get
            Set(value As TempFileCleanupEvent)
                Select Case value
                    Case TempFileCleanupEvent.OnDispose
                        'register FilePath for 2nd-chance-removal on application exit
                        FilesToRemoveOnAppExit.Add(Me.FilePath)
                    Case TempFileCleanupEvent.OnApplicationExit
                        'register FilePath for removal on application exit
                        FilesToRemoveOnAppExit.Add(Me.FilePath)
                    Case TempFileCleanupEvent.None
                        'do nothing
                        RemoveTempFileFromAppExitCleanUpList(Me.FilePath)
                    Case Else
                        Throw New ArgumentOutOfRangeException(NameOf(CleanupTrigger))
                End Select
                __CleanupTrigger = value
            End Set
        End Property

        ''' <summary>
        ''' Create a 0-sized file
        ''' </summary>
        Friend Sub CreateFile()
            System.IO.File.WriteAllBytes(Me.FilePath, New Byte() {})
            If IsUnitTestMode Then WriteToLog("Created 0-sized " & Me.FilePath)
        End Sub

        ''' <summary>
        ''' Determines whether the specified file exists
        ''' </summary>
        ''' <returns></returns>
        Friend Function FileExists() As Boolean
            Return System.IO.File.Exists(Me.FilePath)
        End Function

        ''' <summary>
        ''' Remove the temporary file from disk if it exists
        ''' </summary>
        ''' <remarks>For explicit method calls. If this method isn't called, there will be a cleanup on dispose. Exceptions will be thrown when trying to delete blocked files, read-only files, etc.</remarks>
        Public Overridable Sub CleanUp()
            If System.IO.File.Exists(Me.FilePath) Then
                System.IO.File.Delete(Me.FilePath)
                If IsUnitTestMode Then WriteToLog("Manually deleted " & Me.FilePath)
            End If
            Me._CleanupTrigger = TempFileCleanupEvent.None
        End Sub

        ''' <summary>
        ''' Remove the temporary file from disk if it exists
        ''' </summary>
        ''' <remarks>For explicit method calls. If this method isn't called, there will be a cleanup on dispose. Exceptions will be ignored when trying to delete blocked files, read-only files, etc.</remarks>
        Public Overridable Sub TryCleanUp()
            Try
                Me.CleanUp()
            Catch
                'ignore all exceptions
            End Try
            Me._CleanupTrigger = TempFileCleanupEvent.None
        End Sub

        Protected Overrides Sub Finalize()
            If IsUnitTestMode Then WriteToLog("Finalize " & Me.FilePath)
            Me.Dispose(True)
            MyBase.Finalize()
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If IsUnitTestMode Then WriteToLog("Disposing " & Me.FilePath)
            If Not Me.disposedValue Then
                If disposing AndAlso (IsApplicationClosing OrElse Me.CleanupTrigger = TempFileCleanupEvent.OnDispose) Then
                    Try
                        If System.IO.File.Exists(Me.FilePath) Then
                            System.IO.File.Delete(Me.FilePath)
                            If IsUnitTestMode Then WriteToLog("Dispose deleted " & Me.FilePath)
                        End If
                        Me._CleanupTrigger = TempFileCleanupEvent.None
                    Catch ex As System.IO.IOException
                        'ignore all errors like "Der Prozess kann nicht auf die Datei "C:\Users\wezel\AppData\Local\Temp\1\20121120200830980378654776317.docx" zugreifen, da sie von einem anderen Prozess verwendet wird." e.g. when closing HLS while document is still open
                        If Not IsApplicationClosing Then
                            Me._CleanupTrigger = TempFileCleanupEvent.OnApplicationExit
                        End If
                    Catch ex As Exception
                        'ignore all errors regarding temporary file cleanup
                        If Not IsApplicationClosing Then
                            Me._CleanupTrigger = TempFileCleanupEvent.OnApplicationExit
                        End If
                    End Try
                End If
            End If
            Me.disposedValue = True
        End Sub

        ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

#Region "AppContext"
        Friend Shared ReadOnly FilesToRemoveOnAppExit As New List(Of String)
        Private Shared IsApplicationClosing As Boolean = False

        ''' <summary>
        ''' After successful deletion of a file, the cleanup list item for app exit should be removed, too
        ''' </summary>
        ''' <param name="filePath"></param>
        Private Sub RemoveTempFileFromAppExitCleanUpList(filePath As String)
            If FilesToRemoveOnAppExit.Contains(filePath) Then
                FilesToRemoveOnAppExit.Remove(filePath)
            End If
        End Sub

        ''' <summary>
        ''' Run final cleanup on application exit and try to remove all remaining temporary files from disk
        ''' </summary>
        ''' <remarks>All exceptions (e.g. from file system) are catched and ignored</remarks>
        Public Shared Sub CleanupOnApplicationExit()
            TryToRemoveAllDisposedButLockedTempFiles()
        End Sub

        ''' <summary>
        ''' Try to remove all remaining temporary files from disk 
        ''' </summary>
        Private Shared Sub TryToRemoveAllDisposedButLockedTempFiles()
            IsApplicationClosing = True
            For MyCounter As Integer = FilesToRemoveOnAppExit.Count - 1 To 0 Step -1
                Dim f As String = FilesToRemoveOnAppExit(MyCounter)
                Try
                    If System.IO.File.Exists(f) Then
                        System.IO.File.Delete(f)
                        If IsUnitTestMode Then WriteToLog("CleanupOnApplicationExit deleted " & f)
                    End If
                    FilesToRemoveOnAppExit.RemoveAt(MyCounter)
                Catch
                    'ignore all errors regarding temporary file cleanup
                End Try
            Next
            IsApplicationClosing = False
        End Sub

        Friend Shared IsUnitTestMode As Boolean

        Private Shared Sub WriteToLog(data As String)
            System.Console.WriteLine(data)
        End Sub
#End Region

        ''' <summary>
        ''' Get file information of the temporary file
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Throws an exception if the file does not exist</remarks>
        Public Function GetFileInfo() As System.IO.FileInfo
            Return New System.IO.FileInfo(Me.FilePath)
        End Function

        ''' <summary>
        ''' Does the file exist?
        ''' </summary>
        ''' <returns></returns>
        Public Function Exists() As Boolean
            Return System.IO.File.Exists(Me.FilePath)
        End Function

        Public Structure TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult
            ''' <summary>
            ''' The exception that was thrown when trying to access the given path, e.g. PathTooLongException, UnauthorizedAccessException, SecurityException, etc.
            ''' </summary>
            ''' <returns></returns>
            Public Property FoundException As Exception
            ''' <summary>
            ''' The tested path is not too long for the current platform and a file could be created or opened for writing
            ''' </summary>
            ''' <returns></returns>
            Public Property FileWritable As Boolean
        End Structure

        ''' <summary>
        ''' Test if the given path would cause a PathTooLongException on the current platform, in fact opening or creating a file or directory at this path
        ''' </summary>
        ''' <param name="allowTestToCreateDirectoryPathIfRequired">If the file doesn't exist and the directory path doesn't exist, the test might need the parent directory to create the test file. True to allow it (might lead to a left-over, empty parent directory, False to throw a System.IO.DirectoryNotFoundException</param>
        ''' <returns>True if a PathTooLongException would be thrown on the current platform when trying to access the given path</returns>
        ''' <remarks>
        ''' Relative paths are not supported for this test.
        ''' 
        ''' Windows classic API limits (without Long Paths enabled):
        ''' - MAX_PATH = 260 characters (including terminator, so effectively 259 characters)
        ''' - MAX_COMPONENT = 255 characters (typical NTFS limit per segment)
        ''' 
        ''' Note that Windows Subsystem for Linux (WSL) and some other file systems might have different limits.
        ''' 
        ''' Note on non-Windows platforms:
        ''' This test is based on Windows/NTFS limits, too, because the .NET runtime might be running on a non-Windows platform (e.g. Linux, macOS) but the file system might be a network share or a mounted volume which is actually hosted on a Windows/NTFS system.
        ''' 
        ''' Note on Long Paths:
        ''' Since Windows 10 v1607 and Windows Server 2016, Long Paths can be enabled by a manifest setting and/or a group policy setting.
        ''' If Long Paths are enabled, the MAX_PATH limit of 260 characters does not apply anymore for applications which are manifested accordingly.
        ''' However, the MAX_COMPONENT limit of 255 characters per segment still applies even if Long Paths are enabled.
        ''' 
        ''' Note on Unicode normalization:
        ''' This test does not cover Unicode normalization issues which might lead to different path lengths depending on the normalization form used by different file systems.
        ''' For example, a character like "é" can be represented as a single code point (U+00E9) or as a combination of "e" (U+0065) and an acute accent (U+0301).
        ''' Different file systems might treat these representations differently, potentially leading to unexpected behavior when accessing files with such characters in their names.
        ''' </remarks>
        Public Function TestForWritablePathAndPathTooLongExceptionOnCurrentPlatform(allowTestToCreateDirectoryPathIfRequired As Boolean) As TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult
            Return TestForWritablePathAndPathTooLongExceptionOnCurrentPlatform(Me.FilePath, allowTestToCreateDirectoryPathIfRequired)
        End Function

        ''' <summary>
        ''' Test if the given path would cause a PathTooLongException on the current platform, in fact opening or creating+deleting a file at this path
        ''' </summary>
        ''' <param name="path">An absolute path</param>
        ''' <param name="allowTestToCreateDirectoryPathIfRequired">If the file doesn't exist and the directory path doesn't exist, the test might need the parent directory to create the test file. True to allow it (might lead to a left-over, empty parent directory, False to throw a System.IO.DirectoryNotFoundException</param>
        ''' <returns>True if a PathTooLongException would be thrown on the current platform when trying to access the given path</returns>
        ''' <remarks>
        ''' Relative paths are not supported for this test.
        ''' 
        ''' Windows classic API limits (without Long Paths enabled):
        ''' - MAX_PATH = 260 characters (including terminator, so effectively 259 characters)
        ''' - MAX_COMPONENT = 255 characters (typical NTFS limit per segment)
        ''' 
        ''' Note that Windows Subsystem for Linux (WSL) and some other file systems might have different limits.
        ''' 
        ''' Note on non-Windows platforms:
        ''' This test is based on Windows/NTFS limits, too, because the .NET runtime might be running on a non-Windows platform (e.g. Linux, macOS) but the file system might be a network share or a mounted volume which is actually hosted on a Windows/NTFS system.
        ''' 
        ''' Note on Long Paths:
        ''' Since Windows 10 v1607 and Windows Server 2016, Long Paths can be enabled by a manifest setting and/or a group policy setting.
        ''' If Long Paths are enabled, the MAX_PATH limit of 260 characters does not apply anymore for applications which are manifested accordingly.
        ''' However, the MAX_COMPONENT limit of 255 characters per segment still applies even if Long Paths are enabled.
        ''' 
        ''' Note on Unicode normalization:
        ''' This test does not cover Unicode normalization issues which might lead to different path lengths depending on the normalization form used by different file systems.
        ''' For example, a character like "é" can be represented as a single code point (U+00E9) or as a combination of "e" (U+0065) and an acute accent (U+0301).
        ''' Different file systems might treat these representations differently, potentially leading to unexpected behavior when accessing files with such characters in their names.
        ''' </remarks>
        Public Shared Function TestForWritablePathAndPathTooLongExceptionOnCurrentPlatform(path As String, allowTestToCreateDirectoryPathIfRequired As Boolean) As TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult
            If System.IO.Path.IsPathRooted(path) = False Then
                'Relative paths are not supported for this test
                Throw New ArgumentException("path must be an absolute path")
            End If
            Dim Result As New TestForWritablePathAndPathTooLongExceptionOnCurrentPlatformResult
            Try
                If System.IO.File.Exists(path) Then
                    'File exists, try to open it
                    Using f As System.IO.FileStream = System.IO.File.Open(path, System.IO.FileMode.Open, System.IO.FileAccess.Write, System.IO.FileShare.Read)
                        'File could be opened -> no PathTooLongException
                        Result.FileWritable = True
                        f.Close()
                    End Using
                ElseIf System.IO.Directory.Exists(path) Then
                    'Directory exists, try to get its info
                    Dim di As New System.IO.DirectoryInfo(path)
                    Dim TestFilePath = System.IO.Path.Combine(path, "~" & Guid.NewGuid.ToString("n") & ".tmp")
                    'File/dir does not exist, try to create new file and delete it afterwards
                    'NOTE: This test might fail due to other reasons like missing (write) permissions
                    Using f As System.IO.FileStream = System.IO.File.Create(path)
                        'File could be created -> no PathTooLongException
                        Result.FileWritable = True
                    End Using
                    'Delete the created file again
                    System.IO.File.Delete(path)
                Else
                    'Consider path as new file
                    Dim fi As New System.IO.FileInfo(path) 'should not throw an exception
                    Dim di As New System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(path)) 'should not throw an exception
                    If di.Exists = False Then
                        'Parent directory does not exist
                        If allowTestToCreateDirectoryPathIfRequired Then
                            'try to create it
                            di.Create()
                        Else
                            Throw New System.IO.DirectoryNotFoundException("The directory path for the given file does not exist: " & di.FullName)
                        End If
                    End If

                    'File/dir does not exist, try to create new file and delete it afterwards
                    'NOTE: This test might fail due to other reasons like missing (write) permissions
                    Using f As System.IO.FileStream = System.IO.File.Create(path)
                        'File could be created -> no PathTooLongException
                        f.Close()
                        Result.FileWritable = True
                    End Using
                    'Delete the created file again
                    System.IO.File.Delete(path)
                End If
            Catch ex As Exception
                Result.FoundException = ex
                Result.FileWritable = False
            End Try
            Return Result
        End Function

        ''' <summary>
        ''' Test based on path length limits with rules for classic Windows API and NTFS file system
        ''' </summary>
        ''' <returns>True if the path has got issues which might conflict in Windows/NTFS environments, False if the path's length is safe to use</returns>
        ''' <remarks>
        ''' Relative paths are not supported for this test.
        ''' 
        ''' Windows classic API limits (without Long Paths enabled):
        ''' - MAX_PATH = 260 characters (including terminator, so effectively 259 characters)
        ''' - MAX_COMPONENT = 255 characters (typical NTFS limit per segment)
        ''' 
        ''' Note that Windows Subsystem for Linux (WSL) and some other file systems might have different limits.
        ''' 
        ''' Note on non-Windows platforms:
        ''' This test is based on Windows/NTFS limits, too, because the .NET runtime might be running on a non-Windows platform (e.g. Linux, macOS) but the file system might be a network share or a mounted volume which is actually hosted on a Windows/NTFS system.
        ''' 
        ''' Note on Long Paths:
        ''' Since Windows 10 v1607 and Windows Server 2016, Long Paths can be enabled by a manifest setting and/or a group policy setting.
        ''' If Long Paths are enabled, the MAX_PATH limit of 260 characters does not apply anymore for applications which are manifested accordingly.
        ''' However, the MAX_COMPONENT limit of 255 characters per segment still applies even if Long Paths are enabled.
        ''' 
        ''' Note on Unicode normalization:
        ''' This test does not cover Unicode normalization issues which might lead to different path lengths depending on the normalization form used by different file systems.
        ''' For example, a character like "é" can be represented as a single code point (U+00E9) or as a combination of "e" (U+0065) and an acute accent (U+0301).
        ''' Different file systems might treat these representations differently, potentially leading to unexpected behavior when accessing files with such characters in their names.
        ''' </remarks>
        Public Function IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi() As Boolean
            Return IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(Me.FilePath)
        End Function

        ''' <summary>
        ''' Test based on path length limits with rules for classic Windows API and NTFS file system
        ''' </summary>
        ''' <param name="fullPath"></param>
        ''' <returns>True if the path has got issues which might conflict in Windows/NTFS environments, False if the path's length is safe to use</returns>
        ''' <remarks>
        ''' Relative paths are not supported for this test.
        ''' 
        ''' Windows classic API limits (without Long Paths enabled):
        ''' - MAX_PATH = 260 characters (including terminator, so effectively 259 characters)
        ''' - MAX_COMPONENT = 255 characters (typical NTFS limit per segment)
        ''' 
        ''' Note that Windows Subsystem for Linux (WSL) and some other file systems might have different limits.
        ''' 
        ''' Note on non-Windows platforms:
        ''' This test is based on Windows/NTFS limits, too, because the .NET runtime might be running on a non-Windows platform (e.g. Linux, macOS) but the file system might be a network share or a mounted volume which is actually hosted on a Windows/NTFS system.
        ''' 
        ''' Note on Long Paths:
        ''' Since Windows 10 v1607 and Windows Server 2016, Long Paths can be enabled by a manifest setting and/or a group policy setting.
        ''' If Long Paths are enabled, the MAX_PATH limit of 260 characters does not apply anymore for applications which are manifested accordingly.
        ''' However, the MAX_COMPONENT limit of 255 characters per segment still applies even if Long Paths are enabled.
        ''' 
        ''' Note on Unicode normalization:
        ''' This test does not cover Unicode normalization issues which might lead to different path lengths depending on the normalization form used by different file systems.
        ''' For example, a character like "é" can be represented as a single code point (U+00E9) or as a combination of "e" (U+0065) and an acute accent (U+0301).
        ''' Different file systems might treat these representations differently, potentially leading to unexpected behavior when accessing files with such characters in their names.
        ''' </remarks>
        Public Shared Function IsPathOrPathComponentTooLongForClassicWinApiAndNtfsApi(fullPath As String) As Boolean
            ' Klassische Windows-Grenzen (ohne aktivierte Long Paths):
            Const MAX_PATH As Integer = 260        ' inkl. Terminator -> effektiv 259 Zeichen
            Const MAX_COMPONENT As Integer = 255   ' typische NTFS-Grenze je Segment

            Dim componentTooLong As Boolean = System.IO.Path.GetFileName(fullPath).Length > MAX_COMPONENT
            For Each segment As String In fullPath.Split(New Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries)
                If segment.Length > MAX_COMPONENT Then
                    componentTooLong = True
                    Exit For
                End If
            Next
            Dim pathTooLong As Boolean = fullPath.Length >= (MAX_PATH)

            Return componentTooLong OrElse pathTooLong
        End Function

    End Class

End Namespace
