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
            If IsUnitTestMode Then System.Console.WriteLine("Reserved temp file location " & Me.FilePath)
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
            If IsUnitTestMode Then System.Console.WriteLine("Created 0-sized " & Me.FilePath)
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
                If IsUnitTestMode Then System.Console.WriteLine("Manually deleted " & Me.FilePath)
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
            If IsUnitTestMode Then System.Console.WriteLine("Finalize " & Me.FilePath)
            Me.Dispose(True)
            MyBase.Finalize()
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If IsUnitTestMode Then System.Console.WriteLine("Disposing " & Me.FilePath)
            If Not Me.disposedValue Then
                If disposing AndAlso (IsApplicationClosing OrElse Me.CleanupTrigger = TempFileCleanupEvent.OnDispose) Then
                    Try
                        If System.IO.File.Exists(Me.FilePath) Then
                            System.IO.File.Delete(Me.FilePath)
                            If IsUnitTestMode Then System.Console.WriteLine("Dispose deleted " & Me.FilePath)
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
                        If IsUnitTestMode Then System.Console.WriteLine("CleanupOnApplicationExit deleted " & f)
                    End If
                    FilesToRemoveOnAppExit.RemoveAt(MyCounter)
                Catch
                    'ignore all errors regarding temporary file cleanup
                End Try
            Next
            IsApplicationClosing = False
        End Sub

        Friend Shared IsUnitTestMode As Boolean
#End Region

    End Class

End Namespace
