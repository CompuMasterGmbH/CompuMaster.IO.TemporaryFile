# CompuMaster.IO.TemporaryFile

When running applications, there is often a need to create temporary files and open them in your app or in a 3rd party app. In the end, the file(s) should be removed from disk on custom events or on application exit at latest (if possible, still existent and not opened/blocked in/by 3rd party app)

[![Github Release](https://img.shields.io/github/release/CompuMasterGmbH/CompuMaster.IO.TemporaryFile.svg?maxAge=2592000&label=GitHub%20Release)](https://github.com/CompuMasterGmbH/CompuMaster.IO.TemporaryFile/releases) 
[![NuGet CompuMaster.IO.TemporaryFile](https://img.shields.io/nuget/v/CompuMaster.IO.TemporaryFile.svg?label=NuGet%20CM.IO.TemporaryFile)](https://www.nuget.org/packages/CompuMaster.IO.TemporaryFile/) 

## Usage pattern: create temporary file for auto-cleanup

### C# Sample

```C#
using (var f = new CompuMaster.IO.TemporaryFile(".txt")) 
{
     System.IO.File.WriteAllBytes(f.FilePath, new Byte[] {});
     Console.WriteLine("some work");
}
// if file is not locked by another task: file is already removed
// if file is locked by another task: file will be removed on final call (see usage pattern below)
//    CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit();
```

### VB.Net Sample

```VB.Net
'
Using f As New CompuMaster.IO.TemporaryFile(".txt"))
     System.IO.File.WriteAllBytes(f.FilePath, New Byte() {})
     Console.WriteLine("some more work")
End Using

'if file is not locked by another task: file is already removed
'if file is locked by another task: file will be removed on final call (see usage pattern below)
'    CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit()
```

## Usage pattern: main method

### C# Sample

```C#
static void Main(string[] args)
{
    try
    {
        var f = New CompuMaster.IO.TemporaryFile(".txt")
        //do something more...
    }
    finally
    {
        CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit();
    }
}
```

### VB.Net Sample

```VB.Net
Sub Main(args as String())
    Try
        Dim f As New CompuMaster.IO.TemporaryFile(".txt")
        'do something more...
    Finally
        CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit()
    End Try
End Sub
```

## Usage pattern: main form

### C# Sample

```C#
void MainForm_FormClosing(Object sender, FormClosingEventArgs e)
{
    CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit();
}
```

### VB.Net Sample

```VB.Net
Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    CompuMaster.IO.TemporaryFile.CleanupOnApplicationExit()
End Sub
```
