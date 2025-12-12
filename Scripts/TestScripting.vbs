' test to verify if windows scripting is working properly
Option Explicit

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim wsh: Set wsh = CreateObject("WScript.Shell")
Dim re: Set re = New RegExp
Dim dict: Set dict = CreateObject("Scripting.Dictionary")

Dim sBaseDir: sBaseDir = wsh.SpecialFolders("Desktop")
Dim fOut: Set fOut = fso.CreateTextFile(sBaseDir & "\TestScripting.log", True)

fOut.WriteLine "Windows Scripting Verification Test"
fOut.WriteLine "==================================="
fOut.WriteLine "Script Engine: " & GetScriptEngineInfo()

TestFSO
TestWSH
TestRE
TestDict

fOut.WriteLine "Tests complete."

MsgBox "Tests complete."
WScript.Quit 0

Sub TestFSO
    fOut.WriteLine "Testing FileSystemObject"
    fsoMain
    fOut.WriteLine "Finished testing FileSystemObject"
End Sub

Sub TestWSH
    fOut.WriteLine "Testing WScript.Shell"
    
    fOut.WriteLine "Finished testing WScript.Shell"    
End Sub

Sub TestRE
    fOut.WriteLine "Testing RegExp"
    
    re.Pattern = ".*bc.*"
    If re.Test("abcd") Then
        fOut.WriteLine "RegExp is working"
    End If
    fOut.WriteLine "Finished testing RegExp"
End Sub

Sub TestDict
    Dim sTmp
    fOut.WriteLine "Testing Scripting.Dictionary"
    dict.RemoveAll
    dict("abc") = "123"
    If dict.Exists("abc") Then
        sTmp = dict("abc")
        fOut.WriteLine "Dictionary access, should be 123: " & sTmp
    End If
    fOut.WriteLine "Finished testing Scripting.Dictionary"
End Sub

Function GetScriptEngineInfo
   Dim s
   s = ""   ' Build string with necessary info.
   s = ScriptEngine & " Version "
   s = s & ScriptEngineMajorVersion & "."
   s = s & ScriptEngineMinorVersion & "."
   s = s & ScriptEngineBuildVersion 
   GetScriptEngineInfo = s   ' Return the results.
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FileSystemObject Sample Code
' Copyright 1998 Microsoft Corporation.   All Rights Reserved. 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Regarding code quality:
' 1) The following code does a lot of string manipulation by 
'    concatenating short strings together with the "&" operator. 
'    Since string concatenation is expensive, this is a very 
'    inefficient way to write code. However, it is a very 
'    maintainable way to write code, and is used here because this 
'    program performs extensive disk operations, and because the 
'    disk is much slower than the memory operations required to 
'    concatenate the strings. Keep in mind that this is demonstration 
'    code, not production code.
'
' 2) "Option Explicit" is used, because declared variable access is 
'    slightly faster than undeclared variable access. It also prevents 
'    bugs from creeping into your code, such as when you misspell 
'    DriveTypeCDROM as DriveTypeCDORM.
'
' 3) Error handling is absent from this code, to make the code more 
'    readable. Although precautions have been taken to ensure that the 
'    code will not error in common cases, file systems can be 
'    unpredictable. In production code, use On Error Resume Next and 
'    the Err object to trap possible errors.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Some handy global variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TabStop
Dim NewLine

Const TestDrive = "C"
Const TestFilePath = "C:\ScriptingTest"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants returned by Drive.DriveType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const DriveTypeRemovable = 1
Const DriveTypeFixed = 2
Const DriveTypeNetwork = 3
Const DriveTypeCDROM = 4
Const DriveTypeRAMDisk = 5

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants returned by File.Attributes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const FileAttrNormal   = 0
Const FileAttrReadOnly = 1
Const FileAttrHidden = 2
Const FileAttrSystem = 4
Const FileAttrVolume = 8
Const FileAttrDirectory = 16
Const FileAttrArchive = 32 
Const FileAttrAlias = 1024
Const FileAttrCompressed = 2048

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants for opening files
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowDriveType
' Purpose: 
'    Generates a string describing the drive type of a given Drive object.
' Demonstrates the following 
'  - Drive.DriveType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ShowDriveType(Drive)

   Dim S
   
   Select Case Drive.DriveType
   Case DriveTypeRemovable
      S = "Removable"
   Case DriveTypeFixed
      S = "Fixed"
   Case DriveTypeNetwork
      S = "Network"
   Case DriveTypeCDROM
      S = "CD-ROM"
   Case DriveTypeRAMDisk
      S = "RAM Disk"
   Case Else
      S = "Unknown"
   End Select

   ShowDriveType = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowFileAttr
' Purpose: 
'    Generates a string describing the attributes of a file or folder.
' Demonstrates the following 
'  - File.Attributes
'  - Folder.Attributes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ShowFileAttr(File) ' File can be a file or folder

   Dim S
   Dim Attr
   
   Attr = File.Attributes

   If Attr = 0 Then
      ShowFileAttr = "Normal"
      Exit Function
   End If

   If Attr And FileAttrDirectory Then S = S & "Directory "
   If Attr And FileAttrReadOnly Then S = S & "Read-Only "
   If Attr And FileAttrHidden Then S = S & "Hidden "
   If Attr And FileAttrSystem Then S = S & "System "
   If Attr And FileAttrVolume Then S = S & "Volume "
   If Attr And FileAttrArchive Then S = S & "Archive "
   If Attr And FileAttrAlias Then S = S & "Alias "
   If Attr And FileAttrCompressed Then S = S & "Compressed "

   ShowFileAttr = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GenerateDriveInformation
' Purpose: 
'    Generates a string describing the current state of the 
'    available drives.
' Demonstrates the following 
'  - FileSystemObject.Drives 
'  - Iterating the Drives collection
'  - Drives.Count
'  - Drive.AvailableSpace
'  - Drive.DriveLetter
'  - Drive.DriveType
'  - Drive.FileSystem
'  - Drive.FreeSpace
'  - Drive.IsReady
'  - Drive.Path
'  - Drive.SerialNumber
'  - Drive.ShareName
'  - Drive.TotalSize
'  - Drive.VolumeName
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateDriveInformation(FSO)

   Dim Drives
   Dim Drive
   Dim S

   Set Drives = FSO.Drives
   S = "Number of drives:" & TabStop & Drives.Count & NewLine & NewLine

   ' Construct 1st line of report.
   S = S & String(2, TabStop) & "Drive" 
   S = S & String(3, TabStop) & "File" 
   S = S & TabStop & "Total"
   S = S & TabStop & "Free"
   S = S & TabStop & "Available" 
   S = S & TabStop & "Serial" & NewLine

   ' Construct 2nd line of report.
   S = S & "Letter"
   S = S & TabStop & "Path"
   S = S & TabStop & "Type"
   S = S & TabStop & "Ready?"
   S = S & TabStop & "Name"
   S = S & TabStop & "System"
   S = S & TabStop & "Space"
   S = S & TabStop & "Space"
   S = S & TabStop & "Space"
   S = S & TabStop & "Number" & NewLine   

   ' Separator line.
   S = S & String(105, "-") & NewLine

   For Each Drive In Drives
      S = S & Drive.DriveLetter
      S = S & TabStop & Drive.Path
      S = S & TabStop & ShowDriveType(Drive)
      S = S & TabStop & Drive.IsReady

      If Drive.IsReady Then
         If DriveTypeNetwork = Drive.DriveType Then
            S = S & TabStop & Drive.ShareName 
         Else
            S = S & TabStop & Drive.VolumeName 
         End If
         S = S & TabStop & Drive.FileSystem
         S = S & TabStop & Drive.TotalSize
         S = S & TabStop & Drive.FreeSpace
         S = S & TabStop & Drive.AvailableSpace
         S = S & TabStop & Hex(Drive.SerialNumber)
      End If

      S = S & NewLine

   Next

   GenerateDriveInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GenerateFileInformation
' Purpose: 
'    Generates a string describing the current state of a file.
' Demonstrates the following 
'  - File.Path
'  - File.Name
'  - File.Type
'  - File.DateCreated
'  - File.DateLastAccessed
'  - File.DateLastModified
'  - File.Size
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateFileInformation(File)

   Dim S

   S = NewLine & "Path:" & TabStop & File.Path
   S = S & NewLine & "Name:" & TabStop & File.Name
   S = S & NewLine & "Type:" & TabStop & File.Type
   S = S & NewLine & "Attribs:" & TabStop & ShowFileAttr(File)
   S = S & NewLine & "Created:" & TabStop & File.DateCreated
   S = S & NewLine & "Accessed:" & TabStop & File.DateLastAccessed
   S = S & NewLine & "Modified:" & TabStop & File.DateLastModified
   S = S & NewLine & "Size" & TabStop & File.Size & NewLine

   GenerateFileInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GenerateFolderInformation
' Purpose: 
'    Generates a string describing the current state of a folder.
' Demonstrates the following 
'  - Folder.Path
'  - Folder.Name
'  - Folder.DateCreated
'  - Folder.DateLastAccessed
'  - Folder.DateLastModified
'  - Folder.Size
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateFolderInformation(Folder)

   Dim S

   S = "Path:" & TabStop & Folder.Path
   S = S & NewLine & "Name:" & TabStop & Folder.Name
   S = S & NewLine & "Attribs:" & TabStop & ShowFileAttr(Folder)
   S = S & NewLine & "Created:" & TabStop & Folder.DateCreated
   S = S & NewLine & "Accessed:" & TabStop & Folder.DateLastAccessed
   S = S & NewLine & "Modified:" & TabStop & Folder.DateLastModified
   S = S & NewLine & "Size:" & TabStop & Folder.Size & NewLine

   GenerateFolderInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GenerateAllFolderInformation
' Purpose: 
'    Generates a string describing the current state of a
'    folder and all files and subfolders.
' Demonstrates the following 
'  - Folder.Path
'  - Folder.SubFolders
'  - Folders.Count
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateAllFolderInformation(Folder)

   Dim S
   Dim SubFolders
   Dim SubFolder
   Dim Files
   Dim File

   S = "Folder:" & TabStop & Folder.Path & NewLine & NewLine
   Set Files = Folder.Files

   If 1 = Files.Count Then
      S = S & "There is 1 file" & NewLine
   Else
      S = S & "There are " & Files.Count & " files" & NewLine
   End If

   If Files.Count <> 0 Then
      For Each File In Files
         S = S & GenerateFileInformation(File)
      Next
   End If

   Set SubFolders = Folder.SubFolders

   If 1 = SubFolders.Count Then
      S = S & NewLine & "There is 1 sub folder" & NewLine & NewLine
   Else
      S = S & NewLine & "There are " & SubFolders.Count & " sub folders" & NewLine & NewLine
   End If

   If SubFolders.Count <> 0 Then
      For Each SubFolder In SubFolders
         S = S & GenerateFolderInformation(SubFolder)
      Next
      S = S & NewLine
      For Each SubFolder In SubFolders
         S = S & GenerateAllFolderInformation(SubFolder)
      Next
   End If

   GenerateAllFolderInformation = S

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GenerateTestInformation
' Purpose: 
'    Generates a string describing the current state of the C:\Test
'    folder and all files and subfolders.
' Demonstrates the following 
'  - FileSystemObject.DriveExists
'  - FileSystemObject.FolderExists
'  - FileSystemObject.GetFolder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GenerateTestInformation(FSO)

   Dim TestFolder
   Dim S

   If Not FSO.DriveExists(TestDrive) Then Exit Function
   If Not FSO.FolderExists(TestFilePath) Then Exit Function

   Set TestFolder = FSO.GetFolder(TestFilePath)

   GenerateTestInformation = GenerateAllFolderInformation(TestFolder) 

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DeleteTestDirectory
' Purpose: 
'    Cleans up the test directory.
' Demonstrates the following 
'  - FileSystemObject.GetFolder
'  - FileSystemObject.DeleteFile
'  - FileSystemObject.DeleteFolder
'  - Folder.Delete
'  - File.Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub DeleteTestDirectory(FSO)

   Dim TestFolder
   Dim SubFolder
   Dim File
   ' Two ways to delete a file:

   FSO.DeleteFile(TestFilePath & "\Beatles\OctopusGarden.txt")

   Set File = FSO.GetFile(TestFilePath & "\Beatles\BathroomWindow.txt")
   File.Delete   

   ' Two ways to delete a folder:
   FSO.DeleteFolder(TestFilePath & "\Beatles")
   FSO.DeleteFile(TestFilePath & "\ReadMe.txt")
   Set TestFolder = FSO.GetFolder(TestFilePath)
   TestFolder.Delete

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateLyrics
' Purpose: 
'    Builds a couple of text files in a folder.
' Demonstrates the following 
'  - FileSystemObject.CreateTextFile
'  - TextStream.WriteLine
'  - TextStream.Write
'  - TextStream.WriteBlankLines
'  - TextStream.Close
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CreateLyrics(Folder)

   Dim TextStream
   Set TextStream = Folder.CreateTextFile("OctopusGarden.txt")
' Note that this does not add a line feed to the file.
   TextStream.Write("Octopus' Garden ") 
   TextStream.WriteLine("(by Ringo Starr)")
   TextStream.WriteBlankLines(1)
   TextStream.WriteLine("I'd like to be under the sea in an octopus' garden in the shade,")
   TextStream.WriteLine("He'd let us in, knows where we've been -- in his octopus' garden in the shade.")
   TextStream.WriteBlankLines(2)
   
   TextStream.Close

   Set TextStream = Folder.CreateTextFile("BathroomWindow.txt")
   TextStream.WriteLine("She Came In Through The Bathroom Window (by Lennon/McCartney)")
   TextStream.WriteLine("")
   TextStream.WriteLine("She came in through the bathroom window protected by a silver spoon")
   TextStream.WriteLine("But now she sucks her thumb and wanders by the banks of her own lagoon")
   TextStream.WriteBlankLines(2)
   TextStream.Close

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetLyrics
' Purpose: 
'    Displays the contents of the lyrics files.
' Demonstrates the following 
'  - FileSystemObject.OpenTextFile
'  - FileSystemObject.GetFile
'  - TextStream.ReadAll
'  - TextStream.Close
'  - File.OpenAsTextStream
'  - TextStream.AtEndOfStream
'  - TextStream.ReadLine
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetLyrics(FSO)

   Dim TextStream
   Dim S
   Dim File

   ' There are several ways to open a text file, and several 
   ' ways to read the data out of a file. Here's two ways 
   ' to do each:

   Set TextStream = FSO.OpenTextFile(TestFilePath & "\Beatles\OctopusGarden.txt", OpenFileForReading)
   S = TextStream.ReadAll & NewLine & NewLine
   TextStream.Close

   Set File = FSO.GetFile(TestFilePath & "\Beatles\BathroomWindow.txt")
   Set TextStream = File.OpenAsTextStream(OpenFileForReading)
   Do    While Not TextStream.AtEndOfStream
      S = S & TextStream.ReadLine & NewLine
   Loop
   TextStream.Close

   GetLyrics = S
   
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BuildTestDirectory
' Purpose: 
'    Builds a directory hierarchy to demonstrate the FileSystemObject.
'    We'll build a hierarchy in this order:
'       C:\Test
'       C:\Test\ReadMe.txt
'       C:\Test\Beatles
'       C:\Test\Beatles\OctopusGarden.txt
'       C:\Test\Beatles\BathroomWindow.txt
' Demonstrates the following 
'  - FileSystemObject.DriveExists
'  - FileSystemObject.FolderExists
'  - FileSystemObject.CreateFolder
'  - FileSystemObject.CreateTextFile
'  - Folders.Add
'  - Folder.CreateTextFile
'  - TextStream.WriteLine
'  - TextStream.Close
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function BuildTestDirectory(FSO)

   Dim TestFolder
   Dim SubFolders
   Dim SubFolder
   Dim TextStream

   ' Bail out if (a) the drive does not exist, or if (b) the directory is being built 
   ' already exists.

   If Not FSO.DriveExists(TestDrive) Then
        fOut.WriteLine "Test drive " & TestDrive & " not found."
      BuildTestDirectory = False
      Exit Function
   End If

   If FSO.FolderExists(TestFilePath) Then
        fOut.WriteLine "Test file path " & TestFilePath & " already exists."
      BuildTestDirectory = False
      Exit Function
   End If

   Set TestFolder = FSO.CreateFolder(TestFilePath)

   Set TextStream = FSO.CreateTextFile(TestFilePath & "\ReadMe.txt")
   TextStream.WriteLine("My song lyrics collection")
   TextStream.Close

   Set SubFolders = TestFolder.SubFolders
   Set SubFolder = SubFolders.Add("Beatles")
   CreateLyrics SubFolder   
   BuildTestDirectory = True

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The main routine
' First, it creates a test directory, along with some subfolders 
' and files. Then, it dumps some information about the available 
' disk drives and about the test directory, and then cleans 
' everything up again.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub fsoMain

   Dim myFSO

   ' Set up global data.
   TabStop = Chr(9)
   NewLine = Chr(10)
   Set myFSO = CreateObject("Scripting.FileSystemObject")

   If Not BuildTestDirectory(myFSO) Then 
      fOut.WriteLine "Test directory already exists or cannot be created.   Cannot continue."
      Exit Sub
   End If

   fOut.WriteLine GenerateDriveInformation(myFSO) & NewLine & NewLine
   fOut.WriteLine GenerateTestInformation(myFSO) & NewLine & NewLine
   fOut.WriteLine GetLyrics(myFSO) & NewLine & NewLine
   DeleteTestDirectory(myFSO)

End Sub
