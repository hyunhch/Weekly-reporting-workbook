Attribute VB_Name = "FindPath"
'Attribute VB_Name = "GetLocalOneDrivePath"
' Cross-platform VBA Function to get the local path of OneDrive/SharePoint
' synchronized Microsoft Office files (Works on Windows and on macOS)
'
' Author: Guido Witt-Dörring
' Created: 2022/07/01
' Updated: 2023/03/27
' License: MIT
'
' ----------------------------------------------------------------
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ----------------------------------------------------------------
'
' Copyright (c) 2022 Guido Witt-Dörring
'
' MIT License:
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.

'*******************************************************************************
' COMMENTS REGARDING THE IMPLEMENTATION:
' 1)    Adding the MountPoints from the registry is not the best solution for
'       many reasons.
'       Most importantly, the registry is just not available on Mac.
'       Also, in various cases these registry keys can contain mistakes, like
'       for example when:
'       1. Synchronizing a folder called "Personal" from someone else's personal
'          OneDrive
'       2. Synchronizing a folder called "Business1" from someone else's
'          personal OneDrive and then relogging your own first Business OneDrive
'       3. Relogging you own personal OneDrive can change the "CID" property
'          from a folderID formatted cid (e.g. 3DEA8A9886F05935!125) to a
'          regular private cid (e.g. 3dea8a9886f05935) for synced folders
'          from other people's OneDrives
'
'       Also, even if we assume the information from the registry is correct,
'       it just doesn't contain enough information to solve all possible cases.
'       This is why this function builds the Web to Local translation dictionary
'       by extracting the MountPoints from the OneDrive settings files.
'
' 2)    This function reads the following files from
'       On Windows:
'           - the "...\AppData\Local\Microsoft" directory
'       On Mac:
'           - the ".../Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
'                 "Library/Application Support/OneDrive" directory
'
'       (* is a wildcard representing 0 or more characters, ? is a
'       wildcard representing a character, # is a wildcard representing a digit
'       The ???... filenames represent cids):
'       \OneDrive\settings\Personal\ClientPolicy.ini
'       \OneDrive\settings\Personal\????????????????.dat
'       \OneDrive\settings\Personal\????????????????.ini
'       \OneDrive\settings\Personal\global.ini
'       \OneDrive\settings\Personal\GroupFolders.ini
'       \OneDrive\settings\Business#\????????-????-????-????-????????????.dat
'       \OneDrive\settings\Business#\????????-????-????-????-????????????.ini
'       \OneDrive\settings\Business#\ClientPolicy*.ini
'       \OneDrive\settings\Business#\global.ini
'       \Office\CLP\* (just the filename)
'
'       On mac, the \Office\CLP\* exists for each Microsoft Office application
'       separately. Depending on whether the application was already used in
'       active syncing with OneDrive it may contain different/incomplete files.
'       In the code, the path of this directory is stored inside the variable
'       "clpPath". On Mac, the defined clpPath might not exist or not contain
'       all necessary files for some host applications, because Environ("HOME")
'       depends on the host app.
'       This is not a big problem as the function will still work, however in
'       this case, specifying a preferredMountPointOwner will do nothing.
'       To make sure this directory and the necessary files exist, a file must
'       have been actively synchronized with OneDrive by the application whose
'       "HOME" folder is returned by Environ("HOME") while being logged in
'       to that application with the account whose email is given as
'       preferredMountPointOwner, at some point in the past!
'
'       If you are usually working with Excel but are using this function in a
'       different app, you can instead use an alternative (Excels CLP folder) as
'       the clpPath as it will most likely contain all the necessary information
'       The alternative clpPath is commented out in the code, if you prefer to
'       use Excels CLP folder per default, just un-comment the respective line
'       in the code.
'*******************************************************************************

'*******************************************************************************
' COMMENTS REGARDING THE USAGE:
' This function can be used as a User Defined Function (UDF) from the worksheet.
'
' This function offers three optional parameters to the user, however using
' these should only be necessary in extremely rare situations. The best rule
' regarding their usage: Don't use them.
'
' In the following these parameters will still be explained.
'
'1) rebuildCache
'   The function creates a "translation" dictionary from the OneDrive settings
'   files and then uses this dictionary to "translate" WebPaths to LocalPaths.
'   This dictionary is implemented as a static variable to the function doesn't
'   have to recreate it every time it is called. It is written on the first
'   function call and reused on all the subsequent calls, making them faster.
'   If the function is called with rebuildCache:=True, this dictionary will be
'   rewritten, even if it was already initialized.
'   Note that it is not necessary to use this parameter manually, even if a new
'   MountPoint was added to the OneDrive, or a new OneDrive account was logged
'   in since the last function call because the function will automatically
'   determine if any of those cases occurred, without sacrificing performance.
'
'2) returnAll
'   In some exceptional cases it is possible to map one OneDrive WebPath to
'   multiple different localPaths. This can happen when multiple Business
'   OneDrive accounts are logged in on the device, and multiple of these have
'   access to the same OneDrive folder and they both decide to synchronize it or
'   add it as link to their MySite library.
'   Calling the function with returnAll:=True will return all valid localPaths
'   for the given WebPath, separated by two forward slashes (//). This should be
'   used with caution, as the return value of the function alone is, should
'   multiple local paths exist for the input webPath, not a valid local path
'   anymore.
'   An example of how to obtain all of the local paths could look like this:
'   Dim localPath as String, localPaths() as String
'   localPath = GetLocalPath(webPath, False, True)
'   If Not localPath Like "http*" Then
'       localPaths = Split(localPath, "//")
'   End If
'
'3) preferredMountPointOwner
'   This argument deals with the same problem as 'returnAll'
'   If the function gets called with returnAll:=False (default), and multiple
'   localPaths exist for the given WebPath, the function will just return any
'   one of them, as usually, it shouldn't make a difference, because the result
'   directories at both of these localPaths are mirrored versions of the same
'   webPath. Nevertheless, this option lets the user choose, which mountPoint
'   should be chosen if multiple localPaths are available. Each localPath is
'  'owned' by an OneDrive Account. If a WebPath is synchronized twice, this can
'   only happen by synchronizing it with two different accounts, because
'   OneDrive prevents you from synchronizing the same folder twice on a single
'   account. Therefore, each of the different localPaths for a given WebPath
'   has a unique 'owner'. preferredMountPointOwner lets the user select the
'   localPath by specifying the account the localPath should be owned by.
'   This is done by passing the Email address of the desired account as
'   preferredMountPointOwner.
'   For example, you have two different Business OneDrive accounts logged in,
'   foo.bar@business1.com and foo.bar@business2.com
'   Both synchronize the WebPath:
'   webPath = "https://business1.sharepoint.com/sites/TestLib/Documents/" & _
              "Test/Test/Test/test.xlsm"
'
'   The first one has added it as a link to his personal OneDrive, the local
'   path looks like this:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   The second one just synchronized it normally, the localPath looks like this:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   Calling GetLocalPath like this:
'   GetLocalPath(webPath,,, "foo.bar@business1.com") will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   Calling it like this:
'   GetLocalPath(webPath,,, "foo.bar@business2.com") will return:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   And calling it like this:
'   GetLocalPath(webPath,, True) will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm//C:\Users\username\Business1\TestLinkLib - Test\Test\...
'   ...test.xlsm
'
'   Calling it normally like this:
'   GetLocalPath(webPath) will return any one of the two localPaths, so:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'   OR
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'*******************************************************************************
Option Explicit

''*******************************************************************************
'' USAGE EXAMPLES:
'' Excel:
'Private Sub TestGetLocalPathExcel()
'    Debug.Print GetLocalPath(ThisWorkbook.FullName)
'    Debug.Print GetLocalPath(ThisWorkbook.path)
'End Sub
'
'' Usage as User Defined Function (UDF):
'' NOTE: You might have to replace ; with , in the formulas depending on settings
'' Add this formula to any cell, to get the local path of the workbook:
'' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1))
''
'' To get the local path including the filename (the FullName), use this formula:
'' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1) &
'' TEXTAFTER(TEXTBEFORE(CELL("filename";A1);"]");"["))
'
''Word:
'Private Sub TestGetLocalPathWord()
'    Debug.Print GetLocalPath(ThisDocument.FullName)
'    Debug.Print GetLocalPath(ThisDocument.path)
'End Sub
'
''PowerPoint:
'Private Sub TestGetLocalPathPowerPoint()
'    Debug.Print GetLocalPath(ActivePresentation.FullName)
'    Debug.Print GetLocalPath(ActivePresentation.path)
'End Sub


'*******************************************************************************

'This Function will convert a OneDrive/SharePoint Url path, e.g. Url containing
'https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
'to the locally synchronized path on your current pc or mac, e.g. a path like
'C:\users\username\OneDrive\ on Windows; or /Users/username/OneDrive/ on MacOS,
'if you have the remote directory locally synchronized with the OneDrive app.
'If no local path can be found, the input value will be returned unmodified.
'Author: Guido Witt-Dörring
'Source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'        https://stackoverflow.com/a/73577057/12287457
Public Function GetLocalPath(ByVal path As String, _
                    Optional ByVal returnAll As Boolean = False, _
                    Optional ByVal preferredMountPointOwner As String = "", _
                    Optional ByVal rebuildCache As Boolean = False) _
                             As String
    #If Mac Then
        Const vbErrPermissionDenied As Long = 70
        Const vbErrInvalidFormatInResourceFile As Long = 325
        Const noErrJustDecodeUTF8 As Long = vbObjectError + 62468
        Const isMac As Boolean = True
        Const syncIDFileName As String = ".849C9593-D756-4E56-8D6E-42412F2A707B"
        Const ps As String = "/" 'Application.PathSeparator doesn't work
    #Else 'Windows               'in all host applications (e.g. Outlook), hence
        Const ps As String = "\" 'conditional compilation is preferred here.
        Const isMac As Boolean = False
    #End If
    Const vbErrFileNotFound As Long = 53
    Const vbErrOutOfMemory As Long = 7
    Const vbErrKeyAlreadyExists As Long = 457
    Const chunkOverlap As Long = 1000
    Static locToWebColl As Collection, lastTimeNotFound As Collection
    Static lastCacheUpdate As Date
    Dim resColl As Object, webRoot As String, locRoot As String
    Dim vItem As Variant, s As String, keyExists As Boolean
    Dim pmpo As String: pmpo = LCase(preferredMountPointOwner)

    If Not locToWebColl Is Nothing And Not rebuildCache Then
        Set resColl = New Collection: GetLocalPath = ""
        'If the locToWebColl is initialized, this logic will find the local path
        For Each vItem In locToWebColl
            locRoot = vItem(0): webRoot = vItem(1)
            If InStr(1, path, webRoot, vbTextCompare) = 1 Then _
                resColl.Add Key:=vItem(2), _
                   Item:=Replace(Replace(path, webRoot, locRoot, , 1), "/", ps)
        Next vItem
        If resColl.Count > 0 Then
            If returnAll Then
                For Each vItem In resColl: s = s & "//" & vItem: Next vItem
                GetLocalPath = Mid(s, 3): Exit Function
            End If
            On Error Resume Next: GetLocalPath = resColl(pmpo): On Error GoTo 0
            If GetLocalPath <> "" Then Exit Function
            GetLocalPath = resColl(1): Exit Function
        End If
        'Local path was not found with cached mountpoints
        GetLocalPath = path 'No Exit Function here! Check if cache needs rebuild
    End If

    'Declare all variables that will be used in the loop over OneDrive settings
    Dim cid As String, fileNum As Long, line As Variant, parts() As String
    Dim tag As String, mainMount As String, relPath As String, email As String
    Dim b() As Byte, n As Long, i As Long, size As Long, libNr As String
    Dim parentID As String, folderID As String, folderName As String
    Dim folderIdPattern As String, FileName As String, folderType As String
    Dim siteID As String, libID As String, webID As String, lnkID As String
    Dim syncID As String, mainSyncID As String
    Dim syncFind As String, mainSyncFind As String
    Dim odFolders As Object, cliPolColl As Object, libNrToWebColl As Object
    Dim sig1 As String: sig1 = MidB$(Chr$(&H2), 1, 1)
    Dim sig2 As String: sig2 = ChrW$(&H1) & String(3, vbNullChar) 'x01 (x00 * 7)
    Dim vbNullByte As String: vbNullByte = MidB$(vbNullChar, 1, 1) 'x00
    Dim buffSize As Long, lastChunkEndPos As Long, lenDatFile As Long
    Dim lastFileUpdate As Date, coll As Collection, duplicateCheck As Collection
    Dim dirName As Variant, wDir As Variant, settDirIsDuplicate As Boolean
    #If Mac Then 'Variables for manual decoding of UTF-8, UTF-32 and ANSI
        Dim j As Long, k As Long, m As Long, ansi() As Byte, sAnsi As String
        Dim utf16() As Byte, sUtf16 As String, utf32() As Byte
        Dim utf8() As Byte, sUtf8 As String, numBytesOfCodePoint As Long
        Dim codepoint As Long, lowSurrogate As Long, highSurrogate As Long
        ReDim b(0 To 3): b(0) = &HAB&: b(1) = &HAB&: b(2) = &HAB&: b(3) = &HAB&
        Dim sig3 As String: sig3 = b: sig3 = vbNullChar & vbNullChar & sig3
    #Else 'Windows
        ReDim b(0 To 1): b(0) = &HAB&: b(1) = &HAB&
        Dim sig3 As String: sig3 = b: sig3 = vbNullChar & sig3
    #End If

    Dim settPaths As Collection: Set settPaths = New Collection
    Dim settPath As Variant, clpPath As String
    #If Mac Then 'The settings directories can be in different locations
        Dim cloudStoragePath As String, cloudStoragePathExists As Boolean
        s = Environ("HOME")
        clpPath = s & "/Library/Application Support/Microsoft/Office/CLP/"
        s = Left(s, InStrRev(s, "/Library/Containers/"))
        settPaths.Add s & _
                      "Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
                      "Library/Application Support/OneDrive/settings/"
        settPaths.Add s & "Library/Application Support/OneDrive/settings/"
        cloudStoragePath = s & "Library/CloudStorage/"

        'Excels CLP folder:
        'clpPath = left(s, InStrRev(s, "/Library/Containers")) & _
                  "Library/Containers/com.microsoft.Excel/Data/" & _
                  "Library/Application Support/Microsoft/Office/CLP/"
    #Else 'On Windows, the settings directories are always in this location:
        settPaths.Add Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
        clpPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
    #End If

    #If Mac Then 'Request access to all possible directories at once
        Dim possibleDirs As Variant
        If locToWebColl Is Nothing Then  '(only necessary on the first call)
            Set coll = New Collection
            For Each settPath In settPaths
                coll.Add Item:=settPath
                For i = 1 To 9: coll.Add settPath & "Business" & i & ps: Next i
                coll.Add Item:=settPath & "Personal" & ps
            Next settPath
            coll.Add Item:=clpPath
            coll.Add Item:=cloudStoragePath
            If coll.Count > 0 Then
                ReDim possibleDirs(1 To coll.Count)
                For i = 1 To coll.Count: possibleDirs(i) = coll(i): Next i
                If Not GrantAccessToMultipleFiles(possibleDirs) Then _
                    Err.Raise vbErrPermissionDenied
            End If
        End If
    #End If

    'Find all subdirectories in OneDrive settings folder:
    Dim oneDriveSettDirs As Collection: Set oneDriveSettDirs = New Collection
    For Each settPath In settPaths
        dirName = Dir(settPath, vbDirectory)
        Do Until dirName = ""
            If dirName = "Personal" Or dirName Like "Business#" Then _
                oneDriveSettDirs.Add Item:=settPath & dirName & ps
            dirName = Dir(, vbDirectory)
        Loop
    Next settPath

    
    If Not locToWebColl Is Nothing Or isMac Then
        Dim requiredFiles As Collection: Set requiredFiles = New Collection
        'Get collection of all required files
        For Each wDir In oneDriveSettDirs
           cid = IIf(wDir Like "*" & ps & "Personal" & ps, "????????????????", _
                     "????????-????-????-????-????????????")
            FileName = Dir(wDir, vbNormal)
            Do Until FileName = ""
                If FileName Like cid & ".ini" _
                Or FileName Like cid & ".dat" _
                Or FileName Like "ClientPolicy*.ini" _
                Or StrComp(FileName, "GroupFolders.ini", vbTextCompare) = 0 _
                Or StrComp(FileName, "global.ini", vbTextCompare) = 0 Then _
                    requiredFiles.Add Item:=wDir & FileName
                FileName = Dir
            Loop
        Next wDir
    End If

    'This part should ensure perfect accuracy despite the mount point cache
    'while sacrificing almost no performance at all by querying FileDateTimes.
    If Not locToWebColl Is Nothing And Not rebuildCache Then
        'Check if a settings file was modified since the last cache rebuild
        Dim vFile As Variant
        For Each vFile In requiredFiles
            If FileDateTime(vFile) > lastCacheUpdate Then _
                rebuildCache = True: Exit For 'full cache refresh is required!
        Next vFile
        If Not rebuildCache Then Exit Function
    End If

    'If execution reaches this point, the cache will be fully rebuilt...
    lastCacheUpdate = Now()

    #If Mac Then 'Prepare building syncIDtoSyncDir dictionary. This involves
        'reading the ".849C9593-D756-4E56-8D6E-42412F2A707B" files inside the
        'subdirs of "~/Library/CloudStorage/", list of files and access required
        Dim vDir As Variant
        Set coll = New Collection
        dirName = Dir(cloudStoragePath, vbDirectory)
        Do Until dirName = ""
            If dirName Like "OneDrive*" Then
                cloudStoragePathExists = True
                vDir = cloudStoragePath & dirName & ps
                vFile = cloudStoragePath & dirName & ps & syncIDFileName
                coll.Add Item:=vDir
                coll.Add Item:=vFile, Key:=vDir 'Key for targeted removal later
                requiredFiles.Add Item:=vDir 'For pooling file access requests
                requiredFiles.Add Item:=vFile
            End If
            dirName = Dir(, vbDirectory)
        Loop
        
        'Pool access request for these files and the OneDrive/settings files
        If locToWebColl Is Nothing Then
            Dim vFiles As Variant
            If requiredFiles.Count > 0 Then
                ReDim vFiles(1 To requiredFiles.Count)
               For i = 1 To UBound(vFiles): vFiles(i) = requiredFiles(i): Next i
                If Not GrantAccessToMultipleFiles(vFiles) Then _
                    Err.Raise vbErrPermissionDenied
            End If
        End If
        
        'More access might be required if some folders inside cloudStoragePath
        'don't contain the hidden file ".849C9593-D756-4E56-8D6E-42412F2A707B".
        'In that case, access to their first level subfolders is also required.
        If cloudStoragePathExists Then
            'Remove all files from coll (not the folders!): Remember:
            On Error Resume Next 'coll(coll(i)) = coll(i) & syncIDFileName
            For i = coll.Count To 1 Step -1: coll.Remove coll(i): Next i
            On Error GoTo 0
            For i = coll.Count To 1 Step -1
                If Dir(coll(i) & syncIDFileName, vbHidden) = "" Then
                    dirName = Dir(coll(i), vbDirectory)
                    Do Until dirName = ""
                        If Not dirName Like ".Trash*" And dirName <> "Icon" Then
                            coll.Add coll(i) & dirName & ps
                            coll.Add coll(i) & dirName & ps & syncIDFileName, _
                                     coll(i) & dirName & ps  '<- key for removal
                        End If
                        dirName = Dir(, vbDirectory)
                    Loop          'Remove the
                    coll.Remove i 'folder if it doesn't contain the hidden file.
                End If
            Next i
            If coll.Count > 0 Then
                ReDim possibleDirs(1 To coll.Count)
                For i = 1 To coll.Count: possibleDirs(i) = coll(i): Next i
                If Not GrantAccessToMultipleFiles(possibleDirs) Then _
                    Err.Raise vbErrPermissionDenied
            End If
            'Remove all files from coll (not the folders!): Reminder:
            On Error Resume Next 'coll(coll(i)) = coll(i) & syncIDFileName
            For i = coll.Count To 1 Step -1: coll.Remove coll(i): Next i
            On Error GoTo 0

            'Write syncIDtoSyncDir collection
            Dim syncIDtoSyncDir As Collection
            Set syncIDtoSyncDir = New Collection
            For Each wDir In coll
                If Dir(wDir & syncIDFileName, vbHidden) <> "" Then
                    fileNum = FreeFile(): s = "": vFile = wDir & syncIDFileName
                    'Somehow reading these files with "Open" doesn't always work
                    Dim readSucceeded As Boolean: readSucceeded = False
                    On Error GoTo ReadFailed
                    Open vFile For Binary Access Read As #fileNum
                        ReDim b(0 To LOF(fileNum)): Get fileNum, , b: s = b
                        readSucceeded = True
ReadFailed:             On Error GoTo -1
                    Close #fileNum: fileNum = 0
                    On Error GoTo 0
                    If readSucceeded Then
                        'Debug.Print "Used open statement to read file: " & _
                                    wDir & syncIDFileName
                        ansi = s 'If Open was used: Decode ANSI string manually:
                        If LenB(s) > 0 Then
                            ReDim utf16(0 To LenB(s) * 2 - 1): k = 0
                            For j = LBound(ansi) To UBound(ansi)
                                utf16(k) = ansi(j): k = k + 2
                            Next j
                            s = utf16
                        Else: s = ""
                        End If
                    Else 'Reading the file with "Open" failed with an error. Try
                        'using AppleScript. Also avoids the manual transcoding.
                        'Somehow ApplScript fails too, sometimes. Seems whenever
                        '"Open" works, AppleScript fails and vice versa (?!?!)
                        vFile = MacScript("return path to startup disk as " & _
                                    "string") & Replace(Mid(vFile, 2), ps, ":")
                        s = MacScript("return read file """ & _
                                      vFile & """ as string")
                       'Debug.Print "Used Apple Script to read file: " & vFile
                    End If
                    s = Split(s, """guid"" : """)(1)
                    syncID = Left(s, InStr(1, s, """") - 1)
                    syncIDtoSyncDir.Add Key:=syncID, _
                              Item:=VBA.Array(syncID, Left(wDir, Len(wDir) - 1))
                End If
            Next wDir
        End If
    #End If

    'Writing locToWebColl using .ini and .dat files in the OneDrive settings:
    'Here, a Scripting.Dictionary would be nice but it is not available on Mac!
    Set locToWebColl = New Collection
    Set duplicateCheck = New Collection
    For Each wDir In oneDriveSettDirs 'One folder per logged in OD account
        dirName = Mid(wDir, InStrRev(wDir, ps, Len(wDir) - 1) + 1)
        dirName = Left(dirName, Len(dirName) - 1)
        On Error Resume Next 'Only if duplicate settings directories exist, we
        duplicateCheck dirName: settDirIsDuplicate = Err.Number <> 0  'will
        On Error GoTo 0 'allow duplicate locRoots (helps with debugging)
        
        'Read global.ini to get cid
        If Dir(wDir & "global.ini", vbNormal) = "" Then GoTo NextFolder
        fileNum = FreeFile()
        Open wDir & "global.ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: On Error GoTo DecodeUTF8: Err.Raise noErrJustDecodeUTF8
            On Error GoTo 0: b = sUtf16 'b = StrConv(b, vbUnicode) <- UNRELIABLE
        #End If
        For Each line In Split(b, vbNewLine)
            If line Like "cid = *" Then cid = Mid(line, 7): Exit For
        Next line

        If cid = "" Then GoTo NextFolder
        If (Dir(wDir & cid & ".ini") = "" Or _
            Dir(wDir & cid & ".dat") = "") Then GoTo NextFolder
        If dirName Like "Business#" Then
            folderIdPattern = Replace(Space(32), " ", "[a-f0-9]")
        ElseIf dirName = "Personal" Then
            folderIdPattern = Replace(Space(16), " ", "[A-F0-9]") & "!###*"
        End If

        'Get email for business accounts
        '(only necessary to let user choose preferredMountPointOwner)
        FileName = Dir(clpPath, vbNormal)
        Do Until FileName = ""
            i = InStrRev(FileName, cid, , vbTextCompare)
            If i > 1 And cid <> "" Then _
                email = LCase(Left(FileName, i - 2)): Exit Do
            FileName = Dir
        Loop

        'Read all the ClientPloicy*.ini files:
        Set cliPolColl = New Collection
        FileName = Dir(wDir, vbNormal)
        Do Until FileName = ""
            If FileName Like "ClientPolicy*.ini" Then
                fileNum = FreeFile()
                Open wDir & FileName For Binary Access Read As #fileNum
                    ReDim b(0 To LOF(fileNum)): Get fileNum, , b
                Close #fileNum: fileNum = 0
                #If Mac Then 'On Mac, OneDrive settings files use UTF-8 encoding
                    sUtf8 = b: On Error GoTo DecodeUTF8
                    Err.Raise noErrJustDecodeUTF8 'This is not an error!
                    On Error GoTo 0: b = sUtf16 'StrConv(b, vbUnicode)UNRELIABLE
                #End If
                cliPolColl.Add Key:=FileName, Item:=New Collection
                For Each line In Split(b, vbNewLine)
                    If InStr(1, line, " = ", vbBinaryCompare) Then
                        tag = Left(line, InStr(line, " = ") - 1)
                        s = Mid(line, InStr(line, " = ") + 3)
                        Select Case tag
                        Case "DavUrlNamespace"
                            cliPolColl(FileName).Add Key:=tag, Item:=s
                        Case "SiteID", "IrmLibraryId", "WebID" 'Only used for
                            s = Replace(LCase(s), "-", "")  'backup method later
                            If Len(s) > 3 Then s = Mid(s, 2, Len(s) - 2)
                            cliPolColl(FileName).Add Key:=tag, Item:=s
                        End Select
                    End If
                Next line
            End If
            FileName = Dir
        Loop

        'Read cid.dat file
        buffSize = -1 'Buffer uninitialized
Try:    On Error GoTo Catch
        Set odFolders = New Collection
        lastChunkEndPos = 1: i = 0 'i = current reading pos.
        lastFileUpdate = FileDateTime(wDir & cid & ".dat")
        Do
            'Ensure file is not changed while reading it
            If FileDateTime(wDir & cid & ".dat") > lastFileUpdate Then GoTo Try
            fileNum = FreeFile
            Open wDir & cid & ".dat" For Binary Access Read As #fileNum
                lenDatFile = LOF(fileNum)
                If buffSize = -1 Then buffSize = lenDatFile 'Initialize buffer
                'Overallocate a bit so read chunks overlap to recognize all dirs
                ReDim b(0 To buffSize + chunkOverlap)
                Get fileNum, lastChunkEndPos, b: s = b: size = LenB(s)
            Close #fileNum: fileNum = 0
            lastChunkEndPos = lastChunkEndPos + buffSize

            For vItem = 16 To 8 Step -8
                i = InStrB(vItem + 1, s, sig2) 'Sarch byte pattern in cid.dat
                Do While i > vItem And i < size - 168 'and confirm with another
                    If MidB$(s, i - vItem, 1) = sig1 Then 'pattern at offset
                        i = i + 8: n = InStrB(i, s, vbNullByte) - i 'i:Start pos
                        If n < 0 Then n = 0                         'n: Length
                        If n > 39 Then n = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            folderID = MidB$(s, i, n)
                            ansi = folderID 'Decode ANSI string manually:
                            If LenB(folderID) > 0 Then
                                ReDim utf16(0 To LenB(folderID) * 2 - 1): k = 0
                                For j = LBound(ansi) To UBound(ansi)
                                    utf16(k) = ansi(j)
                                    k = k + 2
                                Next j
                                folderID = utf16
                            Else: folderID = ""
                            End If
                        #Else 'Windows
                            folderID = StrConv(MidB$(s, i, n), vbUnicode)
                        #End If
                        i = i + 39: n = InStrB(i, s, vbNullByte) - i
                        If n < 0 Then n = 0
                        If n > 39 Then n = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            parentID = MidB$(s, i, n)
                            ansi = parentID 'Decode ANSI string manually:
                            If LenB(parentID) > 0 Then
                                ReDim utf16(0 To LenB(parentID) * 2 - 1): k = 0
                                For j = LBound(ansi) To UBound(ansi)
                                    utf16(k) = ansi(j)
                                    k = k + 2
                                Next j
                                parentID = utf16
                            Else: parentID = ""
                            End If
                        #Else 'Windows
                            parentID = StrConv(MidB$(s, i, n), vbUnicode)
                        #End If
                        i = i + 121: n = -Int(-(InStrB(i, s, sig3) - i) / 2) * 2
                        If n < 0 Then n = 0
                        If folderID Like folderIdPattern _
                        And parentID Like folderIdPattern Then
                            #If Mac Then 'Encoding of folder names is UTF-32-LE
                                utf32 = MidB$(s, i, n)
                                'UTF-32 can only be converted manually to UTF-16
                                ReDim utf16(LBound(utf32) To UBound(utf32))
                                j = LBound(utf32): k = LBound(utf32)
                                Do While j < UBound(utf32)
                                    If utf32(j + 2) + utf32(j + 3) = 0 Then
                                        utf16(k) = utf32(j)
                                        utf16(k + 1) = utf32(j + 1)
                                        k = k + 2
                                    Else
                                        If utf32(j + 3) <> 0 Then Err.Raise _
                                            vbErrInvalidFormatInResourceFile
                                        codepoint = utf32(j + 2) * &H10000 + _
                                                    utf32(j + 1) * &H100& + _
                                                    utf32(j)
                                        m = codepoint - &H10000
                                        highSurrogate = &HD800& Or (m \ &H400&)
                                        lowSurrogate = &HDC00& Or (m And &H3FF)
                                        utf16(k) = highSurrogate And &HFF&
                                        utf16(k + 1) = highSurrogate \ &H100&
                                        utf16(k + 2) = lowSurrogate And &HFF&
                                        utf16(k + 3) = lowSurrogate \ &H100&
                                        k = k + 4
                                    End If
                                    j = j + 4
                                Loop
                                If k > LBound(utf16) Then
                                    ReDim Preserve utf16(LBound(utf16) To k - 1)
                                    folderName = utf16
                                Else: folderName = ""
                                End If
                            #Else 'On Windows encoding is UTF-16-LE
                                folderName = MidB$(s, i, n)
                            #End If
                            'VBA.Array() instead of just Array() is used in this
                            'function because it ignores Option Base 1
                            odFolders.Add VBA.Array(parentID, folderName), _
                                          folderID
                        End If
                    End If
                    i = InStrB(i + 1, s, sig2) 'Find next sig2 in cid.dat
                Loop
                If odFolders.Count > 0 Then Exit For
            Next vItem
        Loop Until lastChunkEndPos >= lenDatFile _
                Or buffSize >= lenDatFile
        GoTo Continue
Catch:  'This error can happen at chunk boundries, folder might get added twice:
        If Err.Number = vbErrKeyAlreadyExists Then
            odFolders.Remove folderID 'Make sure the folder gets added new again
            Resume 'to avoid folderNames truncated by chunk ends
        End If
        If Err.Number <> vbErrOutOfMemory Then Err.Raise Err
        If buffSize > &HFFFFF Then buffSize = buffSize / 2: Resume Try
        Err.Raise Err 'Raise out of memory error if less than 1 MB RAM available
Continue: On Error GoTo 0

        'Read cid.ini file
        fileNum = FreeFile()
        Open wDir & cid & ".ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: On Error GoTo DecodeUTF8: Err.Raise noErrJustDecodeUTF8
            On Error GoTo 0: b = sUtf16 'b = StrConv(b, vbUnicode) <- UNRELIABLE
        #End If
        Select Case True
        Case dirName Like "Business#" 'Settings files for a business OD account
        'Max 9 Business OneDrive accounts can be signed in at a time.
            mainMount = "": Set libNrToWebColl = New Collection
            For Each line In Split(b, vbNewLine)
                webRoot = "": locRoot = ""
                Select Case Left$(line, InStr(line, " = ") - 1)
                Case "libraryScope" 'One line per synchronized library
                    parts = Split(line, """"): locRoot = parts(9)
                    syncFind = locRoot: syncID = Split(parts(10), " ")(2)
                    If locRoot = "" Then libNr = Split(line, " ")(2)
                    folderType = parts(3): parts = Split(parts(8), " ")
                    siteID = parts(1): webID = parts(2): libID = parts(3)
                    If mainMount = "" And folderType = "ODB" Then
                        mainMount = locRoot: FileName = "ClientPolicy.ini"
                        mainSyncID = syncID: mainSyncFind = syncFind
                    Else: FileName = "ClientPolicy_" & libID & siteID & ".ini"
                    End If
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(FileName)("DavUrlNamespace")
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace"): Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = "" Then Err.Raise vbErrFileNotFound
                    If locRoot = "" Then
                        libNrToWebColl.Add VBA.Array(libNr, webRoot), libNr
                    Else
                        If settDirIsDuplicate Then On Error Resume Next
                        locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                           syncID, syncFind), Key:=locRoot
                        On Error GoTo 0
                    End If
                Case "libraryFolder" 'One line per synchronized library folder
                    parts = Split(line, """"): libNr = Split(line, " ")(3)
                    locRoot = parts(1): syncFind = locRoot
                    syncID = Split(parts(4), " ")(1)
                    s = "": parentID = Left(Split(line, " ")(4), 32)
                    Do  'If not synced at the bottom dir of the library:
                        '   -> add folders below mount point to webRoot
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do
                        s = odFolders(parentID)(1) & "/" & s
                        parentID = odFolders(parentID)(0)
                    Loop
                    webRoot = libNrToWebColl(libNr)(1) & s
                    If settDirIsDuplicate Then On Error Resume Next
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                               syncID, syncFind), Key:=locRoot
                    On Error GoTo 0
                Case "AddedScope" 'One line per folder added as link to personal
                    parts = Split(line, """")                           'library
                    relPath = parts(5): If relPath = " " Then relPath = ""
                    parts = Split(parts(4), " "): siteID = parts(1)
                    webID = parts(2): libID = parts(3): lnkID = parts(4)
                    FileName = "ClientPolicy_" & libID & siteID & lnkID & ".ini"
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(FileName)("DavUrlNamespace") & relPath
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace") & relPath
                                Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = "" Then Err.Raise vbErrFileNotFound
                    s = "": parentID = Left(Split(line, " ")(3), 32)
                    Do 'If link is not at the bottom of the personal library:
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do       'add folders below
                        s = odFolders(parentID)(1) & ps & s 'mount point to
                        parentID = odFolders(parentID)(0)   'locRoot
                    Loop
                    locRoot = mainMount & ps & s
                    If settDirIsDuplicate Then On Error Resume Next
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                              mainSyncID, mainSyncFind), locRoot
                    On Error GoTo 0
                Case Else: Exit For
                End Select
            Next line
        Case dirName = "Personal" 'Settings files for a personal OD account
        'Only one Personal OneDrive account can be signed in at a time.
            For Each line In Split(b, vbNewLine) 'Loop should exit at first line
                If line Like "library = *" Then
                    parts = Split(line, """"): locRoot = parts(3)
                    syncFind = locRoot: syncID = Split(parts(4), " ")(2)
                    Exit For
                End If
            Next line
            On Error Resume Next 'This file may be missing if the personal OD
            webRoot = cliPolColl("ClientPolicy.ini")("DavUrlNamespace") 'account
            On Error GoTo 0                  'was logged out of the OneDrive app
            If locRoot = "" Or webRoot = "" Or cid = "" Then GoTo NextFolder
            If settDirIsDuplicate Then On Error Resume Next
            locToWebColl.Add VBA.Array(locRoot, webRoot & "/" & cid, email, _
                                       syncID, syncFind), Key:=locRoot
            On Error GoTo 0
            If Dir(wDir & "GroupFolders.ini") = "" Then GoTo NextFolder
            'Read GroupFolders.ini file
            cid = "": fileNum = FreeFile()
            Open wDir & "GroupFolders.ini" For Binary Access Read As #fileNum
                ReDim b(0 To LOF(fileNum)): Get fileNum, , b
            Close #fileNum: fileNum = 0
            #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
                sUtf8 = b: On Error GoTo DecodeUTF8
                Err.Raise noErrJustDecodeUTF8
                On Error GoTo 0: b = sUtf16 'StrConv(b, vbUnicode) is UNRELIABLE
            #End If 'Two lines per synced folder from other peoples personal ODs
            For Each line In Split(b, vbNewLine)
                If InStr(line, "BaseUri = ") And cid = "" Then
                    cid = LCase(Mid(line, InStrRev(line, "/") + 1, 16))
                    folderID = Left(line, InStr(line, "_") - 1)
                ElseIf cid <> "" Then
                    If settDirIsDuplicate Then On Error Resume Next
                    locToWebColl.Add VBA.Array(locRoot & ps & odFolders( _
                                     folderID)(1), webRoot & "/" & cid & "/" & _
                                     Mid(line, Len(folderID) + 9), email, _
                                     syncID, syncFind), _
                                Key:=locRoot & ps & odFolders(folderID)(1)
                    On Error GoTo 0
                    cid = "": folderID = ""
                End If
            Next line
        End Select
NextFolder:
        cid = "": s = "": email = "": Set odFolders = Nothing
    Next wDir

    'Clean the finished "dictionary" up, remove trailing "\" and "/"
    Dim tmpColl As Collection: Set tmpColl = New Collection
    For Each vItem In locToWebColl
        locRoot = vItem(0): webRoot = vItem(1): syncFind = vItem(4)
       If Right(webRoot, 1) = "/" Then webRoot = Left(webRoot, Len(webRoot) - 1)
        If Right(locRoot, 1) = ps Then locRoot = Left(locRoot, Len(locRoot) - 1)
        If Right(syncFind, 1) = ps Then _
            syncFind = Left(syncFind, Len(syncFind) - 1)
        tmpColl.Add VBA.Array(locRoot, webRoot, vItem(2), vItem(3), syncFind), _
                    locRoot
    Next vItem
    Set locToWebColl = tmpColl

    #If Mac Then 'deal with syncIDs
        If cloudStoragePathExists Then
            Set tmpColl = New Collection
            For Each vItem In locToWebColl
                locRoot = vItem(0): syncID = vItem(3): syncFind = vItem(4)
                locRoot = Replace(locRoot, syncFind, _
                                           syncIDtoSyncDir(syncID)(1), , 1)
                tmpColl.Add VBA.Array(locRoot, vItem(1), vItem(2)), locRoot
            Next vItem
            Set locToWebColl = tmpColl
        End If
    #End If

    GetLocalPath = GetLocalPath(path, returnAll, pmpo, False): Exit Function
    Exit Function
DecodeUTF8: 'By abusing error handling, code duplication is avoided
    #If Mac Then     'StrConv doesn't work reliably, therefore UTF-8 must
        utf8 = sUtf8 'be transcoded to UTF-16 manually (yes, this is insane)
        ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
        i = LBound(utf8): k = 0
        Do While i <= UBound(utf8) 'Loop through the UTF-8 byte array
            'Determine the number of bytes in the current UTF-8 codepoint
            numBytesOfCodePoint = 1
            If utf8(i) And &H80 Then
                If utf8(i) And &H20 Then
                    If utf8(i) And &H10 Then
                        numBytesOfCodePoint = 4
                    Else: numBytesOfCodePoint = 3: End If
                Else: numBytesOfCodePoint = 2: End If
            End If
            If i + numBytesOfCodePoint - 1 > UBound(utf8) Then Err.Raise 5, _
                "DecodeUtf8", _
                "Incomplete UTF-8 codepoint at the end of the input string."
            'Calculate the Unicode codepoint value from the UTF-8 bytes
            If numBytesOfCodePoint = 1 Then
                codepoint = utf8(i)
            Else: codepoint = utf8(i) And (2 ^ (7 - numBytesOfCodePoint) - 1)
                For j = 1 To numBytesOfCodePoint - 1
                    codepoint = (codepoint * 64) + (utf8(i + j) And &H3F)
                Next j
            End If
            'Convert the Unicode codepoint to UTF-16LE bytes
            If codepoint < &H10000 Then
                utf16(k) = CByte(codepoint And &HFF&)
                utf16(k + 1) = CByte(codepoint \ &H100&)
                k = k + 2
            Else 'Codepoint must be encoded as surrogate pair
                m = codepoint - &H10000
                '(m \ &H400&) = most significant 10 bits of m
                highSurrogate = &HD800& Or (m \ &H400&)
                '(m And &H3FF) = least significant 10 bits..
                lowSurrogate = &HDC00& Or (m And &H3FF)
                'Concatenate highSurrogate and lowSurrogate as UTF-16LE bytes
                utf16(k) = CByte(highSurrogate And &HFF&)
                utf16(k + 1) = CByte(highSurrogate \ &H100&)
                utf16(k + 2) = CByte(lowSurrogate And &HFF&)
                utf16(k + 3) = CByte(lowSurrogate \ &H100&)
                k = k + 4
            End If
            i = i + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
        Loop
        If k > 0 Then
            ReDim Preserve utf16(k - 1)
            sUtf16 = utf16
        Else: sUtf16 = ""
        End If
        Resume Next 'Clear the error object, and jump back to the statement
    #End If         'after where the pseudo "Error" was raised.
End Function








