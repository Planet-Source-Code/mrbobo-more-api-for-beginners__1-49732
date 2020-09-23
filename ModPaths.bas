Attribute VB_Name = "ModPaths"
Option Explicit
Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
Private Declare Function PathAddExtension Lib "shlwapi.dll" Alias "PathAddExtensionA" (ByVal pszPath As String, ByVal pszExt As String) As Long
Private Declare Function PathBuildRoot Lib "shlwapi.dll" Alias "PathBuildRootA" (ByVal szRoot As String, ByVal iDrive As Long) As Long
Private Declare Function PathCompactPathEx Lib "shlwapi.dll" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathGetDriveNumber Lib "shlwapi.dll" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecA" (ByVal pszFile As String, ByVal pszSpec As String) As Long
Private Declare Function PathParseIconLocation Lib "shlwapi.dll" Alias "PathParseIconLocationA" (ByVal pszIconFile As String) As Long
Private Declare Sub PathQuoteSpaces Lib "shlwapi.dll" Alias "PathQuoteSpacesA" (ByVal lpsz As String)
Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal pszPath As String) As Long
Private Declare Sub PathRemoveExtension Lib "shlwapi.dll" Alias "PathRemoveExtensionA" (ByVal pszPath As String)
Private Declare Function PathRenameExtension Lib "shlwapi.dll" Alias "PathRenameExtensionA" (ByVal pszPath As String, ByVal pszExt As String) As Long
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Private Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Private Declare Sub PathUnquoteSpaces Lib "shlwapi.dll" Alias "PathUnquoteSpacesA" (ByVal lpsz As String)
Public Function AddBackslash(mPath As String) As String
    mPath = mPath & Chr(0)
    PathAddBackslash mPath
    AddBackslash = mPath
End Function
Public Function AddExtension(mPath As String, mExt As String) As String
    If Left(mExt, 1) <> "." Then mExt = "." & mExt
    mPath = mPath & String(Len(mExt), Chr(0))
    PathAddExtension mPath, mExt
    AddExtension = mPath
End Function
Public Function DriveLetterFromPath(mPath As String) As String
    Dim DrvL As String * 3
    PathBuildRoot DrvL, PathGetDriveNumber(mPath)
    DriveLetterFromPath = DrvL
End Function
Public Function FitPath(mPath As String, numChars As Long) As String
    Dim SmPath As String * 255
    PathCompactPathEx SmPath, mPath, numChars, 0
    FitPath = SmPath
End Function
Public Function FileExists(mPath As String) As Boolean
    FileExists = CBool(PathFileExists(mPath))
End Function
Public Function IsNetworkPath(mPath As String) As Boolean
    IsNetworkPath = CBool(PathIsNetworkPath(mPath))
End Function
Public Function IsURL(mPath As String) As Boolean
    IsURL = CBool(PathIsURL(mPath))
End Function
Public Function GetDriveNumber(mPath As String) As Long
    GetDriveNumber = PathGetDriveNumber(mPath)
End Function
Public Function IsFolder(mPath As String) As Boolean
    IsFolder = CBool(PathIsDirectory(mPath))
End Function
Public Function IsFolderEmpty(mPath As String) As Boolean
    IsFolderEmpty = CBool(PathIsDirectoryEmpty(mPath))
End Function
Public Function IsFileExtenstion(mPath As String, mFileSpec As String) As Boolean
    IsFileExtenstion = CBool(PathMatchSpec(mPath, mFileSpec))
End Function
Public Function IconNumberOnly(mPath As String) As Long
    IconNumberOnly = PathParseIconLocation(mPath)
End Function
Public Function QuotePath(mPath As String) As String
    mPath = mPath & Chr(0) & Chr(0)
    PathQuoteSpaces mPath
    QuotePath = mPath
End Function
Public Function RemoveBackslash(mPath As String) As String
    PathRemoveBackslash mPath
    RemoveBackslash = mPath
End Function
Public Function RemoveExtension(mPath As String) As String
    PathRemoveExtension mPath
    RemoveExtension = mPath
End Function
Public Function ChangeExtension(mPath As String, mExt As String) As String
    If Left(mExt, 1) <> "." Then mExt = "." & mExt
    PathRenameExtension mPath, mExt
    ChangeExtension = mPath
End Function
Public Function FileOnly(mPath As String) As String
    PathStripPath mPath
    FileOnly = mPath
End Function
Public Function PathOnly(mPath As String) As String
    Dim mFile As String
    mFile = mPath
    PathStripPath mFile
    mFile = Left(mPath, Len(mPath) - Len(StripNulls(mFile)) - 1)
    If Len(mFile) = 2 Then mFile = mFile & "\"
    PathOnly = mFile
End Function
Public Function DriveOnly(mPath As String) As String
    PathStripToRoot mPath
    DriveOnly = mPath
End Function
Public Function UnQuotePath(mPath As String) As String
    PathUnquoteSpaces mPath
    UnQuotePath = mPath
End Function
Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


