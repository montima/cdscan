Attribute VB_Name = "Module1"
Option Explicit

  Public Const INVALID_HANDLE_VALUE = -1
  Public Const MAX_PATH = 260
  Public Const FILE_ATTRIBUTE_READONLY = &H1
  Public Const FILE_ATTRIBUTE_HIDDEN = &H2
  Public Const FILE_ATTRIBUTE_SYSTEM = &H4
  Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
  Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
  Public Const FILE_ATTRIBUTE_NORMAL = &H80
  Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
  Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
  
 
  Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
  End Type

  Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
  End Type

  Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Long
  End Type

  Declare Function CreateDirectory Lib "kernel32" _
      Alias "CreateDirectoryA" _
      (ByVal lpPathName As String, _
       lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

  Declare Function CopyFile Lib "kernel32" _
      Alias "CopyFileA" _
      (ByVal lpExistingFileName As String, _
       ByVal lpNewFileName As String, _
       ByVal bFailIfExists As Long) As Long

  Declare Function FindFirstFile Lib "kernel32" _
      Alias "FindFirstFileA" _
      (ByVal lpFileName As String, _
      lpFindFileData As WIN32_FIND_DATA) As Long

  Declare Function FindNextFile Lib "kernel32" _
      Alias "FindNextFileA" _
      (ByVal hFindFile As Long, _
       lpFindFileData As WIN32_FIND_DATA) As Long

  Declare Function FindClose Lib "kernel32" _
      (ByVal hFindFile As Long) As Long

  Declare Function GetFileAttributes Lib "kernel32" _
      Alias "GetFileAttributesA" _
      (ByVal lpFileName As String) As Long

  Declare Function LockWindowUpdate Lib "user32" _
      (ByVal hwndLock As Long) As Long




Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
       (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
       (ByVal nDrive As String) As Long

