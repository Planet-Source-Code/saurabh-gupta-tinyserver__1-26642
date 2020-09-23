Attribute VB_Name = "Module1"
Public wwwRoot As String
Public DefaultPage As String
Public PortNum As Integer

Public Declare Function GetPrivateProfileInt Lib "kernel32" _
    Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal nDefault As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long


Public Declare Function GetCurrentDirectory Lib "kernel32" _
        Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long


' Constants with OpenFile API call.
Private Const OF_READ = &H0

' Structure filled in by OpenFile API call.
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(128) As Byte
End Type

' declarations for the API functions that this class uses.
Private Declare Function OpenFile Lib "kernel32" _
    (ByVal lpFileName As String, _
     lpReOpenBuff As OFSTRUCT, _
     ByVal wStyle As Long) As Long

Private Declare Function hread Lib "kernel32" Alias "_hread" _
    (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long

Private Declare Function lclose Lib "kernel32" Alias "_lclose" _
    (ByVal hFile As Long) As Long



Public Function readFile(sFile As String) As String
    Dim hInp As Long
    Dim size As Long
    Dim inpOFS As OFSTRUCT
    
    size = FileLen(sFile)
    hInp = OpenFile(sFile, inpOFS, OF_READ)
    If (hInp <> -1) Then
        readFile = String(size, "*")
        hread hInp, ByVal readFile, size
    End If
    lclose hInp
End Function
