VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiny Server"
   ClientHeight    =   4365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmServer.frx":030A
   ScaleHeight     =   4365
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   360
   End
   Begin VB.CommandButton Configure 
      Caption         =   "Configure Server"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop Server"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "Start Server"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TextBox 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Index           =   0
      Left            =   6960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label Label3 
      Caption         =   "Website :"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://connect.to/tinyserver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2880
      MouseIcon       =   "frmServer.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Click to visit the TinyServer Website"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Message Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show TinyServer"
      End
      Begin VB.Menu mnuStartServer 
         Caption         =   "Start Server"
      End
      Begin VB.Menu mnuStopServer 
         Caption         =   "Stop Server"
      End
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.
'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

'Maximum no of connections
Private Const MAX_CONNECTIONS = 100
'Time in seconds for connection timeout
Private Const MAX_TIME = 4

Dim timeOut(1 To MAX_CONNECTIONS) As Integer

Private Sub Form_Load()
    'Terminate if already running
    If App.PrevInstance Then
        MsgBox "TinyServer already running...", vbInformation
        End
    End If
    
    Dim sFile As String
    sFile = Space(256)
    sFile = Left(sFile, GetCurrentDirectory(Len(sFile), sFile))
    sFile = sFile & "\server.ini"
    Dim i As Long
        For i = 1 To MAX_CONNECTIONS
            Load tcpServer(i)
        Next
    
    wwwRoot = Space(256)
    wwwRoot = Left(wwwRoot, GetPrivateProfileString("config", "wwwroot", "NULL", wwwRoot, Len(wwwRoot), sFile))
    DefaultPage = Space(64)
    DefaultPage = Left(DefaultPage, GetPrivateProfileString("config", "defaultpage", "NULL", DefaultPage, Len(DefaultPage), sFile))
    PortNum = GetPrivateProfileInt("config", "port", 0, sFile)
    
    If StrComp(wwwRoot, "NULL") = 0 Or StrComp(DefaultPage, "NULL") = 0 Or PortNum = 0 Then
        sFile = Left(sFile, GetCurrentDirectory(Len(sFile), sFile))
        wwwRoot = sFile
        DefaultPage = "index.htm"
        PortNum = 80
        WritePrivateProfileString "config", "wwwroot", sFile, sFile + "\server.ini"
        sFile = sFile + "\server.ini"
        WritePrivateProfileString "config", "defaultpage", "index.htm", sFile
        WritePrivateProfileString "config", "port", Str(80), sFile
        MsgBox "Configuration file not found or corrupted, defaults loaded", vbInformation
    End If
    setSysTrayIcon
End Sub



Private Sub Label2_Click()
    Call ShellExecute(0, "open", _
        "http://connect.to/tinyserver", _
        vbNullString, vbNullString, 1)
End Sub

Private Sub StartButton_Click()
    ' Set the LocalPort property to an integer.
    ' Then invoke the Listen method.
    If tcpServer(0).State <> sckListening Then
        tcpServer(0).LocalPort = PortNum
        tcpServer(0).Listen
        TextBox.Text = "TinyServer started . . ." + vbCrLf + "Listening on port : " + Str(PortNum) + vbCrLf
    Else
        TextBox.Text = TextBox.Text + vbCrLf + "TinyServer already started!!!"
    End If
End Sub

Private Sub StopButton_Click()
    If tcpServer(0).State <> sckListening Then
        TextBox.Text = TextBox.Text + vbCrLf + "TinyServer not running!!!"
        Exit Sub
    End If

    tcpServer(0).Close
    If tcpServer(0).State <> sckListening Then
        TextBox.Text = TextBox.Text + vbCrLf + "TinyServer stopped..." + vbCrLf
    End If
End Sub

Private Sub Configure_Click()
    Dim fOptions As New frmOptions
    If tcpServer(0).State = sckListening Then
        MsgBox "Please stop the server before configuring", vbExclamation
    Else
        fOptions.Show vbModal
    End If
End Sub

Private Sub tcpServer_ConnectionRequest _
(Index As Integer, ByVal requestID As Long)
    Dim i As Integer
    ' Accept the request with the requestID
    ' parameter.
    If Index = 0 Then
        For i = 1 To 100
            If tcpServer(i).State = sckClosed Then
                tcpServer(i).LocalPort = 0
                tcpServer(i).Accept requestID
                TextBox.Text = TextBox.Text + vbCrLf + "Connection from : " + tcpServer(i).RemoteHostIP
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub tcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As Integer
    Dim inData As String
    
    tcpServer(Index).GetData inData
    
    Call serveRequest(Index, inData)
    'tcpServer(Index).Close
    
End Sub
Private Sub serveRequest(ind As Integer, inData As String)
    Dim rServer As Winsock
    Dim i As Integer
    Dim fileNum As Integer
    
    Dim Method As String
    Dim Request As String
    Dim lRequest As String
    Dim httpVersion As String
    Dim Accept() As String
    Dim AcceptLanguage As String
    Dim UserAgent As String
    Dim Connection As String
    Dim Referer As String
    Dim Host As String
    Dim AcceptEncoding As String
    Dim Cookie As String
    Dim SplitHeader() As String
    Dim SplitTemp() As String
    Dim sFile As String
    Dim outData As String
    Dim fileDate As Date
    
    i = 1
    Set rServer = tcpServer(ind)
    SplitHeader = Split(inData, vbCrLf)
    SplitTemp = Split(SplitHeader(0))
    Method = SplitTemp(0)
    Request = SplitTemp(1)
    httpVersion = SplitTemp(2)
    While StrComp(SplitHeader(i), "") <> 0
        SplitTemp = Split(SplitHeader(i), ": ")
        Select Case SplitTemp(0)
        Case "Accept"
        Accept = Split(SplitTemp(1), ", ")
        Case "Accept-Language"
        AcceptLanguage = SplitTemp(1)
        Case "Accept-Encoding"
        AcceptEncoding = SplitTemp(1)
        Case "User-Agent"
        UserAgent = SplitTemp(1)
        Case "Host"
        Host = SplitTemp(1)
        Case "Connection"
        Connection = SplitTemp(1)
        Case "Cookie"
        Cookie = SplitTemp(1)
        End Select
        i = i + 1
    Wend
    
    If StrComp(Method, "GET") <> 0 Then
        rServer.SendData errorPage(405, "Method not allowed : <b>" + Method + "</b>")
        Exit Sub
    End If
    
    SplitTemp = Split(Request, "/")
    lRequest = Join(SplitTemp, "\")
    
    If StrComp(Right(lRequest, 1), "\", vbTextCompare) = 0 Then
        sFile = wwwRoot + lRequest + DefaultPage
        If Len(Dir$(sFile)) = 0 And Len(Dir$(wwwRoot + lRequest + "*.*")) <> 0 Then
            rServer.SendData errorPage(403, "You do not have the permission to access <b>" + Request + "</b> on this server")
            Exit Sub
        End If
    Else
        sFile = wwwRoot + lRequest
    End If
    If Len(Dir$(sFile)) = 0 Then
        rServer.SendData errorPage(404, "The following page was not found on this server : <b>" + Request + "</b>")
        Exit Sub
    End If
    
    fileDate = FileDateTime(sFile)
    SplitTemp = Split(sFile, ".")
    rServer.SendData makemimeHeader(200, FileLen(sFile), SplitTemp(1), Format(fileDate, "ddd, d mmm yyyy hh:mm:ss ") + "GMT", Connection)
    rServer.SendData readFile(sFile)
    
End Sub


Function errorPage(errNum As Integer, errMessage As String) As String
    Dim responseHeader As String
    Dim responseData As String
    Dim sDate As Date
    Dim sTime As Date

    sDate = Date
    sTime = Time
    responseData = "<html><head>" + vbCrLf _
        + "<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>" + vbCrLf _
        + "<title>Error : " + Str(errNum) + "</title></head><body><table border='0' width='100%'>" + vbCrLf _
        + "<tr><td width='100%' bgcolor='#00FFFF'><h2>Error : " + Str(errNum) + " " + getReason(errNum) + "<h2></td></tr>" + vbCrLf _
        + "<tr><td width='100%' height='200'>" + errMessage + "</td></tr>" + vbCrLf _
        + "<tr><td width='100%' bgcolor='#C0C0C0'><center><b>TinyServer v1.0.1</b><br>Copyright &copy Saurabh 2001-2002</center>" + vbCrLf _
        + "</td></tr></table></body></html>"
    errorPage = makemimeHeader(errNum, Len(responseData), "htm", Format(sDate, "ddd, d mmm yyyy ") + Format(sTime, " hh:mm:ss ") + "GMT", "keep-alive") + responseData
End Function


Function makemimeHeader(httpCode As Integer, dataLength As Long, fileExt As String, lastModified As String, conType As String) As String
    Dim mimeType As String
    Dim sDate As Date
    Dim sTime As Date
    Dim Authenticate As String
    
    sDate = Date
    sTime = Time
    If httpCode = 401 Then
        Authenticate = "WWW-Authenticate: Basic realm=" + Chr(34) + "TinyServer Admin" + Chr(34) + vbCrLf
    Else
        Authenticate = ""
    End If
    
    Select Case fileExt
        Case "doc"
        mimeType = "application/msword"
        Case "rtf"
        mimeType = "application/rtf"
        Case "zip"
        mimeType = "application/zip"
        Case "jpg"
        mimeType = "image/jpeg"
        Case "jpeg"
        mimeType = "image/jpeg"
        Case "gif"
        mimeType = "image/gif"
        Case "bmp"
        mimeType = "image/x-xbitmap"
        Case "mail"
        mimeType = "message/RFC822"
        Case "txt"
        mimeType = "text/plain"
        Case "htm"
        mimeType = "text/html"
        Case "html"
        mimeType = "text/html"
        Case "mpg"
        mimeType = "video/mpeg"
        Case "mpeg"
        mimeType = "video/mpeg"
        Case "mov"
        mimeType = "video/quicktime"
        Case "wmv"
        mimeType = "video/x-msvideo"
        Case "avi"
        mimeType = "video/avi"
        Case "mid"
        mimeType = "audio/basic"
        Case "wav"
        mimeType = "audio/wav"
        Case Else
        mimeType = "text/plain"
    End Select
    
    makemimeHeader = "HTTP/1.0 " + Str(httpCode) + " " + getReason(httpCode) + vbCrLf _
                   + "Date: " + Format(sDate, "ddd, d mmm yyyy ") + Format(sTime, " hh:mm:ss ") + "GMT" + vbCrLf _
                   + "Server: TinyServer v1.0.1" + vbCrLf _
                   + "MIME-version: 1.0" + vbCrLf _
                   + "Content-type: " + mimeType + vbCrLf _
                   + "Last-modified: " + lastModified + vbCrLf _
                   + "Connection: " + conType + vbCrLf _
                   + Authenticate _
                   + "Content-length: " + Str(dataLength) + vbCrLf + vbCrLf
    'MsgBox (makemimeHeader)
End Function

Function getReason(httpCode As Integer) As String
    Select Case httpCode
        Case 200
        getReason = "OK"
        Case 201
        getReason = "Created"
        Case 202
        getReason = "Accepted"
        Case 204
        getReason = "No Content"
        Case 301
        getReason = "Moved Permanently"
        Case 302
        getReason = "Moved Temporarily"
        Case 304
        getReason = "Not Modified"
        Case 400
        getReason = "Bad Request"
        Case 401
        getReason = "Unauthorized"
        Case 403
        getReason = "Forbidden"
        Case 404
        getReason = "Not Found"
        Case 405
        getReason = "Method not allowed"
        Case 500
        getReason = "Internal Server Error"
        Case 501
        getReason = "Not Implemented"
        Case 502
        getReason = "Bad Gateway"
        Case 503
        getReason = "Service Unavailable"
        Case Else
        getReason = "Unknown"
    End Select
End Function


Private Sub tcpServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    TextBox.Text = TextBox.Text + vbCrLf + "Message form thread " + Str(Index) + ", Code : " + Str(Number) + " Description : " + Description
    tcpServer(Index).Close
    If tcpServer(Index).State = sckClosed Then
        TextBox.Text = TextBox.Text + vbCrLf + "Connection Closed"
        timeOut(Index) = 0
    End If
End Sub


Private Sub setSysTrayIcon()
    'Click this button to add an icon to the taskbar status area.

    'Set the individual values of the NOTIFYICONDATA data type.
    nid.cbSize = Len(nid)
    nid.hwnd = Server.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Server.Icon
    nid.szTip = "Tiny Server" & vbNullChar

    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area.
    Shell_NotifyIcon NIM_ADD, nid
End Sub


Private Sub Form_Terminate()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            DisplayForm
        Case WM_RBUTTONDOWN
            PopupMenu mnuPopup
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_Resize()
    If WindowState = 1 And Visible = True Then
        Me.Hide
    End If
End Sub

'System tray menu handling subroutines
Private Sub mnuExit_Click()
    Unload Me
End Sub
Private Sub mnuStart_Click()
    StartButton_Click
End Sub
Private Sub mnuStop_Click()
    StopButton_Click
End Sub
Private Sub mnuConfigure_Click()
    Configure_Click
End Sub
Private Sub mnuShow_Click()
    DisplayForm
End Sub
Private Sub mnuAbout_Click()
    MsgBox "                       TinyServer v1.1" + vbCrLf _
         + "Programmed by Saurabh (saurabh@yep.com)" + vbCrLf _
         + "               http://connect.to/tinyserver", 0, "About TinyServer"
End Sub
Private Sub DisplayForm()
    If Visible = False Then
        'Display form
        WindowState = 0
        Visible = True
    End If
    SetFocus
End Sub


Private Sub Timer1_Timer()
    Dim i As Integer
    For i = 1 To MAX_CONNECTIONS
        If tcpServer(i).State = sckConnected Then
            If timeOut(i) > MAX_TIME Then
                tcpServer(i).Close
                timeOut(i) = 0
            Else
                timeOut(i) = timeOut(i) + 1
            End If
        End If
    Next i
End Sub
