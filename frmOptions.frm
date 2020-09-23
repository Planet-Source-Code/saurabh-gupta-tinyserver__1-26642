VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure TinyServer"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox PortBox 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton OKbutton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox DPage 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox wwwDir 
      Height          =   1665
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "( Default : 80 )"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Port :"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Wepage Directory :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Default Page :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim configChanged As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub DPage_Change()
    configChanged = True
End Sub

Private Sub Drive1_Change()
    wwwDir.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    configChanged = False
    PortBox.Text = PortNum
    DPage.Text = DefaultPage
    wwwDir.Path = wwwRoot
End Sub

Private Sub OKbutton_Click()
    If configChanged Then
        Dim sFile As String
        sFile = Space(256)
        sFile = Left(sFile, GetCurrentDirectory(Len(sFile), sFile)) + "\server.ini"
        WritePrivateProfileString "config", "wwwroot", wwwDir.Path, sFile
        WritePrivateProfileString "config", "defaultpage", DPage.Text, sFile
        WritePrivateProfileString "config", "port", PortBox.Text, sFile
        PortNum = PortBox.Text
        DefaultPage = DPage.Text
        wwwRoot = wwwDir.Path
    End If
    Unload Me
End Sub

Private Sub PortBox_Change()
    configChanged = True
    If IsNumeric(PortBox.Text) = False Then
        PortBox.Text = Val(PortBox.Text)
    End If
End Sub

Private Sub wwwDir_Change()
    configChanged = True
End Sub
