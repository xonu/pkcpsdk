VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "pkColorPicker Remote Control Demonstration"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6855
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdSetColor 
      Caption         =   "PKCP.Color2 ="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtColor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSetColor 
      Caption         =   "PKCP.Color1 ="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtColor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdSetColors 
      Caption         =   "PKCP.SetColors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "PKCP.Show (PKCP.State = 1)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "PKCP.Hide (PKCP.State = 2)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "PKCP.Launch"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "PKCP.Quit (PKCP.State = 0)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Timer tmrGetColor 
      Interval        =   10
      Left            =   5640
      Top             =   240
   End
   Begin VB.Timer tmrStatus 
      Interval        =   100
      Left            =   2280
      Top             =   2040
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "PKCP.Color2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "PKCP.Color1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "PKCP.State ="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 PKCP.Init
End Sub

Private Sub cmdLaunch_Click()
 If Not PKCP.Launch() Then
  MsgBox "pkColorPicker not found", vbCritical, "ERROR"
 End If

 ' start hidden
 'If Not PKCP.Launch("/s") Then
 ' MsgBox "pkColorPicker not found", vbCritical, "ERROR"
 'End If
End Sub

Private Sub cmdQuit_Click()
 'PKCP.State = PKCP_QUIT
 PKCP.Quit
End Sub

Private Sub cmdShow_Click()
 'PKCP.State = PKCP_SHOW
 PKCP.Show
End Sub

Private Sub cmdHide_Click()
 'PKCP.State = PKCP_HIDE
 PKCP.Hide
End Sub

Private Sub cmdSetColor_Click(Index As Integer)
 Select Case Index
  Case 1: PKCP.Color1 = Val(txtColor(1))
  Case 2: PKCP.Color2 = Val(txtColor(2))
 End Select
End Sub

Private Sub cmdSetColors_Click()
 PKCP.SetColors Val(txtColor(1)), Val(txtColor(2))
End Sub

Private Sub tmrGetColor_Timer()
 lblColor(2).BackColor = PKCP.Color2
 lblColor(1).BackColor = PKCP.Color1
End Sub

Private Sub tmrStatus_Timer()
 Dim Descr As String
 Select Case PKCP.State
  Case PKCP_QUIT: Descr = "(not running)"
  Case PKCP_SHOW: Descr = "(visible)"
  Case PKCP_HIDE: Descr = "(hidden)"
 End Select
 lblStatus.Caption = "PKCP.State = " & PKCP.State & " " & Descr
End Sub
