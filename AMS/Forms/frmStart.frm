VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2505
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5500
      Left            =   4755
      Top             =   630
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   1050
      Left            =   1170
      TabIndex        =   0
      Top             =   705
      Width           =   3285
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading interface and data... please wait !"
         ForeColor       =   &H8000000D&
         Height          =   705
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   2955
      End
   End
   Begin VB.Timer Timer131 
      Interval        =   40
      Left            =   4770
      Top             =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   30
      Width           =   3900
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   0
      Picture         =   "frmStart.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15690
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   540
      TabIndex        =   2
      Top             =   2055
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   315
      Picture         =   "frmStart.frx":0E2D
      Stretch         =   -1  'True
      Top             =   810
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4065
      X2              =   4065
      Y1              =   1845
      Y2              =   2205
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "L o a d i n g  d a t a  /  i n t e r f a c e" & vbCrLf & vbCrLf & "P l e a s e   w a i t !"
    Label11.Caption = "g g g"

End Sub

Private Sub Timer131_Timer()
Dim Num2div As Integer

Start:
    Label11.Caption = " " & Label11.Caption
    If Len(Label11.Caption) >= Num2div + 45 Then Label11.Caption = "g g g": GoTo Start:
End Sub

Private Sub Timer1_Timer()
    frmMain.Show
    Unload Me
End Sub

