VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAccounts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   0
      ScaleHeight     =   4950
      ScaleWidth      =   6090
      TabIndex        =   19
      Top             =   4140
      Width           =   6120
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAccounts 
         Height          =   4920
         Left            =   30
         TabIndex        =   20
         Top             =   15
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   8678
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.PictureBox picAccounts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      ScaleHeight     =   4080
      ScaleWidth      =   6075
      TabIndex        =   7
      Top             =   0
      Width           =   6105
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   90
         ScaleHeight     =   1380
         ScaleWidth      =   5880
         TabIndex        =   12
         Top             =   2550
         Width           =   5910
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   975
            Width           =   3120
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1785
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   150
            Width           =   1470
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Delete Record"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1785
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   585
            Width           =   1470
         End
         Begin VB.CommandButton cmdNew 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&New Record"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   150
            Width           =   1470
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Save Record"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   585
            Width           =   1470
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1290
            Left            =   4245
            ScaleHeight     =   1260
            ScaleWidth      =   1545
            TabIndex        =   13
            Top             =   45
            Width           =   1575
            Begin VB.PictureBox PicPrev 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   180
               Picture         =   "frmAccounts.frx":0000
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   17
               Top             =   495
               Width           =   345
            End
            Begin VB.PictureBox PicNext 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1080
               Picture         =   "frmAccounts.frx":0311
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   16
               Top             =   495
               Width           =   345
            End
            Begin VB.PictureBox PicFirst 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   630
               Picture         =   "frmAccounts.frx":0637
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   15
               Top             =   105
               Width           =   345
            End
            Begin VB.PictureBox PicLast 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   630
               Picture         =   "frmAccounts.frx":0947
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   14
               Top             =   885
               Width           =   345
            End
            Begin VB.Label LblNav 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   660
               TabIndex        =   18
               Top             =   540
               Visible         =   0   'False
               Width           =   285
            End
         End
      End
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   825
         Width           =   1200
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         TabIndex        =   0
         Top             =   1350
         Width           =   3435
      End
      Begin VB.ComboBox cmbHead 
         Height          =   315
         ItemData        =   "frmAccounts.frx":0C6C
         Left            =   2115
         List            =   "frmAccounts.frx":0C7F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1905
         Width           =   3465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Accounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2040
         TabIndex        =   22
         Top             =   0
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Accounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2055
         TabIndex        =   21
         Top             =   15
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -15
         X2              =   7230
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   195
         Left            =   405
         TabIndex        =   11
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   195
         Left            =   405
         TabIndex        =   9
         Top             =   1965
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   0
         Picture         =   "frmAccounts.frx":0CB3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6090
      End
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Clear Me
'Calling form load event
    Call Form_Load

'Changing the mode of button
    Modes False, True, Me

'Setting focus on TxtName
    txtName.SetFocus

'Setting the Caption of LblNav
    LblNav.Caption = "1"
    
'Highlighting TxtName
    High txtName

'Unlock Navigation
    Picture2.Enabled = True

End Sub

Private Sub cmdDelete_Click()
    MsgBox "In process"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
'Clearing controls of the form
    Clear Me

'Calling Max Number
    Call AutoId

'Changing the mode of button
    Modes True, False, Me

'Setting focus on TxtName
    txtName.SetFocus

'Locking Navigation
    Picture2.Enabled = False
End Sub

Private Sub cmdSave_Click()
'If Name field is blank
    If txtName.Text = "" Then
        MsgBox "enter Account Title", vbCritical, "Blank fields"
        txtName.SetFocus
        Exit Sub
    End If

'If account head not selected then
    If cmbHead.ListIndex = -1 Then
        MsgBox "Select Account Head", vbCritical, "Blank fileds"
        cmbHead.SetFocus
        Exit Sub
    End If
    

'Finding Head Account ID through CMBHead.Text
    If cmbHead.Text = "Assets" Then
        mHeadId = 1
    ElseIf cmbHead.Text = "Liabilities" Then
        mHeadId = 2
    ElseIf cmbHead.Text = "Capital" Then
        mHeadId = 3
    ElseIf cmbHead.Text = "Revenue" Then
        mHeadId = 4
    ElseIf cmbHead.Text = "Expense" Then
        mHeadId = 5
    End If


'If New record
    If cmdNew.Enabled = False Then
                
        If InsertRecord.State = 1 Then InsertRecord.Close
            
        Set InsertRecord = New ADODB.Recordset
        InsertRecord.Open "Insert into Accounts values(" & Val(txtId) & ", " & Val(mHeadId) & ", '" & txtName.Text & "', date())", Con, adOpenDynamic, adLockOptimistic
        MsgBox "Record saved", vbInformation, "Saved"
           
'Updating Maximum Number
        UpdateMaxNumber "AcId", Val(txtId)
    
    ElseIf cmdNew.Enabled = True Then
        If UpdateRecord.State = 1 Then UpdateRecord.Close
        
        Set UpdateRecord = New ADODB.Recordset
        UpdateRecord.Open "Update Accounts set HeadId = " & Val(mHeadId) & ", Title = '" & txtName.Text & "' where Id = " & Val(txtId) & " ", Con, adOpenDynamic, adLockOptimistic
        MsgBox "Record updated", vbInformation, "Saved"
    End If
    
    cmdCancel_Click
    LblNav.Caption = "1"
End Sub

Private Sub fgAccounts_DblClick()
Dim mRow As Integer
    mRow = fgAccounts.RowSel
        
    txtId.Text = fgAccounts.TextMatrix(mRow, 0)
    txtName.Text = fgAccounts.TextMatrix(mRow, 1)
    cmbHead.Text = fgAccounts.TextMatrix(mRow, 2)
    
    LblNav.Caption = Val(txtId)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Changing Control focus on Enter
    ChangeFocusOnEnter KeyAscii, Me

End Sub

Private Sub Form_Load()
'Setting up flexgrid data
    Call fgAccountsData
    Call GridSetting
    
'Setting Navigational Recordset
    If RsNAV.State = 1 Then RsNAV.Close
        
        Set RsNAV = New ADODB.Recordset
        RsNAV.Open "Select Id, headId, Title from Accounts", Con, adOpenStatic, adLockOptimistic
            If RsNAV.RecordCount > 0 Then
                Call BoundData  'Showing Data in textboxes
            Else
                Exit Sub
            End If
            
    
End Sub

Public Sub fgAccountsData()
    If RS.State = 1 Then RS.Close
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from ViewHeadWise order by Id", Con, adOpenDynamic, adLockPessimistic

    Set fgAccounts.DataSource = RS
    
End Sub

Public Sub GridSetting()
    With fgAccounts
        .ColWidth(0) = 1000
        .ColWidth(1) = 3150
        .ColWidth(2) = 1500
    
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Account Title"
        .TextMatrix(0, 2) = "Head Account"
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        
        .ColAlignment(0) = 5
        
        .RowHeight(0) = 400
    End With
End Sub

Public Sub BoundData()
    txtId.Text = Val(RsNAV(0))
    txtName.Text = RsNAV(2)
    
    If Val(RsNAV(1)) = 1 Then
        cmbHead.Text = "Assets"
    ElseIf Val(RsNAV(1)) = 2 Then
        cmbHead.Text = "Liabilities"
    ElseIf Val(RsNAV(1)) = 3 Then
        cmbHead.Text = "Capital"
    ElseIf Val(RsNAV(1)) = 4 Then
        cmbHead.Text = "Revenue"
    ElseIf Val(RsNAV(1)) = 5 Then
        cmbHead.Text = "Expense"
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Left = Me.Left + 1200
    Me.Top = Me.Top + 200
End Sub

Private Sub PicFirst_Click()
    RsNAV.MoveFirst
    LblNav.Caption = "1"
    
    If RsNAV.BOF = True Then
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
    Else
        Call BoundData
    End If

End Sub

Private Sub PicLast_Click()
    RsNAV.MoveLast
    LblNav.Caption = Val(RsNAV.RecordCount)
    
    If RsNAV.EOF = True Then
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call BoundData
    End If

End Sub

Private Sub PicNext_Click()
    RsNAV.MoveNext

    LblNav.Caption = Val(LblNav) + 1

    If RsNAV.EOF = True Then
        LblNav.Caption = Val(LblNav) - 1
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call BoundData
    End If
End Sub

Private Sub PicPrev_Click()
    RsNAV.MovePrevious
    LblNav.Caption = Val(LblNav) - 1
    
    
    If RsNAV.BOF = True Then
        LblNav.Caption = Val(LblNav) + 1
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
    Else
        Call BoundData
    End If
    
End Sub

Public Sub AutoId()
'Calling MaxNumber function to get Auto Id for the record
    MaxNumber "AcId", "MaxCode"
    txtId.Text = Val(MaxNmbr)
End Sub

Private Sub txtId_GotFocus()
    txtName.SetFocus
End Sub
