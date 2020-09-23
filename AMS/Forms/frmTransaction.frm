VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransaction 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSrchGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7470
      Left            =   2265
      ScaleHeight     =   7440
      ScaleWidth      =   5490
      TabIndex        =   21
      Top             =   810
      Visible         =   0   'False
      Width           =   5520
      Begin VB.TextBox TxtGrdSrch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   330
         Left            =   105
         TabIndex        =   22
         Top             =   750
         Width           =   5310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshSearch 
         Height          =   6180
         Left            =   90
         TabIndex        =   23
         Top             =   1185
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   10901
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483645
         BackColorBkg    =   15790320
         GridColorFixed  =   0
         FocusRect       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "List of Accounts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   5490
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1853
         TabIndex        =   24
         Top             =   450
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   10650
      TabIndex        =   3
      Top             =   435
      Width           =   10680
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1605
         TabIndex        =   26
         Top             =   135
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker Dtp1 
         Height          =   300
         Left            =   7290
         TabIndex        =   4
         Top             =   105
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         _Version        =   393216
         Format          =   15990784
         CurrentDate     =   40158
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Id"
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
         Left            =   225
         TabIndex        =   27
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Left            =   5565
         TabIndex        =   5
         Top             =   150
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8535
      ScaleWidth      =   10650
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      Begin VB.TextBox txtDebit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   7185
         TabIndex        =   20
         Top             =   6780
         Width           =   1500
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   8700
         TabIndex        =   19
         Top             =   6780
         Width           =   1500
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   -15
         ScaleHeight     =   1380
         ScaleWidth      =   10650
         TabIndex        =   7
         Top             =   7125
         Width           =   10680
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1290
            Left            =   8985
            ScaleHeight     =   1260
            ScaleWidth      =   1545
            TabIndex        =   13
            Top             =   45
            Width           =   1575
            Begin VB.PictureBox PicLast 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   630
               Picture         =   "frmTransaction.frx":0000
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   17
               Top             =   885
               Width           =   345
            End
            Begin VB.PictureBox PicFirst 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   630
               Picture         =   "frmTransaction.frx":0325
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   16
               Top             =   105
               Width           =   345
            End
            Begin VB.PictureBox PicNext 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1080
               Picture         =   "frmTransaction.frx":0635
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   15
               Top             =   495
               Width           =   345
            End
            Begin VB.PictureBox PicPrev 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   180
               Picture         =   "frmTransaction.frx":095B
               ScaleHeight     =   285
               ScaleWidth      =   315
               TabIndex        =   14
               Top             =   495
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   585
            Width           =   1470
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
            TabIndex        =   9
            Top             =   150
            Width           =   1470
         End
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
            TabIndex        =   8
            Top             =   975
            Width           =   3120
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTransaction 
         Height          =   5490
         Left            =   90
         TabIndex        =   6
         Top             =   1230
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   9684
         _Version        =   393216
         Cols            =   6
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
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
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions"
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
         Left            =   4380
         TabIndex        =   2
         Top             =   15
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4410
         TabIndex        =   1
         Top             =   52
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frmTransaction.frx":0C6C
         Stretch         =   -1  'True
         Top             =   15
         Width           =   10665
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Duplicate As Boolean
Dim RowSelect As Integer
Dim rCount As Integer
Dim mSqlQry As String
Dim TransactionId As Single
Dim GLTCode As Single

Private Sub cmdDelete_Click()
    If MsgBox("Do you want to delete this Complete Voucher?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
        If MsgBox("System will be unable to recover the loss data. Continue ?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
'Deleting data
            Con.Execute "Delete from transactions where MainId = " & Val(txtId) & ""
            Con.Execute "Delete from transactionMain where Id = " & Val(txtId) & ""
            MsgBox "Voucher Deleted", vbInformation, "Message"
            cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
Dim GLTCode As Single
Dim FgTransactionRow As Integer

'Checking Data to Save
    If Val(fgTransaction.Rows) < 3 Then
        MsgBox "No data to save", vbInformation, "Message"
        Exit Sub
    End If
    
'Validation for Equal Balances of Debit and Credit Side
    If Not Val(txtDebit) = Val(txtCredit) Then
        MsgBox "Debit and Credit sides should be equal", vbCritical, "Message"
        Exit Sub
    End If
    
    
'If New Transaction
    If cmdNew.Enabled = False Then

    'Getting Maximum Code for GLT
                
        MaxNumber "Gltid", "MaxCode"
        GLTCode = Val(MaxNmbr)
        
    'Inserting Data to TransactionMain Table
            
            Con.Execute "insert into TransactionMain (Id,TDate,Posted) values (" & Val(txtId) & ", '" & Date & "' , '" & "N" & "')"
            
    '----------------------------- Inserting to Transaction Detail -----------------------------
        
            For FgTransactionRow = 1 To fgTransaction.Rows - 2
    
    'Getting Maximum Code for transaction
        MaxNumber "TransId", "MaxCode"
        TransactionId = Val(MaxNmbr)
                    
                Con.Execute "Insert into Transactions (Id, MainId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 5)) & ", '" & fgTransaction.TextMatrix(FgTransactionRow, 2) & "', " & Val(fgTransaction.TextMatrix(FgTransactionRow, 3)) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 4)) & ")"
           
    'Updating MaxCode for Transaction Detail
        UpdateMaxNumber "Transid", Val(TransactionId)
    
            Next
    
    'Updating Transaction Main Id
        UpdateMaxNumber "GltId", Val(txtId)
    
    MsgBox "Record saved", vbInformation, "Message"
    cmdCancel_Click
    
    Else
'If existing record
        GLTCode = Val(txtId)
    
'Deleting old records from TransactionMain and Transaction Detail
        Con.Execute "Delete from transactions where MainId = " & Val(txtId) & ""
        Con.Execute "Delete from transactionMain where Id = " & Val(txtId) & ""
    
'Checking Data to Save
    If Val(fgTransaction.Rows) < 3 Then
        MsgBox "No data to save", vbInformation, "Message"
        Exit Sub
    End If
    
'Validation for Equal Balances of Debit and Credit Side
    If Not Val(txtDebit) = Val(txtCredit) Then
        MsgBox "Debit and Credit sides should be equal", vbCritical, "Message"
        Exit Sub
    End If
        
    'Inserting Data to TransactionMain Table
            
         Con.Execute "insert into TransactionMain (Id,TDate,Posted) values (" & Val(txtId) & ", '" & Date & "' , '" & "N" & "')"
            
    '----------------------------- Inserting to Transaction Detail -----------------------------
        
            For FgTransactionRow = 1 To fgTransaction.Rows - 2
    
    'Getting Maximum Code for transaction
        MaxNumber "TransId", "MaxCode"
        TransactionId = Val(MaxNmbr)
                    
        Con.Execute "Insert into Transactions (Id, MainId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 5)) & ", '" & fgTransaction.TextMatrix(FgTransactionRow, 2) & "', " & Val(fgTransaction.TextMatrix(FgTransactionRow, 3)) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 4)) & ")"
           
    'Updating MaxCode for Transaction Detail
        UpdateMaxNumber "Transid", Val(TransactionId)
    
            Next
    
    MsgBox "Record updated", vbInformation, "Message"
    cmdCancel_Click
    
    
    End If
End Sub

Private Sub Dtp1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fgTransaction.SetFocus
    End If
End Sub

Private Sub fgTransaction_Click()
    RowSelect = fgTransaction.RowSel
End Sub

Private Sub fgTransaction_KeyDown(KeyCode As Integer, Shift As Integer)
'If CTRL + SPACE is pressed
    If fgTransaction.Col = 1 Then
        If KeyCode = 32 And Shift = 2 Then
            PicSrchGrid.Visible = True
            TxtGrdSrch.Text = ""

            FillGridAccounts
            SetGridAccounts
            
            MshSearch.Col = 0
            MshSearch.Row = 1
            MshSearch.SetFocus
        End If
    End If

'Delete Row from Grid
    If KeyCode = vbKeyDelete Then
        
        If fgTransaction.Rows > 2 And fgTransaction.TextMatrix(fgTransaction.Row, 1) <> "" Then
            If MsgBox("Do you want to delete this line>", vbQuestion + vbYesNo, "Delete Line") = vbYes Then
                txtDebit.Text = Val(txtDebit) - Val(fgTransaction.TextMatrix(fgTransaction.Row, 3))
                txtCredit.Text = Val(txtCredit) - Val(fgTransaction.TextMatrix(fgTransaction.Row, 4))
                        
                fgTransaction.RemoveItem RowSelect
            End If
            
        Else
            MsgBox "Blank or Last line can not be deleted", vbCritical, "Message "
            Exit Sub
        End If
        
    End If
End Sub

Private Sub fgTransaction_KeyPress(KeyAscii As Integer)
    EditGrid fgTransaction, KeyAscii
End Sub

Private Sub Form_Load()

'Setting up flexgrid data
    Call GridSetting

'Calling Exist Data
    Call ExistData

 'Calling Data for navigation
    Set RsNAV = New ADODB.Recordset
    If RsNAV.State = 1 Then RsNAV.Close
    RsNAV.Open "Select * from TransactionMain Order By Id", Con, adOpenStatic, adLockOptimistic
End Sub

Private Sub Form_Resize()
    Me.Left = 3600
    Me.Top = 1400


End Sub


Public Sub GridSetting()
'Setting of Transaction Grid
    With fgTransaction
        
        .ColWidth(0) = 250
        .ColWidth(1) = 2500
        .ColWidth(2) = 4400
        .ColWidth(3) = 1475
        .ColWidth(4) = 1475
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    
        .RowHeight(0) = 400
    
        .Rows = 2
        
        .TextMatrix(0, 1) = "Account Name"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Debit"
        .TextMatrix(0, 4) = "Credit"
        
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
    
    End With
End Sub

Public Sub SetGridAccounts()
'Setting of Search Grid
    With MshSearch
        .ColWidth(1) = 2500
        .ColWidth(2) = 1500
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Account Title"
        .TextMatrix(0, 2) = "Account Type"
        
        .RowHeight(0) = 400
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
    
    End With
    
End Sub

Public Sub FillGridAccounts()
'Filling all Accounts Data in Search grid
    SQLQRY = "Select Id, Title, HeadTitle from ViewHeadWise"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open SQLQRY, Con, adOpenStatic, adLockReadOnly
            
        Set MshSearch.DataSource = RS
End Sub

Private Sub TxtGrdSrch_Change()
    Call SearchRecord
End Sub
Public Sub SearchRecord()
'Filling the Search grid with Critarial Data
    
    Dim SearchedRowCount As Integer
    
    If PicSrchGrid.Visible = True Then
        
        MshSearch.Rows = 2
        MshSearch.Row = 0
        
        SearchedRowCount = 0
        
        SQLQRY = "Select Id, Title, HeadTitle from ViewHeadWise where Title Like '" & TxtGrdSrch.Text & "%'"
            
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
            RS.Open SQLQRY, Con, adOpenStatic, adLockReadOnly
                
                If RS.RecordCount <= 0 Then
                    MshSearch.Clear
                    
                    With MshSearch
                        .TextMatrix(0, 0) = "ID"
                        .TextMatrix(0, 1) = "Account Title"
                        .TextMatrix(0, 2) = "Account Type"
                    End With
                    
                    
                    MshSearch.Rows = 2
                    Exit Sub
                End If
                
                For SearchedRowCount = 1 To RS.RecordCount
                    
                    If SearchedRowCount >= MshSearch.Rows - 1 Then
                        MshSearch.Rows = MshSearch.Rows + 1
                        MshSearch.Row = MshSearch.Row + 1
                    End If
                    
                    MshSearch.TextMatrix(SearchedRowCount, 0) = RS(0)
                    MshSearch.TextMatrix(SearchedRowCount, 1) = RS(1)
                    MshSearch.TextMatrix(SearchedRowCount, 2) = RS(2)
                    
                    
                    RS.MoveNext
                Next
                
                MshSearch.Col = 0
                MshSearch.Row = 0
                MshSearch.ColAlignment(0) = 3
    End If
    
End Sub

Private Sub MshSearch_DblClick()
    If MshSearch.Row = 0 Then
        Exit Sub
    End If
    
'Checking for duplicate entry in grid
        Call CheckDuplicate
        
        If Duplicate = True Then
            Exit Sub
        End If
        
        fgTransaction.TextMatrix(fgTransaction.Row, 1) = MshSearch.TextMatrix(MshSearch.Row, 1)
        fgTransaction.TextMatrix(fgTransaction.Row, 5) = MshSearch.TextMatrix(MshSearch.Row, 0)
        
        PicSrchGrid.Visible = False
        fgTransaction.Col = 1
        fgTransaction.SetFocus
End Sub

Public Sub CheckDuplicate()
Dim dRow As Integer
   
    For dRow = 1 To fgTransaction.Rows - 2
        If fgTransaction.TextMatrix(dRow, 5) = MshSearch.TextMatrix(MshSearch.Row, 0) Then
            MsgBox "Account Title aleready selected", vbInformation, "Message"
            Duplicate = True
            Exit Sub
        Else
            Duplicate = False
        End If
    Next
    
        
End Sub

Private Sub MshSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MshSearch_DblClick
    End If
    
    If KeyAscii = 27 Then
        PicSrchGrid.Visible = False
        fgTransaction.SetFocus
        Exit Sub
    End If
    
    If KeyAscii = 8 Then
        If TxtGrdSrch.Text <> "" Then TxtGrdSrch.Text = Left$(TxtGrdSrch.Text, (Len(TxtGrdSrch.Text) - 1))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    Else
        TxtGrdSrch.Text = TxtGrdSrch.Text + Chr$(KeyAscii)
    End If
    
End Sub

Private Sub TxtGrdSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        PicSrchGrid.Visible = False
        fgTransaction.SetFocus
        Exit Sub
    End If
End Sub

Public Sub ExistData()

'Getting data from Purchase
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
        RS.Open "SELECT Id, TDate from TransactionMain", Con, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                Exit Sub
            End If
            
            txtId.Text = Val(RS(0))
            Dtp1.Value = RS(1)
        RS.Close
            
'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        mSqlQry = "SELECT Accounts.Title, Transactions.Descript, Transactions.DrAmount, Transactions.CrAmount, Accounts.Id, TransactionMain.ID FROM (TransactionMain INNER JOIN Transactions ON TransactionMain.ID = Transactions.MainId) INNER JOIN Accounts ON Transactions.AcId = Accounts.Id Where TransactionMain.ID = " & Val(txtId) & ""
        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgTransaction.Rows = 2
            fgTransaction.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgTransaction.TextMatrix(fgTransaction.Row, 1) = RS(0)
                fgTransaction.TextMatrix(fgTransaction.Row, 2) = RS(1)
                fgTransaction.TextMatrix(fgTransaction.Row, 3) = RS(2)
                fgTransaction.TextMatrix(fgTransaction.Row, 4) = RS(3)
                fgTransaction.TextMatrix(fgTransaction.Row, 5) = RS(4)
                
                txtDebit.Text = Val(txtDebit) + Val(RS(2))
                txtCredit.Text = Val(txtCredit) + Val(RS(3))
                
                fgTransaction.Rows = fgTransaction.Rows + 1
                fgTransaction.Row = fgTransaction.Row + 1
                                
                RS.MoveNext
            Next

End Sub

Private Sub PicFirst_Click()
    On Error Resume Next
    RsNAV.MoveFirst
    LblNav.Caption = "1"
    
    If RsNAV.BOF = True Then
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
        Exit Sub
    Else
        Call NAVData
    End If
End Sub

Private Sub PicLast_Click()
    On Error Resume Next
    RsNAV.MoveLast
    LblNav.Caption = Val(RsNAV.RecordCount)
    
    If RsNAV.EOF = True Then
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call NAVData
    End If
    
End Sub

Private Sub PicNext_Click()
    On Error Resume Next
    RsNAV.MoveNext

    LblNav.Caption = Val(LblNav) + 1

    If RsNAV.EOF = True Then
        LblNav.Caption = Val(LblNav) - 1
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call NAVData
    End If
End Sub

Private Sub PicPrev_Click()
    On Error Resume Next
    RsNAV.MovePrevious
    
    LblNav.Caption = Val(LblNav) - 1
    
    
    If RsNAV.BOF = True Then
        LblNav.Caption = Val(LblNav) + 1
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
    Else
        Call NAVData
    End If
    
End Sub

Public Sub NAVData()

'Getting data from Purchase
            
            txtId.Text = Val(RsNAV(0))
            Dtp1.Value = RsNAV(1)
            
'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        mSqlQry = "SELECT Accounts.Title, Transactions.Descript, Transactions.DrAmount, Transactions.CrAmount, Accounts.Id, TransactionMain.ID FROM (TransactionMain INNER JOIN Transactions ON TransactionMain.ID = Transactions.MainId) INNER JOIN Accounts ON Transactions.AcId = Accounts.Id Where TransactionMain.ID = " & Val(txtId) & ""
        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgTransaction.Rows = 2
            fgTransaction.Row = 1
            
                txtDebit.Text = ""
                txtCredit.Text = ""
            
            For rCount = 1 To RS.RecordCount
                fgTransaction.TextMatrix(fgTransaction.Row, 1) = RS(0)
                fgTransaction.TextMatrix(fgTransaction.Row, 2) = RS(1)
                fgTransaction.TextMatrix(fgTransaction.Row, 3) = RS(2)
                fgTransaction.TextMatrix(fgTransaction.Row, 4) = RS(3)
                fgTransaction.TextMatrix(fgTransaction.Row, 5) = RS(4)
                
                txtDebit.Text = Val(txtDebit) + Val(RS(2))
                txtCredit.Text = Val(txtCredit) + Val(RS(3))
                
                fgTransaction.Rows = fgTransaction.Rows + 1
                fgTransaction.Row = fgTransaction.Row + 1
                                
                RS.MoveNext
            Next
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
'Clearing all controls for New Data
    Clear Me
    fgTransaction.Clear
    
    GridSetting
    
    Call AutoId
    
    Dtp1.SetFocus
    Dtp1.Value = Now
            
    Modes True, False, Me

'Lock Navigation
    Picture4.Enabled = False

End Sub
Private Sub cmdCancel_Click()
    Clear Me
    
    fgTransaction.Clear
    
    GridSetting
            
    Call ExistData
    
    Modes False, True, Me
    
    Dtp1.SetFocus
    
    If RsNAV.RecordCount <= 0 Then
        Exit Sub
    Else
        RsNAV.Requery
        RsNAV.MoveFirst
    End If

    LblNav.Caption = "1"

'UnLock Navigation
    Picture4.Enabled = True

End Sub

Public Sub AutoId()
'Calling MaxNumber function to get Auto Id for the record
    MaxNumber "GLTID", "MaxCode"
    txtId.Text = Val(MaxNmbr)
End Sub
