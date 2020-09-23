VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customized Accounts Management System"
   ClientHeight    =   9255
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMainFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10470
      Left            =   2625
      ScaleHeight     =   10440
      ScaleWidth      =   12600
      TabIndex        =   13
      Top             =   120
      Width           =   12630
      Begin VB.Image Image1 
         Height          =   2775
         Left            =   -60
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   12660
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   90
      ScaleHeight     =   8235
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   1245
      Width           =   2430
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   285
         TabIndex        =   12
         Top             =   7050
         Width           =   1800
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Income Statmnt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   6000
         Width           =   1890
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Sheet"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   285
         TabIndex        =   10
         Top             =   4965
         Width           =   1800
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Balance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   285
         TabIndex        =   9
         Top             =   3915
         Width           =   1800
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   285
         TabIndex        =   8
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Accounts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   285
         TabIndex        =   7
         Top             =   1845
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Top             =   6945
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   5910
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   4860
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   3825
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   2775
         Width           =   2070
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F9F9&
         Height          =   450
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   1740
         Width           =   2070
      End
      Begin VB.Image Image2 
         Height          =   870
         Left            =   -15
         Picture         =   "frmMain.frx":F6BF
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   2430
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   8250
      Left            =   165
      Top             =   1305
      Width           =   2400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
                Dim mform As Form
                    For Each mform In Forms
                        Unload mform
                    Next
End Sub

Private Sub Label2_Click(Index As Integer)
    Select Case Index
        Case 0
            lblCaption_Click (0)
        Case 1
            lblCaption_Click (1)
        Case 2
            lblCaption_Click (2)
        Case 3
            lblCaption_Click (3)
        Case 4
            lblCaption_Click (4)
        Case 5
            lblCaption_Click (5)
            
    End Select
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            OnColorChange Label2(0), lblCaption(0)

        Case 1
            OnColorChange Label2(1), lblCaption(1)
            
        Case 2
            OnColorChange Label2(2), lblCaption(2)
            
        Case 3
            OnColorChange Label2(3), lblCaption(3)
            
        Case 4
            OnColorChange Label2(4), lblCaption(4)
            
        Case 5
            OnColorChange Label2(5), lblCaption(5)
            
    End Select

End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblCaption_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAccounts.Show vbModal
        
        Case 1
            frmTransaction.Show vbModal
        
        Case 2
            Call TrialBalanceData
'Refreshing RepTrialBalance Table
            Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open "Select * from RepTrialBalance", Con, adOpenStatic, adLockOptimistic
                    RS.Requery
'Showing Report
            DataRepTrialBalance.Show
        
        Case 3
            Call BalanceSheetData
        
        Case 4
            Call ShortIncomeStatement
        
        Case 5
            If MsgBox("Do you want to exit ?", vbYesNo + vbQuestion, "Exit") = vbYes Then
                Unload Me
            End If
    
    End Select
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            OnColorChange Label2(0), lblCaption(0)

        Case 1
            OnColorChange Label2(1), lblCaption(1)
            
        Case 2
            OnColorChange Label2(2), lblCaption(2)
            
        Case 3
            OnColorChange Label2(3), lblCaption(3)
            
        Case 4
            OnColorChange Label2(4), lblCaption(4)
            
        Case 5
            OnColorChange Label2(5), lblCaption(5)
            
    End Select


End Sub



Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Integer

    For a = 0 To 5
        Label2(a).BackColor = &HF9F9F9
    Next
    
    For a = 0 To 5
        lblCaption(a).ForeColor = vbBlack
    Next
    
End Sub

Private Sub picMainFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Integer

    For a = 0 To 5
        If Label2(a).BackColor = vbBlack Then
            Label2(a).BackColor = &HF9F9F9
        End If
    Next
    
    For a = 0 To 5
        If lblCaption(a).ForeColor = vbWhite Then
            lblCaption(a).ForeColor = vbBlack
        End If
    Next
    
End Sub

Public Sub TrialBalanceData()
Dim AcId As Single
Dim AcTitle As String
Dim SumofDR As Single
Dim SumofCr As Single
Dim Balance As Single
Dim DrBal As Single
Dim CrBal As Single

    Con.Execute "Delete * from RepTrialBalance"
    
'Copying data from ViewTrialBalance to RepTrialBalance
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select * from ViewTrialBalance", Con, adOpenStatic, adLockOptimistic
            While Not RS.EOF = True
                AcId = Val(RS(0))
                AcTitle = RS(1)
                SumofDR = Val(RS(2))
                SumofCr = Val(RS(3))
                Balance = Val(RS(2)) - Val(RS(3))
                
                If Balance > 0 Then
                    DrBal = Abs(Balance)
                    CrBal = 0
                ElseIf Balance < 0 Then
                    CrBal = Abs(Balance)
                    DrBal = 0
                Else
                    DrBal = 0
                    CrBal = 0
                End If
                
                Con.Execute "Insert into RepTrialBalance(AcId,Title,Dr,Cr) Values (" & Val(AcId) & ", '" & AcTitle & "', " & Val(DrBal) & ", " & Val(CrBal) & ")"
                RS.MoveNext
                
            Wend
            
            
'Getting Total of both (Debit / Credit Side)
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(Dr),Sum(Cr) from RepTrialBalance", Con, adOpenStatic, adLockOptimistic
            
            On Error Resume Next
            
            TrialDr = Val(RS(0))
            TrialCr = Val(RS(1))
            

'=============================================REPORT===============================================
            
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
    RS.Open "Select * from RepTrialBalance", Con, adOpenStatic, adLockOptimistic
        
        With DataRepTrialBalance
            Set .DataSource = RS
            .DataMember = RS.DataMember
            
            .Sections("Section1").Controls("Text1").DataField = "AcId"
            .Sections("Section1").Controls("Text2").DataField = "Title"
            .Sections("Section1").Controls("Text3").DataField = "DR"
            .Sections("Section1").Controls("Text4").DataField = "CR"
        End With
            
            

End Sub

Public Sub BalanceSheetData()
'Deleteing old Data from RepBalanceSheet
    Con.Execute "Delete * from RepBalanceSheet"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
            
'Copying data from ViewBalancesheet to RepBalanceSheet
        RS.Open "Select * from ViewBalanceSheet", Con, adOpenStatic, adLockPessimistic
            If RS.RecordCount > 0 Then
                While Not RS.EOF = True
                    Con.Execute "Insert into RepBalanceSheet Values(" & Val(RS(0)) & ", '" & RS(1) & "', " & Val(RS(2)) & ", '" & RS(3) & "', " & IIf(Val(RS(6)) > 0, Val(Abs(RS(6))), 0) & ", " & IIf(Val(RS(6)) < 0, Val(Abs(RS(6))), 0) & ")"
                    RS.MoveNext
                Wend
            End If
           
'=============================================REPORT===============================================
    Dim ProfitLoss As Single
    Dim Profit As Single
    Dim Loss As Single
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
    RS.Open "Select * from RepBalanceSheet where HeadId < 4", Con, adOpenStatic, adLockOptimistic
        
        With DataRepBalanceSheet
            Set .DataSource = RS
            .DataMember = RS.DataMember
            
            .Sections("Section1").Controls("txtId").DataField = "AcId"
            .Sections("Section1").Controls("txtTitle").DataField = "AcTitle"
            .Sections("Section1").Controls("txtDr").DataField = "DR"
            .Sections("Section1").Controls("txtCr").DataField = "CR"
          
'Calculating Total Assets and Liabilities
            Set RsMisc = New ADODB.Recordset
            If RsMisc.State = 1 Then RsMisc.Close
                RsMisc.Open "Select Sum(Dr),Sum(Cr) from RepBalanceSheet where HeadId < 4", Con, adOpenStatic, adLockOptimistic
                    
'Calculating Profit / Loss
                    On Error Resume Next
                    
                    ProfitLoss = Val(RsMisc(0)) - Val(RsMisc(1))
                    
'If difference (Assets) > Liabilities which means "PROFIT"
                    If Val(ProfitLoss) > 0 Then
                        Profit = Abs(Val(ProfitLoss))
                        Loss = 0
                        .Sections("Section5").Controls("lblprofit").Caption = Abs(Val(Profit))
                        .Sections("Section5").Controls("lblProLoss").Caption = "Net Profit"
                                                                                        
'If difference (Assets) < Liabilities which means "LOSS"
                    ElseIf Val(ProfitLoss) < 0 Then
                        Profit = 0
                        Loss = Abs(Val(ProfitLoss))
                        
                        .Sections("Section5").Controls("lblLoss").Caption = Abs(Val(Loss))
                        .Sections("Section5").Controls("lblProLoss").Caption = "Net Loss"
                    Else
                        .Sections("Section5").Controls("lblprofit").Caption = 0
                        .Sections("Section5").Controls("lblLoss").Caption = 0
                    End If
                    
'Totals of both Sides
                    .Sections("Section5").Controls("lblTotalDr").Caption = Val(RsMisc(0))
                    .Sections("Section5").Controls("lblTotalCr").Caption = Val(RsMisc(1))
       
'Final Total
                    .Sections("Section5").Controls("lblTotalLiabilities").Caption = Val(RsMisc(1)) + Abs(Val(Profit))
                    .Sections("Section5").Controls("lblTotalAssets").Caption = Val(RsMisc(0)) + Abs(Val(Loss))
                                     
                    On Error Resume Next
            .Show
        End With
            
    
    End Sub

Public Sub ShortIncomeStatement()
Dim TotalRev As Single
Dim TotalExp As Single
Dim ProfLoss As Single


'Getting Revenue
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Balance from ViewIncomeStatement where Id = 4", Con, adOpenStatic, adLockPessimistic
            TotalRev = Abs(Val(RS(0)))
        
'Getting Expenses
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Balance from ViewIncomeStatement where Id = 5", Con, adOpenStatic, adLockPessimistic
            TotalExp = Abs(Val(RS(0)))

'Getting Difference (Final Profit / Loss)
            ProfLoss = Val(TotalRev) - Val(TotalExp)
    
'=============================================REPORT===============================================
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
    RS.Open "Select * from ViewIncomeStatement", Con, adOpenStatic, adLockOptimistic
        
        With DataRepIncomeStatement
            Set .DataSource = RS
            .DataMember = RS.DataMember
        
            .Sections("Section2").Controls("lblTotalRev").Caption = Val(TotalRev)
            .Sections("Section2").Controls("lblTotalExp").Caption = Val(TotalExp)
            .Sections("Section2").Controls("lblTotalProLoss").Caption = Abs(Val(ProfLoss))
            
            If Val(ProfLoss) > 0 Then
                .Sections("Section2").Controls("lblProfitLoss").Caption = "Net Profit :"
            ElseIf Val(ProfLoss) < 0 Then
                .Sections("Section2").Controls("lblProfitLoss").Caption = "Net Loss :"
            Else
                .Sections("Section2").Controls("lblProfitLoss").Caption = "No Profit / Loss"
            End If
                
            
            .Show
        End With
End Sub
