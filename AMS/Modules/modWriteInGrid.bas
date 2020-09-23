Attribute VB_Name = "modWriteInGrid"
Public Function EditGrid(Msh As MSHFlexGrid, KeyAscii As Integer)
        
'==================================== Editing Grid Code For A/c Opening ====================================
        
'If First Row selected then
    Select Case Msh.Row
        Case 0
            KeyAscii = 0
            Exit Function
    End Select
    
'Block Every Key Except Enter in Column 0 (Account Name)

        Select Case Msh.Col

            Case 1
            
'If ENTER is pressed on Account Name Column then move to next column
            If KeyAscii = 13 Then
                SendKeys "{right}"
                Exit Function
            End If

                KeyAscii = 0
                Exit Function
        End Select

'---------------------------------------------------------
        
        Select Case KeyAscii
            
            Case 8: 'IF KEY IS BACKSPACE THEN
                If Msh.Text <> "" Then Msh.Text = Left$(Msh.Text, (Len(Msh.Text) - 1))
            
            Case 13: 'IF KEY IS ENTER THEN
                
                Select Case Msh.Col
                    
                    Case Is < 4
                        '----------Move Curssor to Right Side untill col >= 3------
                        Msh.SetFocus
                        SendKeys "{right}"
                    
                    Case 4
                        Msh.SetFocus
                        If (Msh.Row + 1) = Msh.Rows Then
                            ''-------- Null Value Chk in last col---------------
                            
'If values are missing in Debit and Credit Side
                            If Len(Msh.TextMatrix(Msh.Row, 1)) = 0 Or Msh.TextMatrix(Msh.Row, 1) = "" Then
                                 MsgBox "Select Account Name", vbInformation, "Message"
                                 Msh.Col = 1
                                 Exit Function
                            End If
                            
'If values are missing in Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 3)) = 0 And Val(Msh.TextMatrix(Msh.Row, 4)) = 0 Then
                                 MsgBox "Enter Debit / Credit Amount", vbInformation, "Message"
                                 Msh.Col = 3
                                 Exit Function
                            End If
                            
'If values are in both Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 3)) > 0 And Val(Msh.TextMatrix(Msh.Row, 4)) > 0 Then
                                 MsgBox "Enter Only Debit OR Credit Amount", vbInformation, "Message"
                                 Msh.Col = 3
                                 Exit Function
                            End If
                            
                            ''-------------New Row Create----------
                            Msh.Rows = Msh.Rows + 1
                            Msh.Col = 1
                            ''-------------------------------------
                            
                            ''-------------Adding Amount to Debit / credit Textbox----------
                            frmTransaction.txtCredit = Val(frmTransaction.txtCredit) + Val(Msh.TextMatrix(Msh.Row, 4))
                            frmTransaction.txtDebit = Val(frmTransaction.txtDebit) + Val(Msh.TextMatrix(Msh.Row, 3))
                            
                            ''-------------------------------------
                            
                            
                        End If
                        
                        SendKeys "{home}" + "{down}"   '' + "{right}"
                
                End Select
            
            Case Else 'KeyAscii Select
                ''-------write any code for Any Validation-------------
                Select Case Msh.Col
                    Case 3, 4
                        ''-------Allow Number Validation-------------
                        ONUGrid KeyAscii, Msh
                End Select
                ''-------------Write Data in Cells----------
                Msh.Text = Msh.Text + Chr$(KeyAscii)
                ''-------------------------------------
            End Select 'KeyAscii Select
End Function
'Only numbers in Grid
Public Function ONUGrid(KeyAscii As Integer, txt As MSHFlexGrid)
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 46 Then


    If KeyAscii = 46 Then


        If InStr(txt.Text, ".") Then
            KeyAscii = 0
            Exit Function
        Else
            txt.Text = txt.Text
        End If
    Else
    End If
Else
    KeyAscii = 0
End If
End Function
