VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CR_Catchup_Form_addin 
   Caption         =   "CR Catch up"
   ClientHeight    =   3624
   ClientLeft      =   -228
   ClientTop       =   -864
   ClientWidth     =   7716
   OleObjectBlob   =   "CR_Catchup_Form_addin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CR_Catchup_Form_addin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Budget = ""
If IsNumeric(RefEdit1.text) Then
    Budget = Val(RefEdit1.text)
Else
    If Range(RefEdit1.text).Count = 1 Then
        Budget = Range(RefEdit1.text).Value
    Else
        Budget = Application.WorksheetFunction.Sum(Range(RefEdit1.text))
    End If
End If

Budget = Round(Budget, 2)
Debug.Print "Budget: " & Budget
Application.ScreenUpdating = False

Set CRatualRange = Range(RefEdit2.text)
If CRatualRange.Count = 1 Then
    Set CRatualRange = Range(CRatualRange, CRatualRange.End(xlDown))
End If
Set MidpointRange = Range(RefEdit3.text)
ColumnCR = CRatualRange.Column
ColumnSpend = Range(RefEdit6.text).Column

AllEligible = True
If RefEdit5.text <> "" Then
    ColumnReceivership = Range(RefEdit5.text).Column
    AllEligible = False
End If
    
AllFT = True
If RefEdit4.text <> "" Then
    Set FTERange = Range(RefEdit4.text)
    AllFT = False
End If

enough = False

Errors = Application.WorksheetFunction.CountIf(CRatualRange, "#DIV/0!") + Application.WorksheetFunction.CountIf(CRatualRange, "#N/A")
If Errors > 0 Then
    MsgBox "There is an error among your CR data. Please fix that and try again."
Else
    MinCR = Application.WorksheetFunction.Min(CRatualRange)
    
    
    CR_Goal = MinCR
    
    Incremento = 0.01
    For a = 1 To 15
        While BudgetNecessario < Budget And enough = False
            BudgetNecessario = 0
            For Each celula In CRatualRange
                If AllFT = True Then
                    FTE = 1
                Else
                    FTE = Cells(celula.Row, FTERange.Column).Value
                End If
                
                If CR_Goal > Cells(celula.Row, ColumnCR).Value Then
                    If AllEligible = False Then
                        If Cells(celula.Row, ColumnReceivership).text = "No" Then
                            Gasto = 0
                        Else
                            Gasto = Round((CR_Goal - Cells(celula.Row, CRatualRange.Column).Value) * Cells(celula.Row, MidpointRange.Column).Value * FTE, 2)
                            BudgetNecessario = BudgetNecessario + Gasto
                        End If
                    Else
                        Gasto = Round((CR_Goal - Cells(celula.Row, CRatualRange.Column).Value) * Cells(celula.Row, MidpointRange.Column).Value * FTE, 2)
                        BudgetNecessario = BudgetNecessario + Gasto
                    End If
                Else
                    Gasto = 0
                End If
            Next
            DoEvents
            Debug.Print BudgetNecessario & " for CR = " & CR_Goal
            If a = 15 Then
                b = b + 1
                If b > 10 Then
                    enough = True
                End If
            End If
            If Round(BudgetNecessario, 2) = Budget Then
                CR_Goal_Final = CR_Goal
                Debug.Print "Final CR goal is: " & CR_Goal_Final
            End If
            CR_Goal = CR_Goal + Incremento
        Wend
        
        CR_Goal = CR_Goal - 2 * Incremento
        BudgetNecessario = 0
        Debug.Print "switching to " & Incremento & " increments"
        Incremento = Incremento / 10
    Next
    
    
    
    Incremento = Incremento * 10
    CR_Goal = CR_Goal - Incremento
    
    If CR_Goal_Final > 0 Then
        CR_Goal = CR_Goal_Final
    End If
    
    UltimoCara = 0
    
    For Each celula In CRatualRange
        If AllFT = True Then
            FTE = 1
        Else
            FTE = Cells(celula.Row, FTERange.Column).Value
        End If
        If CR_Goal > Cells(celula.Row, ColumnCR).Value Then
            If AllEligible = False Then
                If Cells(celula.Row, ColumnReceivership).text = "No" Then
                    Gasto = 0
                Else
                    Gasto = Round((CR_Goal - Cells(celula.Row, CRatualRange.Column).Value) * Cells(celula.Row, MidpointRange.Column).Value * FTE, 2)
                    BudgetNecessario = BudgetNecessario + Gasto
                    UltimoCara = celula.Row
                End If
            Else
                Gasto = Round((CR_Goal - Cells(celula.Row, CRatualRange.Column).Value) * Cells(celula.Row, MidpointRange.Column).Value * FTE, 2)
                BudgetNecessario = BudgetNecessario + Gasto
                UltimoCara = celula.Row
            End If
        Else
            Gasto = 0
        End If
        If IsNumeric(Cells(celula.Row, ColumnSpend).Value) Then
            Cells(celula.Row, ColumnSpend).FormulaR1C1 = Gasto
        End If
    Next
    
    'Ajuste:
    If BudgetNecessario <> Budget Then
        Cells(UltimoCara, ColumnSpend).FormulaR1C1 = Cells(UltimoCara, ColumnSpend).Value + (Budget - BudgetNecessario)
        Debug.Print "Adjusting row " & UltimoCara & " by " & (Budget - BudgetNecessario)
    End If
    
End If

CR_Catchup_Form_addin.Hide
Application.ScreenUpdating = True

End Sub


Private Sub RefEdit1_Change()
If Not (IsNumeric(RefEdit1.text)) Then
    RefEdit1.text = Range(RefEdit1.text).Address
End If
End Sub

Private Sub RefEdit2_Change()
If Not (IsNumeric(RefEdit2.text)) Then
    RefEdit2.text = Range(RefEdit2.text).Address
End If
End Sub

Private Sub RefEdit3_Change()
If Not (IsNumeric(RefEdit3.text)) Then
    RefEdit3.text = Range(RefEdit3.text).Address
End If
End Sub


Private Sub RefEdit4_Change()
If Not (IsNumeric(RefEdit4.text)) Then
    RefEdit4.text = Range(RefEdit4.text).Address
End If
End Sub

Private Sub RefEdit5_Change()
If Not (IsNumeric(RefEdit5.text)) Then
    RefEdit5.text = Range(RefEdit5.text).Address
End If
End Sub
Private Sub RefEdit6_Change()
If Not (IsNumeric(RefEdit6.text)) Then
    RefEdit6.text = Range(RefEdit6.text).Address
End If
End Sub

