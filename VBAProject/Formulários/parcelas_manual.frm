VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} parcelas_manual 
   Caption         =   "Preenchimento das Parcelas"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "parcelas_manual.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "parcelas_manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub anterior_botton_Click()
    On Error GoTo errodata
    If CDate(data.Value) Then
    End If
    If parcela.Value - 1 >= 1 And parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        ElseIf CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value - 1 >= 1 Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    End If
    
    On Error GoTo errodin
    If CCur(valor.Value) Then
    End If
    
    If valor.Value < 0 Then
        MsgBox ("Valor gasto deve ser um número positivo.")
        valor.SetFocus
        Exit Sub
    End If
    
    Sheets("Gastos").Range("R" & parcela.Value).Value = CDate(data.Value)
    Sheets("Gastos").Range("s" & parcela.Value).Value = CCur(valor.Value)

    parcela.Value = parcela.Value - 1
    Exit Sub
errodata:
    MsgBox ("Data inválida.")
    data.SetFocus
    Exit Sub
errodin:
    MsgBox ("Valor inválido.")
    valor.SetFocus
    Exit Sub
End Sub

Sub concluir_parcelas_Click()
    On Error GoTo errodin
    If CCur(valor.Value) Then
    End If
    
    If valor.Value < 0 Then
        GoTo errodin
    End If
    
    On Error GoTo errodata
    If CDate(data.Value) Then
    End If
    If parcela.Value - 1 >= 1 And parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        ElseIf CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value - 1 >= 1 Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    End If
    
    Sheets("Gastos").Range("R" & parcela.Value).Value = CDate(data.Value)
    Sheets("Gastos").Range("S" & parcela.Value).Value = CCur(valor.Value)

    ultima_linha_parcela = Sheets("Gastos").Range("Q1000000").End(xlUp).Row
    For i = 1 To ultima_linha_parcela
        If Sheets("Gastos").Range("S" & i).Value = "" Then
            MsgBox ("Para concluir, é necessário preencher todas parcelas.")
            Exit Sub
        End If
    Next i
    
    Sheets("Gastos").Range("N1").FormulaLocal = "=SOMA(S1:S" & ultima_linha_parcela & ")"
    soma_parcelas = Sheets("Gastos").Range("N1").Value
    Sheets("Gastos").Range("N1").Clear
    If CDec(soma_parcelas) <> CDec(trab.valor_gasto.Value) Then
        MsgBox ("Soma dos valores das parcelas preenchidas não condiz com o valor total gasto na compra.")
        Exit Sub
    End If
    
    MsgBox ("Parcelas preenchidas com sucesso.")
    Call Unload(parcelas_manual)
    Exit Sub
errodin:
    MsgBox ("Preencha corretamente o valor da parcela.")
    valor.SetFocus
    Exit Sub
errodata:
    MsgBox ("Preencha corretamente a data da parcela no formato ""dd/mm/aaaa"".")
    data.SetFocus
    Exit Sub
End Sub

Private Sub data_Change()
    Call trab.data(data)
End Sub

Private Sub parcela_Change()
    data.Enabled = True
    If parcela.Value = 1 Then
        data.Enabled = False
        valor.SetFocus
        anterior_botton.Enabled = False
        proxima_botton.Enabled = True
    ElseIf parcela.Value = trab.parcelas_gasto.Value Then
        data.SetFocus
        proxima_botton.Enabled = False
        anterior_botton.Enabled = True
    Else
        data.SetFocus
        proxima_botton.Enabled = True
        anterior_botton.Enabled = True
    End If
    
    data.Value = CDate(Sheets("GAstos").Range("R" & parcela.Value).Value)
    valor.Value = CCur(Sheets("GAstos").Range("S" & parcela.Value).Value)
    
End Sub

Private Sub proxima_botton_Click()
    On Error GoTo errodata
    If CDate(data.Value) Then
    End If
    If parcela.Value - 1 >= 1 And parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        ElseIf CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value - 1 >= 1 Then
        If CDate(data.Value) < CDate(Sheets("Gastos").Range("R" & parcela.Value - 1).Value) Then
            MsgBox ("Necessário colocar uma data maior que a data da parcela anterior.")
            data.SetFocus
            Exit Sub
        End If
    ElseIf parcela.Value + 1 <= Int(trab.parcelas_gasto.Value) Then
        If CDate(data.Value) > CDate(Sheets("Gastos").Range("R" & parcela.Value + 1).Value) Then
            MsgBox ("Necessário colocar uma data menor que a data da parcela posterior.")
            data.SetFocus
            Exit Sub
        End If
    End If
    
    On Error GoTo errodin
    If CCur(valor.Value) Then
    End If
    
    If valor.Value < 0 Then
        MsgBox ("Valor gasto deve ser um número positivo.")
        valor.SetFocus
        Exit Sub
    End If
    
    Sheets("Gastos").Range("R" & parcela.Value).Value = CDate(data.Value)
    Sheets("Gastos").Range("s" & parcela.Value).Value = CCur(valor.Value)

    parcela.Value = parcela.Value + 1
    Exit Sub
errodata:
    MsgBox ("Data inválida.")
    data.SetFocus
    Exit Sub
errodin:
    MsgBox ("Valor inválido.")
    valor.SetFocus
    Exit Sub
End Sub

Sub UserForm_Initialize()
    data.Value = CDate(trab.data_gasto.Value)
    Sheets("Gastos").Range("T1").Value = CDate(trab.data_gasto.Value)
    Sheets("Gastos").Range("R1").Value = CDate(trab.data_gasto.Value)
    
    ultima = Sheets("Gastos").Range("Q1000000").End(xlUp).Row
    If ultima = 1 Then
        For i = 1 To trab.parcelas_gasto.Value
            Sheets("Gastos").Range("Q" & i).Value = i
            Sheets("Gastos").Range("S" & i).Value = ""
        Next i
        Sheets("Gastos").Range("T1").Value = CDate(trab.data_gasto.Value)
        Sheets("Gastos").Range("R1").FormulaLocal = "=DATAM($T$1;Q1-1)"
        Sheets("Gastos").Range("R1:R" & trab.parcelas_gasto.Value).FillDown
        Sheets("Gastos").Calculate
        Sheets("Gastos").Range("R:R").NumberFormat = "dd/mm/yyyy"
    ElseIf Int(ultima) > Int(trab.parcelas_gasto.Value) Then
        Sheets("Gastos").Range("Q" & trab.parcelas_gasto.Value + 1 & ":S" & ultima).Clear
    ElseIf Int(ultima) < Int(trab.parcelas_gasto.Value) Then
        Sheets("Gastos").Range("Q" & ultima + 1).FormulaLocal = "=Q" & ultima & "+1"
        Sheets("Gastos").Range("R" & ultima + 1).FormulaLocal = "=DATAM($T$1;Q" & ultima + 1 & "-1)"
        If Int(ultima + 1) <> Int(trab.parcelas_gasto.Value) Then
            Sheets("Gastos").Range("Q" & ultima + 1 & ":R" & trab.parcelas_gasto.Value).FillDown
            Sheets("Gastos").Calculate
        End If
        Sheets("Gastos").Range("R:R").NumberFormat = "dd/mm/yyyy"
        Sheets("Gastos").Range("Q:Q").NumberFormat = "General"
    End If
    
    parcela.Value = 1
    Exit Sub

End Sub



Private Sub valor_Enter()
    If Mid(valor.Text, 1, 2) <> "R$" Then
        valor = "R$" + valor
    End If
    valor.Value = Format(valor, "R$#,##0.00")
End Sub
