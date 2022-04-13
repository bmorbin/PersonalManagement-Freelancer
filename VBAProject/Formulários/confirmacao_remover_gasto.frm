VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirmacao_remover_gasto 
   Caption         =   "Confirmar Remoção Parcela"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "confirmacao_remover_gasto.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirmacao_remover_gasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub distribui_Click()
    op2_selected.Visible = True
    op1_selected.Visible = False
    op3_selected.Visible = False
End Sub

Private Sub remove_all_Click()
    op2_selected.Visible = False
    op1_selected.Visible = True
    op3_selected.Visible = False
End Sub
Private Sub remove_single_Click()
    op2_selected.Visible = False
    op1_selected.Visible = False
    op3_selected.Visible = True
End Sub

Private Sub distribui_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    op1_image.Visible = False
    op3_image.Visible = False
    op2_image.Visible = True
End Sub

Private Sub remove_all_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    op2_image.Visible = False
    op3_image.Visible = False
    op1_image.Visible = True
    
End Sub


Private Sub remove_single_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    op2_image.Visible = False
    op1_image.Visible = False
    op3_image.Visible = True
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    op2_image.Visible = False
    op1_image.Visible = False
    op3_image.Visible = False
End Sub



Private Sub remover_cancela_Click()
    Call Unload(confirmacao_remover_gasto)
End Sub

Private Sub remover_confirma_Click()
    ultima = Sheets("Gastos").Range("A1000000").End(xlUp).Row
    If remove_all Then
        For i = 1 To ultima
            Sheets("gastos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & trab.id_egasto.Value & ";A2:A" & ultima & ";0);0)"
            result = Sheets("Gastos").Range("N1").Value
            Sheets("Gastos").Range("N1").Clear
            If result <> 0 Then
                Sheets("Gastos").Range(result + 1 & ":" & result + 1).Delete Shift:=xlUp
            Else
                Exit For
            End If
        Next i
        MsgBox ("Parcelas removidas com sucesso.")
    ElseIf remove_single Then
        comeco = 2
        For i = 1 To ultima
            Sheets("gastos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & trab.id_egasto.Value & ";A" & comeco & ":A" & ultima & ";0);0)"
            result = Sheets("Gastos").Range("N1").Value
            Sheets("Gastos").Range("N1").Clear
            If result <> 0 And comeco <= ultima Then
                linha = result + (comeco - 1)
                If Int(Sheets("Gastos").Range("D" & linha).Value) = Int(trab.parcela_egasto.Value) Then
                    Sheets("Gastos").Range(linha & ":" & linha).Delete Shift:=xlUp
                    comeco = linha
                ElseIf Int(Sheets("Gastos").Range("D" & linha).Value) > Int(trab.parcela_egasto.Value) Then
                    Sheets("Gastos").Range("D" & linha).Value = Sheets("Gastos").Range("D" & linha).Value - 1
                    comeco = linha + 1
                Else
                    comeco = linha + 1
                End If
            Else
                Exit For
            End If
        Next i
        MsgBox ("Parcela removida com sucesso.")
    ElseIf distribui Then
        Set array_linhas_do_id = CreateObject("System.Collections.ArrayList")
        
        comeco = 2
        For i = 1 To ultima
            Sheets("gastos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & trab.id_egasto.Value & ";A" & comeco & ":A" & ultima & ";0);0)"
            result = Sheets("Gastos").Range("N1").Value
            Sheets("Gastos").Range("N1").Clear
            If result <> 0 And comeco <= ultima Then
                linha = result + (comeco - 1)
                array_linhas_do_id.Add linha
                comeco = linha + 1
            Else
                Exit For
            End If
        Next i
        
        linha_id_remove = array_linhas_do_id(Int(trab.parcela_egasto.Value) - 1) 'array(0) pega primeiro valor da listaarray
        If array_linhas_do_id.Count = 1 Then
            Sheets("Gastos").Range(linha_id_remove & ":" & linha_id_remove).Delete Shift:=xlUp
            MsgBox ("Parcela removida com sucesso.")
        Else
            parte_valor = Sheets("Gastos").Range("C" & linha_id_remove).Value / (array_linhas_do_id.Count - 1) '.count para pegar tamanho da lista
            
            For Each linha In array_linhas_do_id
                If Int(Sheets("Gastos").Range("D" & linha).Value) > Int(trab.parcela_egasto.Value) Then
                    Sheets("Gastos").Range("D" & linha).Value = Sheets("Gastos").Range("D" & linha).Value - 1
                End If
                Sheets("Gastos").Range("C" & linha).Value = Sheets("Gastos").Range("C" & linha).Value + parte_valor
            Next linha
            Sheets("Gastos").Range(linha_id_remove & ":" & linha_id_remove).Delete Shift:=xlUp
            MsgBox ("Parcela removida com sucesso. Valores das parcelas seguintes atualizados.")
        End If
    End If
    ultima = Sheets("Gastos").Range("A1000000").End(xlUp).Row
    If ultima = 1 Then
        trab.lista_gasto.ListIndex = -1
        Call trab.reset_novo_gasto
        Call trab.novo_botton_Click
    Else
        trab.lista_gasto.ListIndex = ultima - 2
        trab.id_egasto.Value = Int(trab.lista_gasto.List(trab.lista_gasto.ListIndex, 0))
        trab.data_egasto.Value = CDate(trab.lista_gasto.List(trab.lista_gasto.ListIndex, 1))
        trab.valor_egasto.Value = CCur(trab.lista_gasto.List(trab.lista_gasto.ListIndex, 2))
        trab.parcela_egasto.Value = Int(trab.lista_gasto.List(trab.lista_gasto.ListIndex, 3))
        trab.nome_egasto.Value = (trab.lista_gasto.List(trab.lista_gasto.ListIndex, 4))
        trab.categoria_egasto.Value = (trab.lista_gasto.List(trab.lista_gasto.ListIndex, 5))
        trab.forma_egasto.Value = (trab.lista_gasto.List(trab.lista_gasto.ListIndex, 6))
    End If
    Call Unload(confirmacao_remover_gasto)
End Sub




