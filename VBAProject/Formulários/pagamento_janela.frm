VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pagamento_janela 
   Caption         =   "Pagamentos Registrados"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "pagamento_janela.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "pagamento_janela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pag_regs_normal.Visible = True
    pag_regis_selected.Visible = False
End Sub

Private Sub pag_regs_normal_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pag_regs_normal.Visible = False
    pag_regis_selected.Visible = True
End Sub
Private Sub pag_regis_selected_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MsgBox ("Filtre pelo ID do Trabalho para visualizar as parcelas de pagamento já pagas daquele trabalho na Tabela de Pagamentos. Nota: apenas os trabalhos que tiverem pagamentos já registrados aparecerão como opção de filtro.")
    pag_regis_selected.Visible = False
    pag_regs_normal.Visible = True
End Sub
Sub show_table_pagamentos()
    Sheets("Pagamentos_filtrado").Cells.Delete
    Sheets("Pagamentos_filtrado").Cells.Delete
    Sheets("Pagamentos").UsedRange.Copy Sheets("Pagamentos_filtrado").Range("A1")
    Sheets("Pagamentos_filtrado").UsedRange.AutoFilter
    Sheets("Pagamentos_filtrado").UsedRange.AutoFilter 1, IDtrab_paga.Value
    Sheets("Pagamentos_filtrado").Range("G:L").Delete

    linha = Sheets("Pagamentos_filtrado").Range("A1048576").End(xlUp).Row
    pagamento_janela.table_pagamentos.ColumnCount = 6
    pagamento_janela.table_pagamentos.ColumnHeads = False
    pagamento_janela.table_pagamentos.ColumnWidths = "18;90;78;66;36;60"
    
    Sheets("Pagamentos_filtrado").Activate
    Sheets("Pagamentos_filtrado").Range("A2:F" & linha).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy Sheets("Pagamentos_filtrado").Range("N2")
    Sheets("Pagamentos_filtrado").UsedRange.AutoFilter
    Sheets("Pagamentos_filtrado").Range("A1:F1").Copy Sheets("Pagamentos_filtrado").Range("N1")
    Sheets("Pagamentos_filtrado").Range("A:F").Clear
    linha = Sheets("Pagamentos_filtrado").Range("N1048576").End(xlUp).Row
    Sheets("Pagamentos_filtrado").Range("N1:S" & linha).Copy Sheets("Pagamentos_filtrado").Range("A1")
    Sheets("Pagamentos_filtrado").Range("N:S").Clear
    Application.CutCopyMode = False
    
    pagamento_janela.table_pagamentos.RowSource = "Pagamentos_filtrado!A2:F" & linha
    
    Sheets("Pagamentos_filtrado").Range("N1").FormulaLocal = "=SOMA(Pagamentos_filtrado!D:D)"
    total_pago.Value = Format(Sheets("Pagamentos_filtrado").Range("N1").Value, "R$#,##0.00")
    Sheets("Pagamentos_filtrado").Range("N1").Clear
    
    table_pagamentos.ListIndex = 0
    IDtrab_paga.Value = Int(table_pagamentos.List(table_pagamentos.ListIndex, 0))
    trab_paga.Value = table_pagamentos.List(table_pagamentos.ListIndex, 1)
    data_paga.Value = CDate(table_pagamentos.List(table_pagamentos.ListIndex, 2))
    valor_paga.Value = CCur(table_pagamentos.List(table_pagamentos.ListIndex, 3))
    parcela_paga.Value = Int(table_pagamentos.List(table_pagamentos.ListIndex, 4))
    pagou_paga.Value = table_pagamentos.List(table_pagamentos.ListIndex, 5)
    
End Sub

Private Sub data_paga_Change()
    Call trab.data(data_paga)
    If IDtrab_paga.Value <> "" Then
        Call busca_parcela
    End If
End Sub

Private Sub editar_paga_Click()
    On Error GoTo errodin
    If CCur(valor_paga.Value) Then
    End If
    If valor_paga.Value < 0 Then
        GoTo errodin
    End If
    
    On Error GoTo errodata
    If CDate(data_paga.Value) Then
    End If
    
    If table_pagamentos.ListIndex < 0 Then
        GoTo ErroNenhumSelecionado
    End If
    
    On Error GoTo ErroID
    If Int(IDtrab_paga.Value) Then
    End If
    
    If Int(IDtrab_paga.Value) <> Int(table_pagamentos.List(table_pagamentos.ListIndex, 0)) Then
        GoTo ErroID
    End If
    
    Call atualiza_edit
    Call ops_pagou_paga
    
    MsgBox ("Editado com sucesso.")
    
    Exit Sub
errodin:
    MsgBox ("Preencha corretamente o campo da quantia recebida na parcela.")
    valor_paga.SetFocus
    Exit Sub
errodata:
    MsgBox ("Preencha corretamente o campo da data da parcela no formato ""dd/mm/aaaa"".")
    data_paga.SetFocus
    Exit Sub
ErroNenhumSelecionado:
    MsgBox ("É necessário selecionar um item da lista para editá-lo. Dê dois cliques no item desejado.")
    data_paga.SetFocus
    Exit Sub
ErroID:
    MsgBox ("Insira o ID correspondente ao selecionado na lista.")
    IDtrab_paga.SetFocus
    Exit Sub
End Sub
Sub atualiza_edit()
    id = IDtrab_paga.Value
    
    ultima = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If ultima = 1 Then
        Sheets("Pagamentos").Cells(2, 3).Value = CDate(data_paga.Value)
        Sheets("Pagamentos").Cells(2, 4).Value = CCur(valor_paga.Value)
        Sheets("Pagamentos").Cells(2, 5).Value = Int(parcela_paga.Value)
        Sheets("Pagamentos").Cells(2, 6).Value = pagou_paga.Value
        Call show_table_pagamentos
        Exit Sub
    End If
    
    localizacao_item_selecionado = table_pagamentos.ListIndex + 1
    
    comeco = 2
    For i = 1 To ultima
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value
        Sheets("Pagamentos").Range("N1").Clear
        If result <> 0 And comeco <= ultima Then
            linha = result + (comeco - 1)
            If i = localizacao_item_selecionado Then
                Sheets("Pagamentos").Range("C" & linha).Value = CDate(data_paga.Value)
                Sheets("Pagamentos").Range("D" & linha).Value = CCur(valor_paga.Value)
                Sheets("Pagamentos").Range("E" & linha).Value = Int(parcela_paga.Value)
                Sheets("Pagamentos").Range("F" & linha).Value = pagou_paga.Value
            End If
            data_encontrada = Sheets("Pagamentos").Range("C" & linha).Value
            Sheets("Pagamentos").Range("O" & i).Value = CDate(data_encontrada)
            comeco = linha + 1
        Else
            Sheets("Pagamentos").Range("P1").FormulaLocal = "=ORDEM.EQ(O1;$O$1:$O$" & i - 1 & ";1)+CONT.SE($O$1:O1;O1)-1"
            If i - 1 > 1 Then
                Sheets("Pagamentos").Range("P1:P" & i - 1).FillDown
                Sheets("Pagamentos").Calculate
            End If
            Exit For
        End If
    Next i
    
    comeco = 2
    For i = 1 To ultima
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value
        Sheets("Pagamentos").Range("N1").Clear
        If result = 0 Then
            Exit For
        End If
        linha = result + (comeco - 1)
        comeco = linha + 1
        Sheets("Pagamentos").Range("E" & linha).Value = Sheets("Pagamentos").Range("P" & i).Value
    Next i
    
    Sheets("Pagamentos").Range("O:P").Clear
    
    Call show_table_pagamentos
    
End Sub

Private Sub IDtrab_paga_AfterUpdate()
    TabKeyBehavior = False
    Call show_table_pagamentos

End Sub



Private Sub remover_paga_Click()
    confirma_excluir.Show 'confirmar exclusao
    
    linha_pag = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    
    If linha_pag = 1 Then
        Call Unload(pagamento_janela)
        Exit Sub
    End If
    
End Sub

Sub remover_pag()
    Sheets("Pagamentos").Activate
    If table_pagamentos.ListIndex < 0 Then
        MsgBox ("Clique duas vezes no registro de pagamento antes de remover.")
        Exit Sub
    End If
    
    id_remover = IDtrab_paga.Value
    parcela_remover = Int(table_pagamentos.List(table_pagamentos.ListIndex, 4))
    valor_remover = CCur(table_pagamentos.List(table_pagamentos.ListIndex, 3))
    data_remover = CDate(table_pagamentos.List(table_pagamentos.ListIndex, 2))
    pagador_remover = table_pagamentos.List(table_pagamentos.ListIndex, 5)
    
    On Error GoTo Sair
    If parcela_remover <> Int(parcela_paga.Value) Or data_remover <> CDate(data_paga.Value) Or valor_remover <> CCur(valor_paga.Value) _
    Or pagador_remover <> pagou_paga.Value Then
        MsgBox ("Parcela não encontrada. Certifique-se de não ter alterado os valores nas caixas de entrada e clique 2x novamente no item a ser excluído.")
        Exit Sub
    End If
    
    ultima = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If ultima = 2 Then
        Sheets("Pagamentos").Range("2:2").Delete Shift:=xlUp
        MsgBox ("Removido com sucesso")
        trab.editar_pagamento.Enabled = False
        Exit Sub
    End If
    
    comeco = 2
    For i = 1 To ultima
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id_remover & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value
        Sheets("Pagamentos").Range("N1").Clear
        If result <> 0 And comeco <= ultima Then
            linha = result + (comeco - 1)
            data_encontrada = Sheets("Pagamentos").Range("C" & linha).Value
            parcela_encontrada = Sheets("Pagamentos").Range("E" & linha).Value
            valor_encontrada = Sheets("Pagamentos").Range("D" & linha).Value
            pagador_encontrada = Sheets("Pagamentos").Range("F" & linha).Value
            If data_encontrada = data_remover And parcela_encontrada = parcela_remover And valor_encontrada = valor_remover _
            And pagador_encontrada = pagador_remover Then
                Sheets("Pagamentos").Range(linha & ":" & linha).Delete Shift:=xlUp
                comeco = linha
                i = i - 1
            Else
                Sheets("Pagamentos").Range("O" & i).Value = CDate(data_encontrada)
                comeco = linha + 1
            End If
        Else
            If i = 1 Then
                Exit For
            End If
            Sheets("Pagamentos").Range("P1").FormulaLocal = "=ORDEM.EQ(O1;$O$1:$O$" & i - 1 & ";1)+CONT.SE($O$1:O1;O1)-1"
            If i - 1 > 1 Then
                Sheets("Pagamentos").Range("P1:P" & i - 1).FillDown
                Sheets("Pagamentos").Calculate
            End If
            Exit For
        End If
    Next i
    
    If i > 1 Then
        comeco = 2
        For i = 1 To ultima
            Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id_remover & ";A" & comeco & ":A" & ultima & ";0);0)"
            result = Sheets("Pagamentos").Range("N1").Value
            Sheets("Pagamentos").Range("N1").Clear
            If result = 0 Then
                Exit For
            End If
            linha = result + (comeco - 1)
            comeco = linha + 1
            Sheets("Pagamentos").Range("E" & linha).Value = Sheets("Pagamentos").Range("P" & i).Value
        Next i
    End If
    Sheets("Pagamentos").Range("O:P").Clear
    
    MsgBox ("Parcela removida com sucesso.")
    
    linha_pag = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    
    If linha_pag = 1 Then
        Exit Sub
    End If
    
    linha_filtro = Sheets("Pagamentos_filtrado").Range("A1000000").End(xlUp).Row
    If linha_filtro = 2 Then
        table_pagamentos.RowSource = ""
        Call ops_id
        IDtrab_paga.ListIndex = 0
        Call show_table_pagamentos
        Exit Sub
    End If
    
    Call show_table_pagamentos
    Call ops_id

    Exit Sub

Sair:
    MsgBox ("ERRO: Parcela não encontrada. Certifique-se de não ter alterado os valores nas caixas de entrada e clique 2x novamente no item a ser excluído.")
    Exit Sub
End Sub


Public Sub table_pagamentos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Saida
    IDtrab_paga.Value = Int(table_pagamentos.List(table_pagamentos.ListIndex, 0))
    trab_paga.Value = table_pagamentos.List(table_pagamentos.ListIndex, 1)
    data_paga.Value = CDate(table_pagamentos.List(table_pagamentos.ListIndex, 2))
    valor_paga.Value = CCur(table_pagamentos.List(table_pagamentos.ListIndex, 3))
    parcela_paga.Value = Int(table_pagamentos.List(table_pagamentos.ListIndex, 4))
    pagou_paga.Value = table_pagamentos.List(table_pagamentos.ListIndex, 5)
Saida:
    Exit Sub
End Sub

Private Sub table_pagamentos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListScroll Me, Me.table_pagamentos
End Sub

Private Sub UserForm_Initialize()
    Call ops_pagou_paga
    Call ops_id
    IDtrab_paga.ListIndex = 0
    Call show_table_pagamentos
End Sub

Private Sub UserForm_Terminate()
    Sheets("Pagamentos_filtrado").UsedRange.Clear
    Sheets("Pagamentos_filtrado").UsedRange.Delete
End Sub

Private Sub valor_paga_Enter()
    If Mid(valor_paga.Text, 1, 2) <> "R$" Then
        valor_paga = "R$" + valor_paga
    End If
    valor_paga.Value = Format(valor_paga, "R$#,##0.00")
End Sub

Sub ops_pagou_paga()
    pagou_paga.Clear
    linha = Sheets("Pagamentos").Range("F1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Pagamentos").Range("F1:F" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Pagamentos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Pagamentos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Pagamentos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              pagou_paga.AddItem (opc)
        End If
    Next opc
    Sheets("Pagamentos").Range("N:N").Clear
    Application.CutCopyMode = False
End Sub

Sub ops_id()
    IDtrab_paga.Clear
    linha = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then
        Call Unload(pagamento_janela)
        Exit Sub
    End If
    
    Sheets("Pagamentos").Range("A1:A" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Pagamentos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Pagamentos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Pagamentos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              IDtrab_paga.AddItem (opc)
        End If
    Next opc
    Sheets("Pagamentos").Range("N:N").Clear
    Application.CutCopyMode = False
End Sub

Sub busca_parcela()
    On Error GoTo Sair
    If CDate(data_paga.Value) Then
    End If
        
    ultima = Sheets("Pagamentos_filtrado").Range("A1000000").End(xlUp).Row
    If ultima <= 2 Then
        parcela_paga.Value = 1
        Exit Sub
    End If
    
    Sheets("Pagamentos_filtrado").Range("C2:C" & ultima).Copy Sheets("Pagamentos_filtrado").Range("N1")
    Sheets("Pagamentos_filtrado").Range("N" & table_pagamentos.ListIndex + 1).Value = CDate(data_paga.Value)
    Application.CutCopyMode = False
    Sheets("Pagamentos_filtrado").Range("O" & table_pagamentos.ListIndex + 1).FormulaLocal = "=ORDEM.EQ(N" & table_pagamentos.ListIndex + 1 & ";$N$1:$N$" & ultima - 1 & ";1)+CONT.SE($N$1:N1;N1)-1"
    parcela_paga.Value = Sheets("Pagamentos_filtrado").Range("O" & table_pagamentos.ListIndex + 1).Value
    Sheets("Pagamentos_filtrado").Range("N1:O" & ultima).Clear
    
    Exit Sub
Sair:
    parcela_paga = ""
    Exit Sub
End Sub
