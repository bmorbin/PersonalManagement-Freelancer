VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trab 
   Caption         =   "Gestão Pessoal"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   OleObjectBlob   =   "trab.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "trab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private CelulaCaminhoImagem As Range
Private ArquivoExiste       As Object


Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pag_interroga_normal.Visible = True
    pag_interroga_select.Visible = False
End Sub







Private Sub pag_interroga_normal_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pag_interroga_normal.Visible = False
    pag_interroga_select.Visible = True
End Sub

Private Sub pag_interroga_select_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MsgBox ("Dê dois cliques no trabalho na Tabela de Registros. Caso seja um novo trabalho, certifique-se de tê-lo salvo na aba de Novo Registro. Aperte o botão EDITAR para visualizar os pagamentos já registrados.")
    pag_interroga_select.Visible = False
    pag_interroga_normal.Visible = True
    
End Sub

Private Sub tdespesas_frame1_Click()
'    despesas_frame1.Visible = False
'    tdespesas_frame1.Visible = False
'    despesas_frame2.Visible = True
'    tdespesas_frame2.Visible = True
    registros_frame1.Visible = False
    tregistros_frame1.Visible = False
    registros_frame2.Visible = True
    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
    receitas_frame1.Visible = False
    treceitas_frame1.Visible = False
    receitas_frame2.Visible = True
    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = False
    tregistros_frame0.Visible = False
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = False
    treceitas_frame0.Visible = False
    despesas_frame0.Visible = True
    tdespesas_frame0.Visible = True
    
    trab.MultiPage1.Value = 1
End Sub


Private Sub treceitas_frame1_Click()
    despesas_frame1.Visible = False
    tdespesas_frame1.Visible = False
    despesas_frame2.Visible = True
    tdespesas_frame2.Visible = True
    registros_frame1.Visible = False
    tregistros_frame1.Visible = False
    registros_frame2.Visible = True
    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
'    receitas_frame1.Visible = False
'    treceitas_frame1.Visible = False
'    receitas_frame2.Visible = True
'    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = False
    tregistros_frame0.Visible = False
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = True
    treceitas_frame0.Visible = True
    despesas_frame0.Visible = False
    tdespesas_frame0.Visible = False
    
    trab.MultiPage1.Value = 0
    If editar_pagamento.Enabled = True Then
        editar_pagamento_Click
    Else
        MsgBox ("Necessário registrar algum pagamento na janela ""Registro Rápido de Pagamentos"" para acessar as Receitas.")
    End If
End Sub


Private Sub tregistros_frame1_Click()
    despesas_frame1.Visible = False
    tdespesas_frame1.Visible = False
    despesas_frame2.Visible = True
    tdespesas_frame2.Visible = True
'    registros_frame1.Visible = False
'    tregistros_frame1.Visible = False
'    registros_frame2.Visible = True
'    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
    receitas_frame1.Visible = False
    treceitas_frame1.Visible = False
    receitas_frame2.Visible = True
    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = True
    tregistros_frame0.Visible = True
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = False
    treceitas_frame0.Visible = False
    despesas_frame0.Visible = False
    tdespesas_frame0.Visible = False
    
    trab.MultiPage1.Value = 0
    nome_registro.SetFocus
End Sub

Private Sub tdespesas_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    despesas_frame2.Visible = False
    tdespesas_frame2.Visible = False
    despesas_frame1.Visible = True
    tdespesas_frame1.Visible = True
    
End Sub
'Private Sub tdespesas_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    despesas_frame2.Visible = True
'    tdespesas_frame2.Visible = True
'    despesas_frame1.Visible = False
'    tdespesas_frame1.Visible = False
'
'End Sub

Private Sub tdash_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dash_frame2.Visible = False
    tdash_frame2.Visible = False
    dash_frame1.Visible = True
    tdash_frame1.Visible = True
    
End Sub
'Private Sub tdash_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    dash_frame2.Visible = True
'    tdash_frame2.Visible = True
'    dash_frame1.Visible = False
'    tdash_frame1.Visible = False
'
'End Sub

Private Sub treceitas_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    receitas_frame2.Visible = False
    treceitas_frame2.Visible = False
    receitas_frame1.Visible = True
    treceitas_frame1.Visible = True
    
End Sub
'Private Sub treceitas_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    receitas_frame2.Visible = True
'    treceitas_frame2.Visible = True
'    receitas_frame1.Visible = False
'    treceitas_frame1.Visible = False
'
'End Sub

Private Sub tregistros_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    registros_frame2.Visible = False
    tregistros_frame2.Visible = False
    registros_frame1.Visible = True
    tregistros_frame1.Visible = True
    
End Sub
'Private Sub tregistros_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    registros_frame2.Visible = True
'    tregistros_frame2.Visible = True
'    registros_frame1.Visible = False
'    tregistros_frame1.Visible = False
'
'End Sub



Private Sub despesas_frame1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    despesas_frame1.Visible = False
'    tdespesas_frame1.Visible = False
'    despesas_frame2.Visible = True
'    tdespesas_frame2.Visible = True
    registros_frame1.Visible = False
    tregistros_frame1.Visible = False
    registros_frame2.Visible = True
    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
    receitas_frame1.Visible = False
    treceitas_frame1.Visible = False
    receitas_frame2.Visible = True
    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = False
    tregistros_frame0.Visible = False
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = False
    treceitas_frame0.Visible = False
    despesas_frame0.Visible = True
    tdespesas_frame0.Visible = True
    
    trab.MultiPage1.Value = 1
End Sub

Private Sub receitas_frame1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    despesas_frame1.Visible = False
    tdespesas_frame1.Visible = False
    despesas_frame2.Visible = True
    tdespesas_frame2.Visible = True
    registros_frame1.Visible = False
    tregistros_frame1.Visible = False
    registros_frame2.Visible = True
    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
'    receitas_frame1.Visible = False
'    treceitas_frame1.Visible = False
'    receitas_frame2.Visible = True
'    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = False
    tregistros_frame0.Visible = False
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = True
    treceitas_frame0.Visible = True
    despesas_frame0.Visible = False
    tdespesas_frame0.Visible = False
    
    trab.MultiPage1.Value = 0
    If editar_pagamento.Enabled = True Then
        editar_pagamento_Click
    End If
End Sub

Private Sub registros_frame1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    despesas_frame1.Visible = False
    tdespesas_frame1.Visible = False
    despesas_frame2.Visible = True
    tdespesas_frame2.Visible = True
'    registros_frame1.Visible = False
'    tregistros_frame1.Visible = False
'    registros_frame2.Visible = True
'    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
    receitas_frame1.Visible = False
    treceitas_frame1.Visible = False
    receitas_frame2.Visible = True
    treceitas_frame2.Visible = True
    
    registros_frame0.Visible = True
    tregistros_frame0.Visible = True
    dash_frame0.Visible = False
    tdash_frame0.Visible = False
    receitas_frame0.Visible = False
    treceitas_frame0.Visible = False
    despesas_frame0.Visible = False
    tdespesas_frame0.Visible = False
    
    trab.MultiPage1.Value = 0
    nome_registro.SetFocus
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    despesas_frame1.Visible = False
    tdespesas_frame1.Visible = False
    despesas_frame2.Visible = True
    tdespesas_frame2.Visible = True
    registros_frame1.Visible = False
    tregistros_frame1.Visible = False
    registros_frame2.Visible = True
    tregistros_frame2.Visible = True
    dash_frame1.Visible = False
    tdash_frame1.Visible = False
    dash_frame2.Visible = True
    tdash_frame2.Visible = True
    receitas_frame1.Visible = False
    treceitas_frame1.Visible = False
    receitas_frame2.Visible = True
    treceitas_frame2.Visible = True
End Sub



Private Sub despesas_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    despesas_frame2.Visible = False
    tdespesas_frame2.Visible = False
    despesas_frame1.Visible = True
    tdespesas_frame1.Visible = True
    
End Sub
'Private Sub despesas_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    despesas_frame2.Visible = True
'    tdespesas_frame2.Visible = True
'    despesas_frame1.Visible = False
'    tdespesas_frame1.Visible = False
'
'End Sub

Private Sub dash_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dash_frame2.Visible = False
    tdash_frame2.Visible = False
    dash_frame1.Visible = True
    tdash_frame1.Visible = True
    
End Sub
'Private Sub dash_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    dash_frame2.Visible = True
'    tdash_frame2.Visible = True
'    dash_frame1.Visible = False
'    tdash_frame1.Visible = False
'
'End Sub

Private Sub receitas_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    receitas_frame2.Visible = False
    treceitas_frame2.Visible = False
    receitas_frame1.Visible = True
    treceitas_frame1.Visible = True
    
End Sub
'Private Sub receitas_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    receitas_frame2.Visible = True
'    treceitas_frame2.Visible = True
'    receitas_frame1.Visible = False
'    treceitas_frame1.Visible = False
'
'End Sub

Private Sub registros_frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    registros_frame2.Visible = False
    tregistros_frame2.Visible = False
    registros_frame1.Visible = True
    tregistros_frame1.Visible = True
    
End Sub
'Private Sub registros_frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    registros_frame2.Visible = True
'    tregistros_frame2.Visible = True
'    registros_frame1.Visible = False
'    tregistros_frame1.Visible = False
'
'End Sub



Private Sub Foto_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim strCaminhoImagem    As String
    
    strCaminhoImagem = Application.GetOpenFilename("*.jpg,*.jpg,*.bmp,*.bmp")
    
    On Error GoTo Sair
    If ArquivoExiste.FileExists(strCaminhoImagem) = True Then
        MsgBox ("Foto salva com sucesso. Saia e entre no programa novamente para atualizá-la.")
        CelulaCaminhoImagem.Value = strCaminhoImagem
        Foto.Picture = LoadPicture(CelulaCaminhoImagem.Value)
        Foto.PictureSizeMode = fmPictureSizeModeZoom
    Else
        MsgBox ("Processo não concluído, a imagem não foi encontrada.")
    End If
    
    Exit Sub
Sair:
    MsgBox ("Erro ao importar imagem. Certifique-se do arquivo estar no formato "".jpg"".")
    Exit Sub
End Sub

Private Sub Apelido_AfterUpdate()
    Apelido.Locked = True
    nome_real.Caption = Apelido.Value
    If Apelido.Value <> "" Then
        Sheets("Planilha1").Range("A2").Value = Apelido.Value
    End If
    Apelido.Value = ""
    
End Sub

Private Sub Apelido_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    nome_real.Caption = ""
    Apelido.Locked = False
    Apelido.Value = ""
    Apelido.SetFocus
End Sub


Sub show_lista_gasto()
    linha = Sheets("Gastos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    
    lista_gasto.ColumnCount = 7
    lista_gasto.ColumnHeads = False
    lista_gasto.ColumnWidths = "18;54;60;30;72;60;54"
    lista_gasto.RowSource = "Gastos!A2:G" & linha

End Sub



Private Sub lista_gasto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListScroll Me, Me.lista_gasto
End Sub





Private Sub remover_gasto_Click()
    confirmacao_remover_gasto.Show
End Sub

Sub auto_botton_Click()
    preencher_botton.Enabled = False
    Sheets("Gastos").Range("Q:U").Clear
End Sub

Private Sub data_gasto_AfterUpdate()
    On Error GoTo Sair
    If CDate(data_gasto.Value) Then
        Sheets("gastos").Range("Q:T").Clear
    End If
Sair:
    Exit Sub
End Sub

Private Sub editar_gasto_Click()
    On Error GoTo arrumardata
    If CDate(data_egasto.Value) Then
    End If
    
    If Year(CDate(data_egasto.Value)) < 2000 Then
        GoTo arrumardata
    End If
    
    On Error GoTo arrumardin
    If CCur(valor_egasto.Value) Then
    End If
    If valor_egasto.Value < 0 Then
        MsgBox ("Digite corretamente o valor da parcela. Nota: Necessário colocar valor positivo somente.")
        valor_egasto.SetFocus
        Exit Sub
    End If
    
    If categoria_egasto.Value = "" Then
        MsgBox ("Necessário definir categoria da compra. Exemplos: Mercado, Vestimentas, Eletrônicos, etc.")
        categoria_egasto.SetFocus
        Exit Sub
    End If
    
    If nome_egasto.Value = "" Then
        nome_egasto.Value = "Gasto " & id_egasto.Value
    End If
    
    ultima = Sheets("Gastos").Range("A1000000").End(xlUp).Row
    
    Set array_linhas_do_id = CreateObject("System.Collections.ArrayList")
    
    comeco = 2
    For i = 1 To ultima
        Sheets("Gastos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id_egasto.Value & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Gastos").Range("N1").Value
        Sheets("Gastos").Range("N1").Clear
        If result <> 0 And comeco <= ultima Then
            linha = result + (comeco - 1)
            array_linhas_do_id.Add linha
            If Int(Sheets("Gastos").Range("D" & linha).Value) = Int(parcela_egasto.Value) - 1 Then
                If CDate(Sheets("Gastos").Range("B" & linha).Value) > CDate(data_egasto.Value) Then
                    MsgBox ("A data inserida na parcela é menor que a data da parcela anterior.")
                    data_egasto.SetFocus
                    Exit Sub
                End If
            ElseIf Int(Sheets("Gastos").Range("D" & linha).Value) = Int(parcela_egasto.Value) + 1 Then
                If CDate(Sheets("Gastos").Range("B" & linha).Value) < CDate(data_egasto.Value) Then
                    MsgBox ("A data inserida na parcela é maior que a data da parcela posterior.")
                    data_egasto.SetFocus
                    Exit Sub
                End If
            End If
            comeco = linha + 1
        Else
            Exit For
        End If
    Next i
    
    
    For Each linha In array_linhas_do_id
        If Int(Sheets("Gastos").Range("D" & linha).Value) = Int(parcela_egasto.Value) Then
            Sheets("GAstos").Range("B" & linha).Value = CDate(data_egasto.Value)
            Sheets("GAstos").Range("C" & linha).Value = CCur(valor_egasto.Value)
            Sheets("GAstos").Range("E" & linha).Value = nome_egasto.Value
            Sheets("GAstos").Range("F" & linha).Value = categoria_egasto.Value
            Sheets("GAstos").Range("G" & linha).Value = forma_egasto.Value
        Else
            Sheets("GAstos").Range("E" & linha).Value = nome_egasto.Value
            Sheets("GAstos").Range("F" & linha).Value = categoria_egasto.Value
        End If
    Next linha
    
    cetegoria_antes = lista_gasto.List(lista_gasto.ListIndex, 5)
    forma_antes = lista_gasto.List(lista_gasto.ListIndex, 6)
    If categoria_antes <> categoria_egasto.Value Then
        Call ops_categoria_gasto
    End If
    If forma_antes <> forma_egasto.Value Then
        Call ops_forma_gasto
    End If
    
    valor_antigo = CCur(lista_gasto.List(lista_gasto.ListIndex, 2))
    If valor_antigo <> CCur(valor_egasto.Value) Then
        MsgBox ("Parcela da compra editada com sucesso. Nota: o valor total de """ & nome_egasto.Value & """ foi alterado com essa edição.")
    Else
        MsgBox ("Parcela da compra editada com sucesso.")
    End If
    
    lista_gasto.ListIndex = ultima - 2
    id_egasto.Value = Int(lista_gasto.List(lista_gasto.ListIndex, 0))
    data_egasto.Value = CDate(lista_gasto.List(lista_gasto.ListIndex, 1))
    valor_egasto.Value = CCur(lista_gasto.List(lista_gasto.ListIndex, 2))
    parcela_egasto.Value = Int(lista_gasto.List(lista_gasto.ListIndex, 3))
    nome_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 4))
    categoria_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 5))
    forma_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 6))
    
    Exit Sub
arrumardata:
    MsgBox ("Data de início do gasto inválida.")
    data_egasto.SetFocus
    Exit Sub
arrumardin:
    MsgBox ("Digite o valor total da compra antes de salvar.")
    valor_egasto.SetFocus
    Exit Sub
    
End Sub

Private Sub incremento_parcelas_SpinDown()
    If parcelas_gasto = 2 Then
        Exit Sub
    End If
    parcelas_gasto.Value = parcelas_gasto.Value - 1
End Sub

Private Sub incremento_parcelas_SpinUp()
    parcelas_gasto.Value = parcelas_gasto.Value + 1
End Sub


Private Sub manual_botton_Click()
    preencher_botton.Enabled = True
End Sub

Private Sub nao_parcelado_botton_Click()
    op_nao_parc.Visible = False
    op_sim_parc.Visible = True
    
    Call bloqueia_parcelas
    auto_botton = True
    categoria_gasto.SetFocus
    Sheets("Gastos").Range("Q:T").Clear
End Sub



Private Sub parcelado_botton_Click()
    op_sim_parc.Visible = False
    op_nao_parc.Visible = True
    
    
    parcelas_gasto.Enabled = True
    parcelas_gasto.SetFocus
    parcelas_gasto.Value = 2
    preencher_parcelas.Enabled = True
    auto_botton.Enabled = True
    manual_botton.Enabled = True
    pre_visu_parcelas.Enabled = True
    incremento_parcelas.Enabled = True
    preencher_botton.Enabled = True
    
    Call auto_botton_Click
End Sub

Private Sub data_egasto_Change()
    Call data(data_egasto)
End Sub


Private Sub parcelas_gasto_AfterUpdate()
    'conferindo se a entrada é válida
    On Error GoTo Sair
    If Int(parcelas_gasto.Value) Then
    End If
    If parcelas_gasto.Value < 2 Then
        GoTo Sair
    End If
    
    Exit Sub
Sair:
    MsgBox ("Digite corretamente o número de parcelas.")
    parcelas_gasto.Value = 2
    Exit Sub
End Sub


Private Sub pre_visu_parcelas_Click()
    On Error GoTo errodata
    If CDate(data_gasto.Value) Then
    End If
    
    If Year(CDate(data_gasto.Value)) < 2000 Then
        GoTo errodata
    End If
    
    On Error GoTo errodin
    If CCur(valor_gasto.Value) Then
    End If
    If valor_gasto.Value < 0 Then
        GoTo errodin
    End If
    
    If auto_botton Then
        Sheets("Gastos").Range("Q:T").Clear
        valor_cada_parcela = valor_gasto.Value / parcelas_gasto.Value
        For i = 1 To parcelas_gasto.Value
            Sheets("Gastos").Range("Q" & i).Value = i
            Sheets("Gastos").Range("S" & i).Value = CCur(valor_cada_parcela)
        Next i
        Sheets("Gastos").Range("T1").Value = CDate(data_gasto.Value)
        Sheets("Gastos").Range("R1").FormulaLocal = "=DATAM($T$1;Q1-1)"
        Sheets("Gastos").Range("R1:R" & parcelas_gasto.Value).FillDown
        Sheets("Gastos").Calculate
        Sheets("Gastos").Range("R:R").NumberFormat = "dd/mm/yyyy"
        
        visualizar_parcelas.Show
        Sheets("Gastos").Range("Q1:T" & parcelas_gasto.Value).Clear
    End If
    If manual_botton Then
        ultima = Sheets("gastos").Range("Q1000000").End(xlUp).Row
        If ultima = 1 Then
            MsgBox ("Parcelas não preenchidas ainda. Preencha-as para visualizá-las.")
            preencher_botton.SetFocus
            Exit Sub
        Else
            visualizar_parcelas.Show
        End If
             
    End If
    
    Exit Sub
errodata:
    MsgBox ("Data inválida. Corrija a data do início da compra antes de visualizar as parcelas.")
    data_gasto.SetFocus
    Exit Sub
errodin:
    MsgBox ("Valor gasto inválido. Preencha corretamente o campo do valor total gasto na compra.")
    valor_gasto.SetFocus
    Exit Sub
End Sub



Private Sub preencher_botton_Click()
    On Error GoTo erroparcelas
    If Int(parcelas_gasto.Value) < 2 Then
        MsgBox ("Digite o número de parcelas da compra antes de preenchê-las. Nota: mínimo 2 parcelas.")
        Exit Sub
    End If
    On Error GoTo arrumardata
    If CDate(data_gasto.Value) Then
    End If
    
    If Year(CDate(data_gasto.Value)) < 2000 Then
        GoTo arrumardata
    End If
    
    On Error GoTo arrumardin
    If CCur(valor_gasto.Value) Then
    End If
    If valor_gasto.Value < 0 Then
        MsgBox ("Digite o valor total da compra antes de preencher as parcelas.")
        valor_gasto.SetFocus
        Exit Sub
    End If
    
    parcelas_manual.Show
    Exit Sub
erroparcelas:
    MsgBox ("Digite o número de parcelas da compra antes de preenchê-las.")
    parcelas_gasto.SetFocus
    Exit Sub
arrumardata:
    MsgBox ("Data de início do gasto inválida. Corrija-a antes de preencher as parcelas.")
    data_gasto.SetFocus
    Exit Sub
arrumardin:
    MsgBox ("Digite o valor total da compra antes de preencher as parcelas.")
    valor_gasto.SetFocus
    Exit Sub
End Sub



Private Sub salvar_gasto_Click()
    On Error GoTo arrumardata
    If CDate(data_gasto.Value) Then
    End If
    
    If Year(CDate(data_gasto.Value)) < 2000 Then
        GoTo arrumardata
    End If
    
    On Error GoTo arrumardin
    If CCur(valor_gasto.Value) Then
    End If
    If valor_gasto.Value < 0 Then
        MsgBox ("Digite corretamente o valor total da compra. Nota: Necessário colocar valor positivo somente.")
        valor_gasto.SetFocus
        Exit Sub
    End If
    
    If categoria_gasto.Value = "" Then
        MsgBox ("Necessário definir categoria da compra. Exemplos: Mercado, Vestimentas, Eletrônicos, etc.")
        categoria_gasto.SetFocus
        Exit Sub
    End If
    
    If nome_gasto.Value = "" Then
        nome_gasto.Value = "Gasto " & id_gasto.Value
    End If
    
    linha_adc = Sheets("gastos").Range("A1000000").End(xlUp).Row + 1
    If nao_parcelado_botton Then
        Sheets("Gastos").Cells(linha_adc, 1).Value = Int(id_gasto.Value)
        Sheets("Gastos").Cells(linha_adc, 2).Value = CDate(data_gasto.Value)
        Sheets("Gastos").Cells(linha_adc, 3).Value = CCur(valor_gasto.Value)
        Sheets("Gastos").Cells(linha_adc, 4).Value = Int(parcelas_gasto.Value)
        Sheets("Gastos").Cells(linha_adc, 5).Value = nome_gasto.Value
        Sheets("Gastos").Cells(linha_adc, 6).Value = categoria_gasto.Value
        Sheets("Gastos").Cells(linha_adc, 7).Value = forma_gasto.Value
    Else
        On Error GoTo arrumarparcelas
        If Int(parcelas_gasto.Value) Then
        End If
        If parcelas_gasto.Value < 0 Then
            GoTo arrumarparcelas
        End If
        If auto_botton Then
            Sheets("Gastos").Range("Q:T").Clear
            valor_cada_parcela = valor_gasto.Value / parcelas_gasto.Value
            For i = 1 To parcelas_gasto.Value
                Sheets("Gastos").Range("Q" & i).Value = i
                Sheets("Gastos").Range("S" & i).Value = CCur(valor_cada_parcela)
            Next i
            Sheets("Gastos").Range("T1").Value = CDate(data_gasto.Value)
            Sheets("Gastos").Range("R1").FormulaLocal = "=DATAM($T$1;Q1-1)"
            Sheets("Gastos").Range("R1:R" & parcelas_gasto.Value).FillDown
            Sheets("Gastos").Calculate
            Sheets("Gastos").Range("R:R").NumberFormat = "dd/mm/yyyy"
        
            For i = 1 To parcelas_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 1).Value = Int(id_gasto.Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 2).Value = CDate(Sheets("GAstos").Range("R" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 3).Value = CCur(Sheets("GAstos").Range("S" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 4).Value = Int(Sheets("GAstos").Range("Q" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 5).Value = nome_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 6).Value = categoria_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 7).Value = forma_gasto.Value
            Next i
            Sheets("Gastos").Range("Q1:T" & parcelas_gasto.Value).Clear
        ElseIf manual_botton Then
            ultima_manual = Sheets("Gastos").Range("Q1000000").End(xlUp).Row
            If Int(ultima_manual) > Int(parcelas_gasto.Value) Then
                MsgBox ("Foram preenchidas mais parcelas do que a quantidade definida. Preencha novamente as parcelas.")
                preencher_botton.SetFocus
                Exit Sub
            ElseIf Int(ultima_manual) < Int(parcelas_gasto.Value) Then
                MsgBox ("Não foram preenchidas todas parcelas. Preencha-as antes de salvar.")
                preencher_botton.SetFocus
                Exit Sub
            Else
                ultima_linha_parcela = Sheets("Gastos").Range("Q1000000").End(xlUp).Row
                For i = 1 To ultima_linha_parcela
                    If Sheets("Gastos").Range("S" & i).Value = "" Then
                        MsgBox ("Para salvar, é necessário preencher todas parcelas.")
                        Exit Sub
                    End If
                Next i
                If CDate(Sheets("Gastos").Range("R1").Value) <> CDate(data_gasto.Value) Then
                    MsgBox ("Data da primeira parcela preenchida não condiz com a data inicial do pagamento.")
                    preencher_botton.SetFocus
                    Exit Sub
                Else
                    Sheets("Gastos").Range("N1").FormulaLocal = "=SOMA(S1:S" & ultima_linha_parcela & ")"
                    soma_parcelas = Sheets("Gastos").Range("N1").Value
                    Sheets("Gastos").Range("N1").Clear
                    If CDec(soma_parcelas) <> CDec(trab.valor_gasto.Value) Then
                        MsgBox ("Soma dos valores das parcelas preenchidas não condiz com o valor total gasto na compra.")
                        preencher_botton.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            'passou por todas condições, então salva
            For i = 1 To parcelas_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 1).Value = Int(id_gasto.Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 2).Value = CDate(Sheets("GAstos").Range("R" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 3).Value = CCur(Sheets("GAstos").Range("S" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 4).Value = Int(Sheets("GAstos").Range("Q" & i).Value)
                Sheets("Gastos").Cells(linha_adc + i - 1, 5).Value = nome_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 6).Value = categoria_gasto.Value
                Sheets("Gastos").Cells(linha_adc + i - 1, 7).Value = forma_gasto.Value
            Next i
        End If
        Sheets("Gastos").Range("Q:T").Clear
    End If
    
    Call reset_novo_gasto
    Call show_lista_gasto
    
    nome_gasto.SetFocus
    
    Exit Sub
    
arrumardata:
    MsgBox ("Data de início do gasto inválida.")
    data_gasto.SetFocus
    Exit Sub
arrumardin:
    MsgBox ("Digite o valor total da compra antes de salvar.")
    valor_gasto.SetFocus
    Exit Sub
arrumarparcelas:
    MsgBox ("Número de parcelas inválido.")
    parcelas_gasto.SetFocus
    Exit Sub
End Sub

Private Sub table_trabalhos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListScroll Me, Me.table_trabalhos
End Sub


Private Sub valor_egasto_Enter()
    If Mid(valor_egasto.Text, 1, 2) <> "R$" Then
        valor_egasto = "R$" + valor_egasto
    End If
    valor_egasto.Value = Format(valor_egasto, "R$#,##0.00")
End Sub

Private Sub valor_gasto_Enter()
    If Mid(valor_gasto.Text, 1, 2) <> "R$" Then
        valor_gasto = "R$" + valor_gasto
    End If
    valor_gasto.Value = Format(valor_gasto, "R$#,##0.00")
End Sub

Private Sub data_gasto_Change()
    Call data(data_gasto)
End Sub

Public Sub novo_botton_Click()
    novo_botton = True
    esconde_edit_remove_despesa.Visible = True
    esconde_registrar_despesa.Visible = False
    trab.MultiPage2.Value = 0
    'trab.lista_gasto.ListIndex = -1
End Sub

Private Sub edit_remove_botton_Click()
    If trab.lista_gasto.ListIndex < 0 Then
        MsgBox ("Necessário clicar duas vezes na linha da lista que deseja editar ou remover.")
        Call novo_botton_Click
        Exit Sub
    End If
    edit_remove_botton = True
    esconde_edit_remove_despesa.Visible = False
    esconde_registrar_despesa.Visible = True
    
    trab.MultiPage2.Value = 1
End Sub


Sub reset_novo_gasto()
    trab.MultiPage2.Value = 0
    trab.lista_gasto.ListIndex = -1
    
    linha = Sheets("Gastos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then
        id_gasto.Value = 1
    Else
        maxID = WorksheetFunction.Max(Sheets("Gastos").Range("A:A"))
        id_gasto.Value = maxID + 1
    End If
    
    data_gasto.Value = Format(Date, "dd/mm/yyyy")
    nome_gasto.Value = "Gasto " & id_gasto.Value
    valor_gasto.Value = ""
    nao_parcelado_botton = True
    
    Call bloqueia_parcelas
    
    categoria_gasto.Value = ""
    forma_gasto.Value = ""
    Call ops_categoria_gasto
    Call ops_forma_gasto
    
End Sub

Sub bloqueia_parcelas()
    parcelas_gasto.Value = 1
    parcelas_gasto.Enabled = False
    preencher_parcelas.Enabled = False
    auto_botton.Enabled = False
    manual_botton.Enabled = False
    pre_visu_parcelas.Enabled = False
    incremento_parcelas.Enabled = False
    preencher_botton.Enabled = False
    
End Sub

Sub ops_categoria_gasto()
    categoria_gasto.Clear
    categoria_egasto.Clear
    linha = Sheets("Gastos").Range("F1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Gastos").Range("F1:F" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Gastos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Gastos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Gastos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              categoria_gasto.AddItem (opc)
              categoria_egasto.AddItem (opc)
        End If
    Next opc
    Sheets("Gastos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub

Sub ops_forma_gasto()
    forma_egasto.Clear
    forma_gasto.Clear
    linha = Sheets("Gastos").Range("G1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Gastos").Range("G1:G" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Gastos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Gastos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Gastos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              forma_gasto.AddItem (opc)
              forma_egasto.AddItem (opc)
        End If
    Next opc
    Sheets("Gastos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub



Private Sub lista_gasto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lista_gasto.ListIndex < 0 Then
        GoTo Sair
    End If
    
    edit_remove_botton_Click
    id_egasto.Value = Int(lista_gasto.List(lista_gasto.ListIndex, 0))
    data_egasto.Value = CDate(lista_gasto.List(lista_gasto.ListIndex, 1))
    valor_egasto.Value = CCur(lista_gasto.List(lista_gasto.ListIndex, 2))
    parcela_egasto.Value = Int(lista_gasto.List(lista_gasto.ListIndex, 3))
    nome_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 4))
    categoria_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 5))
    forma_egasto.Value = (lista_gasto.List(lista_gasto.ListIndex, 6))
    
Sair:
    Exit Sub
End Sub



Public Sub botton_edit_Click() 'Nota-se que é necessario utilizar Public para poder chamar essa sub dentro de outras subs
    If Sheets("Pagamentos").Range("A1000000").End(xlUp).Row = 1 Then
        editar_pagamento.Enabled = False
    Else
        editar_pagamento.Enabled = True
    End If
    
    linha = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then
        botton_novo = True
        Exit Sub
    End If
    
    cont_edit_selected.Visible = True
    novo_registro_selected.Visible = False
    
    botton_remover.Enabled = True
    
    table_trabalhos.ListIndex = linha - 2
    
    ID_registro.Value = Sheets("Trabalhos").Range("A" & linha).Value
    data_ini_registro.Value = CDate(Sheets("Trabalhos").Range("B" & linha).Value)
    data_fim_registro.Value = CDate(Sheets("Trabalhos").Range("C" & linha).Value)
    nome_registro.Value = Sheets("Trabalhos").Range("E" & linha).Value
    link_registro.Value = Sheets("Trabalhos").Range("F" & linha).Value
    cliente_registro.Value = Sheets("Trabalhos").Range("G" & linha).Value
    ctt_cliente_registro.Value = Sheets("Trabalhos").Range("H" & linha).Value
    valor_registro.Value = CCur(Sheets("Trabalhos").Range("I" & linha).Value)
    descobriu_registro.Value = Sheets("Trabalhos").Range("J" & linha).Value
    recomendou_registro.Value = Sheets("Trabalhos").Range("K" & linha).Value
    estilo_registro.Value = Sheets("Trabalhos").Range("L" & linha).Value
    comentario_registro.Value = Sheets("Trabalhos").Range("M" & linha).Value
        
    caixa_hora.Value = Int((Sheets("Trabalhos").Range("D" & linha).Value) / 60)
    caixa_minuto.Value = Int(Sheets("Trabalhos").Cells(linha, 4) Mod 60) 'operador Mod para resto da divisao
    
    id_pagamento.Value = Sheets("Trabalhos").Range("A" & linha).Value
    trabalho_pagamento.Value = Sheets("Trabalhos").Range("E" & linha).Value
    data_pagamento.Enabled = True
    data_pagamento.Value = Date
    salvar_pagamento.Enabled = True
    Call busca_parcela(id_pagamento.Value)
    pagador_pagamento.Enabled = True
    valor_pagamento.Enabled = True
    Call ops_pagador
    valor_pagamento.Value = ""
    pagador_pagamento.Value = ""
End Sub

Public Sub botton_novo_Click()
    cont_edit_selected.Visible = False
    novo_registro_selected.Visible = True
    
    botton_remover.Enabled = False
    table_trabalhos.ListIndex = -1
    id_pagamento.Value = ""
    trabalho_pagamento.Value = ""
    data_pagamento.Value = Date
    parcela_pagamento.Value = ""
    pagador_pagamento.Value = ""
    valor_pagamento.Value = ""
    salvar_pagamento.Enabled = False
    pagador_pagamento.Enabled = False
    data_pagamento.Enabled = False
    valor_pagamento.Enabled = False
    
    Call Reset_for_new
End Sub

Private Sub botton_remover_Click()
    confirmacao_remover.Show
End Sub
Sub removendo()
    linha = table_trabalhos.ListIndex + 2
    
    Sheets("Trabalhos").Range(linha & ":" & linha).Delete Shift:=xlUp
    MsgBox ("Removido com sucesso.")
    botton_remover.Enabled = False
    Call show_table_trabalhos
    Call ops_cliente
    Call ops_recomendacao
    Call ops_estilo
    
    Call remover_regis_pag(ID_registro.Value)
    
    ultima_linha = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row
    If ultima_linha = 1 Then
        botton_novo = True
        Exit Sub
    End If
    
    Call botton_edit_Click
End Sub
Sub remover_regis_pag(id As Integer)
    ultima_linha_pag = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If ultima_linha_pag = 1 Then
        Exit Sub
    End If
    For i = 0 To ultima_linha_pag
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id & ";Pagamentos!A2:A" & ultima_linha_pag & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value + 1 'no da linha que contem o id do trabalho q esta sendo removido
        Sheets("Pagamentos").Range("N1").Clear
        If result <> 1 Then
            Sheets("Pagamentos").Range(result & ":" & result).Delete Shift:=xlUp
        Else
            Exit Sub
        End If
    Next i
    
    If Sheets("Pagamentos").Range("A1000000").End(xlUp).Row = 1 Then
        editar_pagamento.Enabled = False
    Else
        editar_pagamento.Enabled = True
    End If

End Sub

Private Sub cronometro_Click()
    If caixa_hora.Value = "" Then caixa_hora.Value = 0
    If caixa_minuto.Value = "" Then caixa_minuto.Value = 0
    
    If IsNumeric(caixa_hora.Value) And IsNumeric(caixa_minuto.Value) Then
        If caixa_hora.Value >= 0 And caixa_minuto.Value >= 0 Then
            cronometro_janela.Show
        Else
            MsgBox ("Entrada da duração inicial inválida. Coloque a quantidade de horas e minutos antes de começar a cronometrar")
            caixa_hora.SetFocus
        End If
    Else
        MsgBox ("Entrada da duração inicial inválida. Coloque a quantidade de horas e minutos antes de começar a cronometrar")
        caixa_hora.SetFocus
    End If
End Sub

Private Sub data_fim_registro_Change()
    Call data(data_fim_registro)
End Sub

Private Sub data_ini_registro_Change()
    Call data(data_ini_registro)
End Sub

Private Sub data_pagamento_Change()
    Call data(data_pagamento)
    If id_pagamento.Value <> "" Then
        Call busca_parcela(id_pagamento.Value)
    End If
End Sub

Sub editar_pagamento_Click()
    pagamento_janela.Show
End Sub

Private Sub salvar_pagamento_Click()
    If IsDate(data_pagamento.Value) = False Then
        MsgBox ("Preencha corretamente o campo da data de pagamento no formato 'dd/mm/aaaa'.")
        data_pagamento.SetFocus
        Exit Sub
    ElseIf Year(data_pagamento.Value) < 2000 Then
        MsgBox ("Data muito antiga. Corrija a data.")
        data_pagamento.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errodin
    If CCur(valor_pagamento.Value) Then
    End If
    If valor_pagamento.Value < 0 Then
        GoTo errodin
    End If
    
    linha = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If linha > 1 Then
        Sheets("Pagamentos").Range("A2:F" & linha).Cut Sheets("Pagamentos").Range("A3")
        Application.CutCopyMode = False
    End If
    
    Sheets("Pagamentos").Cells(2, 1).Value = Int(id_pagamento.Value)
    Sheets("Pagamentos").Cells(2, 2).Value = trabalho_pagamento.Value
    Sheets("Pagamentos").Cells(2, 3).Value = CDate(data_pagamento.Value)
    Sheets("Pagamentos").Cells(2, 4).Value = CCur(valor_pagamento.Value)
    'Sheets("Pagamentos").Cells(2, 5).Value = Int(parcela_pagamento.Value)
    Sheets("Pagamentos").Cells(2, 6).Value = pagador_pagamento.Value
    MsgBox ("Pagamento registrado com sucesso.")
    
    Call atualiza_parcelas
    
    If botton_edit Then
        Call botton_edit_Click
        Exit Sub
    End If
    
    table_trabalhos.ListIndex = -1
    id_pagamento.Value = ""
    trabalho_pagamento.Value = ""
    data_pagamento.Value = Date
    parcela_pagamento.Value = ""
    pagador_pagamento.Value = ""
    valor_pagamento.Value = ""
    salvar_pagamento.Enabled = False
    pagador_pagamento.Enabled = False
    data_pagamento.Enabled = False
    valor_pagamento.Enabled = False
    
    If Sheets("Pagamentos").Range("A1000000").End(xlUp).Row = 1 Then
        editar_pagamento.Enabled = False
    Else
        editar_pagamento.Enabled = True
    End If
    
    Exit Sub 'é necessario sair da sub antes de entrar no erro
    
errodin:
    MsgBox ("Preencha corretamente o valor recebido.")
    valor_pagamento.SetFocus
    Exit Sub
End Sub

Sub atualiza_parcelas()
    id = id_pagamento.Value
    
    ultima = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If ultima = 1 Then
        Sheets("Pagamentos").Cells(2, 5).Value = Int(parcela_pagamento.Value)
        Exit Sub
    End If
    
    comeco = 3
    For i = 1 To ultima
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value
        Sheets("Pagamentos").Range("N1").Clear
        If result <> 0 And comeco <= ultima Then
            linha = result + (comeco - 1)
            data_encontrada = Sheets("Pagamentos").Range("C" & linha).Value
            Sheets("Pagamentos").Range("O" & (i + 1)).Value = CDate(data_encontrada)
            comeco = linha + 1
        Else
            If i = 1 Then
                Sheets("Pagamentos").Cells(2, 5).Value = Int(parcela_pagamento.Value)
                Exit Sub
            Else
                Sheets("Pagamentos").Range("O1").Value = CDate(data_pagamento.Value)
                Sheets("Pagamentos").Range("P1").FormulaLocal = "=ORDEM.EQ(O1;$O$1:$O$" & i & ";1)+CONT.SE($O$1:O1;O1)-1"
                Sheets("Pagamentos").Range("P1:P" & i).FillDown
                Sheets("Pagamentos").Calculate
                parcela_pagamento.Value = Sheets("Pagamentos").Range("P1").Value
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

End Sub

Private Sub salvar_registro_Click()
    If nome_registro.Value = "" Then
        nome_registro.Value = "Trabalho " & ID_registro.Value
    End If
    
    If cliente_registro.Value = "" Then
        MsgBox ("Preencha o nome do cliente.")
        cliente_registro.SetFocus
        Exit Sub
    End If
    
    If valor_registro.Value = "" Then
        MsgBox ("Preencha o valor cobrado pelo trabalho.")
        valor_registro.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errodin 'forma de se o que estiver embaixo der errado, então vai parra a funçãozinha errodin que fica DEPOIS do Exit Sub e dentro da Sub
    If CCur(valor_registro.Value) Then
    End If
    If valor_registro.Value < 0 Then
        GoTo errodin
    End If
    
    If IsDate(data_ini_registro.Value) = False Then
        MsgBox ("Data inválida. Corrija a data no formato 'dd/mm/aaaa'.")
        data_ini_registro.SetFocus
        Exit Sub
    ElseIf Year(data_ini_registro.Value) < 2000 Then
        MsgBox ("Data muito antiga. Corrija a data.")
        data_ini_registro.SetFocus
        Exit Sub
    End If
    
    If IsDate(data_fim_registro.Value) = False Then
        MsgBox ("Data inválida. Corrija a data no formato 'dd/mm/aaaa'.")
        data_fim_registro.SetFocus
        Exit Sub
    ElseIf Year(data_fim_registro.Value) < 2000 Then
        MsgBox ("Data muito antiga. Corrija a data.")
        data_fim_registro.SetFocus
        Exit Sub
    ElseIf CDate(data_ini_registro.Value) > CDate(data_fim_registro.Value) Then
        MsgBox ("ERRO: Data final antes da data inicial. Corrija as datas.")
        data_fim_registro.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErroDuracao
    'testando entradas validas na duracao
    If Int(caixa_hora.Value) Or Int(caixa_minuto.Value) Then
    End If
    If Int(caixa_minuto.Value) >= 60 Then
        MsgBox ("Preencha corretamente os minutos da duração (até 59min).")
        caixa_minuto.SetFocus
        Exit Sub
    ElseIf Int(caixa_minuto.Value) = 0 And Int(caixa_hora.Value) = 0 Then
        MsgBox ("Preencha o campo da duração.")
        caixa_hora.SetFocus
        Exit Sub
    ElseIf caixa_minuto.Value < 0 Or caixa_hora.Value < 0 Then
        MsgBox ("NOTA: Permitido números positivos apenas.")
    End If
    
    Duracao = caixa_minuto.Value + Int(caixa_hora.Value * 60)
    
    If botton_novo Then
        linha = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row + 1 'primeira linha depois do ultimo registro
        'inserindo na planilha
        Sheets("Trabalhos").Cells(linha, 1).Value = Int(ID_registro.Value)
        Sheets("Trabalhos").Cells(linha, 2).Value = CDate(data_ini_registro.Value)
        Sheets("Trabalhos").Cells(linha, 3).Value = CDate(data_fim_registro.Value)
        Sheets("Trabalhos").Cells(linha, 4).Value = Duracao
        Sheets("Trabalhos").Cells(linha, 5).Value = nome_registro.Value
        Sheets("Trabalhos").Cells(linha, 6).Value = link_registro.Value
        Sheets("Trabalhos").Cells(linha, 7).Value = cliente_registro.Value
        Sheets("Trabalhos").Cells(linha, 8).Value = ctt_cliente_registro.Value
        Sheets("Trabalhos").Cells(linha, 9).Value = CCur(valor_registro.Value)
        Sheets("Trabalhos").Cells(linha, 10).Value = descobriu_registro.Value
        Sheets("Trabalhos").Cells(linha, 11).Value = recomendou_registro.Value
        Sheets("Trabalhos").Cells(linha, 12).Value = estilo_registro.Value
        Sheets("Trabalhos").Cells(linha, 13).Value = comentario_registro.Value
        Call show_table_trabalhos
        Call ops_cliente
        Call ops_recomendacao
        Call ops_estilo
        MsgBox ("Registrado com sucesso.")
        Call Reset_for_new
        nome_registro.SetFocus
        Exit Sub
    End If
    
    
    If botton_edit Then
        confirmacao_editar.Show
    End If
    
    Exit Sub

ErroDuracao:
    MsgBox ("Preencha corretamente a duração.")
    caixa_hora.SetFocus
    Exit Sub
errodin:
    MsgBox ("Preencha corretamente o valor cobrado pelo trabalho.")
    valor_registro.SetFocus
    Exit Sub

End Sub



Private Sub valor_pagamento_Enter()
    If Mid(valor_pagamento.Text, 1, 2) <> "R$" Then
        valor_pagamento = "R$" + valor_pagamento
    End If
    valor_pagamento.Value = Format(valor_pagamento, "R$#,##0.00")
End Sub

Private Sub valor_registro_Enter()
    If Mid(valor_registro.Text, 1, 2) <> "R$" Then
        valor_registro = "R$" + valor_registro
    End If
    valor_registro.Value = Format(valor_registro, "R$#,##0.00")
End Sub

Private Sub link_registro_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If link_registro <> "" Then
        On Error GoTo Mensagem
        ActiveWorkbook.FollowHyperlink link_registro
    End If
    Exit Sub
Mensagem:
    MsgBox ("Não foi possível acessar este link. Verifique se o link está correto. Obs.: Para links da web, colocar ""https://..."".")
    Exit Sub
End Sub

Private Sub descobriu_registro_Change()
    If descobriu_registro.Value = "Recomendação" Then
        recomendou_registro.Enabled = True
    Else
        recomendou_registro.Enabled = False
        recomendou_registro.Value = ""
    End If

End Sub

Private Sub table_trabalhos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Saida
    id_pagamento.Value = table_trabalhos.List(table_trabalhos.ListIndex, 0) 'Note que NESTE CASO o indice da coluna 2 eh 1 (igual python), mas nem sempre eh assim
    trabalho_pagamento.Value = table_trabalhos.List(table_trabalhos.ListIndex, 4)
    data_pagamento.Enabled = True
    data_pagamento.Value = Date
    salvar_pagamento.Enabled = True
    Call busca_parcela(id_pagamento.Value)
    pagador_pagamento.Enabled = True
    valor_pagamento.Enabled = True
    Call ops_pagador

    If botton_edit Then
        botton_remover.Enabled = True
        ID_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 0) 'Note que NESTE CASO o indice da coluna 2 eh 1 (igual python), mas nem sempre eh assim
        data_ini_registro.Value = CDate(table_trabalhos.List(table_trabalhos.ListIndex, 1))
        data_fim_registro.Value = CDate(table_trabalhos.List(table_trabalhos.ListIndex, 2))
        nome_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 4)
        link_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 5)
        cliente_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 6)
        ctt_cliente_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 7)
        valor_registro.Value = CCur(table_trabalhos.List(table_trabalhos.ListIndex, 8))
        descobriu_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 9)
        recomendou_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 10)
        estilo_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 11)
        comentario_registro.Value = table_trabalhos.List(table_trabalhos.ListIndex, 12)
        
        linha = table_trabalhos.ListIndex + 2 'visto que listindex não considera cabeçalho, e começa do 0, soma-se 2 para chegar ao valor da celula
        caixa_hora.Value = Int((Sheets("Trabalhos").Range("D" & linha).Value) / 60)
        caixa_minuto.Value = Int(Sheets("Trabalhos").Cells(linha, 4) Mod 60) 'operador Mod para resto da divisao
    End If

Saida:
    Exit Sub
End Sub

Sub show_table_trabalhos()
    linha = Sheets("Trabalhos").Range("A1048576").End(xlUp).Row
    
    'como se tem apenas uma linha para a condição, então não precisa colocar o "end if"
    If linha = 1 Then linha = 2
    
    trab.table_trabalhos.ColumnCount = 13
    trab.table_trabalhos.ColumnHeads = False
    trab.table_trabalhos.ColumnWidths = "18;60;60;0;78;0;72;0;60;0;0;0;0"
    trab.table_trabalhos.RowSource = "Trabalhos!A2:M" & linha
End Sub
Sub ops_recomendacao()
    recomendou_registro.Clear
    linha = Sheets("Trabalhos").Range("K1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Trabalhos").Range("K1:K" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Trabalhos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Trabalhos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Trabalhos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              recomendou_registro.AddItem (opc)
        End If
    Next opc
    Sheets("Trabalhos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub
Sub ops_estilo()
    estilo_registro.Clear
    linha = Sheets("Trabalhos").Range("L1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Trabalhos").Range("L1:L" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Trabalhos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Trabalhos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Trabalhos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              estilo_registro.AddItem (opc)
        End If
    Next opc
    Sheets("Trabalhos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub
Sub ops_pagador()
    pagador_pagamento.Clear
    linha = Sheets("Pagamentos").Range("F1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Pagamentos").Range("F1:F" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Pagamentos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Pagamentos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Pagamentos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              pagador_pagamento.AddItem (opc)
        End If
    Next opc
    Sheets("Pagamentos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub
Sub ops_cliente()
    cliente_registro.Clear
    linha = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then Exit Sub
    Sheets("Trabalhos").Range("G1:G" & linha).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Trabalhos").Range("N1"), Unique:=True
    linha_unique_opc = Sheets("Trabalhos").Range("N1000000").End(xlUp).Row
    
    For Each opc In Sheets("Trabalhos").Range("N2:N" & linha_unique_opc)
        If opc <> "" Then
              cliente_registro.AddItem (opc)
        End If
    Next opc
    Sheets("Trabalhos").Range("N:N").Clear
    Application.CutCopyMode = False
    
End Sub
Sub data(ByRef data As MSForms.TextBox) 'referenciando a uma caixa de texto

    barra = Mid(data.Text, 3, 1)
    If Len(data.Text) = 2 And barra <> "/" Then
        data = data + "/"
    End If
    
    barra = Mid(data.Text, 5, 1) 'pegar caractere na posiçao 5 (começando do 1) e tamanho 1
    If Len(data.Text) = 5 And barra <> "/" Then
        data = data + "/"
    End If

End Sub
Sub busca_parcela(id As Integer)
    On Error GoTo Sair
    If CDate(data_pagamento.Value) Then
    End If
    
    ultima = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
    If ultima = 1 Then
        parcela_pagamento.Value = Int(1)
        Exit Sub
    End If
    
    comeco = 2
    For i = 1 To ultima
        Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(" & id & ";A" & comeco & ":A" & ultima & ";0);0)"
        result = Sheets("Pagamentos").Range("N1").Value
        Sheets("Pagamentos").Range("N1").Clear
        If result <> 0 And comeco <= ultima Then
            linha = result + (comeco - 1)
            data_encontrada = Sheets("Pagamentos").Range("C" & linha).Value
            Sheets("Pagamentos").Range("N" & (i + 1)).Value = CDate(data_encontrada)
            comeco = linha + 1
        Else
            If i = 1 Then
                parcela_pagamento.Value = Int(1)
                Exit Sub
            Else
                Sheets("Pagamentos").Range("N1").Value = CDate(data_pagamento.Value)
                Sheets("Pagamentos").Range("O1").FormulaLocal = "=ORDEM.EQ(N1;$N$1:$N$" & i & ";1)+CONT.SE($N$1:N1;N1)-1"
                parcela_pagamento.Value = Sheets("Pagamentos").Range("O1").Value
            End If
            Exit For
        End If
    Next i
    
    Sheets("Pagamentos").Range("N1:O" & i).Clear
    
    Exit Sub
Sair:
    parcela_pagamento = ""
    Exit Sub
    
End Sub


Sub Reset_for_new()
    data_ini_registro.Value = Format(Date, "dd/mm/yyyy")
    data_fim_registro.Value = Format(Date, "dd/mm/yyyy")
    caixa_hora.Value = 0
    caixa_minuto.Value = 0
    
    linha = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row
    If linha = 1 Then
        ID_registro.Value = 1
    Else
        maxID = WorksheetFunction.Max(Sheets("Trabalhos").Range("A:A"))
        ID_registro.Value = maxID + 1
    End If
    
    nome_registro.Value = "Trabalho " & ID_registro.Text
    cliente_registro.Value = ""
    valor_registro.Value = ""
    link_registro.Value = ""
    ctt_cliente_registro.Value = ""
    descobriu_registro.Value = ""
    estilo_registro.Value = ""
    comentario_registro.Value = ""
End Sub




Private Sub UserForm_Initialize()
    Application.Visible = False
    'Chama a sub que contém os atributos para habilitar os botões
    'minimizar e maximizar e possibilita redimensionar o UserForm
    Call HabilitaBotoes(trab)
    
    Sheets("Gastos").Range("Q:U").Clear
    Call show_table_trabalhos
    Call show_lista_gasto
    
    descobriu_registro.Clear
    descobriu_registro.AddItem ("Instagram")
    descobriu_registro.AddItem ("Facebook")
    descobriu_registro.AddItem ("Behance")
    descobriu_registro.AddItem ("ArtStation")
    descobriu_registro.AddItem ("TikTok")
    descobriu_registro.AddItem ("Meu site")
    descobriu_registro.AddItem ("Recomendação")
    Call ops_recomendacao
    Call ops_cliente
    Call ops_estilo
    Call botton_novo_Click
    If Sheets("Pagamentos").Range("A1000000").End(xlUp).Row = 1 Then
        editar_pagamento.Enabled = False
    Else
        editar_pagamento.Enabled = True
    End If
    
    Call reset_novo_gasto
    Call tregistros_frame1_Click
    
    Set CelulaCaminhoImagem = Sheets("Planilha1").Range("A1")
    Set ArquivoExiste = CreateObject("Scripting.FileSystemObject")
    
    If ArquivoExiste.FileExists(CelulaCaminhoImagem.Value) = True Then
        Foto.Picture = LoadPicture(CelulaCaminhoImagem.Value)
        Foto.PictureSizeMode = fmPictureSizeModeZoom
    End If
    
    nome_real.Caption = Sheets("Planilha1").Range("A2").Value
    
End Sub


Private Sub UserForm_Terminate()
    ThisWorkbook.Save
    ThisWorkbook.Application.Quit
End Sub

