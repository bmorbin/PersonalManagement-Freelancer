VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirmacao_editar 
   Caption         =   "Confirmar Edição"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "confirmacao_editar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirmacao_editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelar_Click()
    Call Unload(confirmacao_editar)
End Sub

Private Sub editar_Click()
    Duracao = trab.caixa_minuto.Value + Int(trab.caixa_hora.Value * 60)
    ultima = Sheets("Trabalhos").Range("A1000000").End(xlUp).Row
    
    Sheets("Trabalhos").Range("N1").FormulaLocal = "=CORRESP(" & trab.ID_registro.Value & ";A2:A" & ultima _
    & ";0)"
    linha = Sheets("Trabalhos").Range("N1").Value + 1
    Sheets("Trabalhos").Range("N1").Clear
    
    arg = CInt(linha) 'linha é double e para converter, é necessario usar CInt
    Call update_pag_nome((arg))
    
    'inserindo na planilha
    Sheets("Trabalhos").Cells(linha, 1).Value = Int(trab.ID_registro.Value)
    Sheets("Trabalhos").Cells(linha, 2).Value = CDate(trab.data_ini_registro.Value)
    Sheets("Trabalhos").Cells(linha, 3).Value = CDate(trab.data_fim_registro.Value)
    Sheets("Trabalhos").Cells(linha, 4).Value = Duracao
    Sheets("Trabalhos").Cells(linha, 5).Value = trab.nome_registro.Value
    Sheets("Trabalhos").Cells(linha, 6).Value = trab.link_registro.Value
    Sheets("Trabalhos").Cells(linha, 7).Value = trab.cliente_registro.Value
    Sheets("Trabalhos").Cells(linha, 8).Value = trab.ctt_cliente_registro.Value
    Sheets("Trabalhos").Cells(linha, 9).Value = CCur(trab.valor_registro.Value)
    Sheets("Trabalhos").Cells(linha, 10).Value = trab.descobriu_registro.Value
    Sheets("Trabalhos").Cells(linha, 11).Value = trab.recomendou_registro.Value
    Sheets("Trabalhos").Cells(linha, 12).Value = trab.estilo_registro.Value
    Sheets("Trabalhos").Cells(linha, 13).Value = trab.comentario_registro.Value
    Call trab.show_table_trabalhos
    Call trab.ops_cliente
    Call trab.ops_recomendacao
    Call trab.ops_estilo
    MsgBox ("Editado com sucesso.")
    Call trab.botton_edit_Click
    trab.nome_registro.SetFocus
    Call Unload(confirmacao_editar)
End Sub

Sub update_pag_nome(linha As Integer)
    nome_antigo = Sheets("Trabalhos").Range("E" & linha).Value
    If nome_antigo <> trab.nome_registro.Value Then
        ultima_linha_pag = Sheets("Pagamentos").Range("A1000000").End(xlUp).Row
        If ultima_linha_pag = 1 Then
            Exit Sub
        End If
                    
        For i = 0 To ultima_linha_pag
            Sheets("Pagamentos").Range("N1").FormulaLocal = "=SEERRO(CORRESP(""" & nome_antigo & """;Pagamentos!B2:B" & ultima_linha_pag & ";0);0)"
            result = Sheets("Pagamentos").Range("N1").Value + 1 'no da linha que contem o nome de trabalho q esta sendo editado
            Sheets("Pagamentos").Range("N1").Clear
            If result <> 1 Then
                If Int(Sheets("Pagamentos").Range("B2:B" & ultima_linha_pag).Find(nome_antigo, Lookat:=xlWhole) _
                .Offset(0, -1).Value) = Int(trab.ID_registro.Value) Then
                'necessario colocar lookat whole(inteiro) para achar correspondencia exata. Fonte: https://docs.microsoft.com/pt-br/office/vba/api/excel.range.find
                    Sheets("Pagamentos").Cells(result, 2).Value = trab.nome_registro.Value
                End If
            Else
                Exit Sub
            End If
        Next i
    End If
End Sub
