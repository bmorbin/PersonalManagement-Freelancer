VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cronometro_janela 
   Caption         =   "Cronômetro"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "cronometro_janela.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cronometro_janela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Time_Start As Double
Public Tempo As Double
Public Time_End As Double
Public linha As Integer
Public tempo_formatado As String
Public status As String
Public horas As Double
Public minutos As Double
Public Sub_End As Double
Public Sub_Start As Double


Private Sub cronometro_interroga_selected_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MsgBox ("Tempo cronometrado será adicionado à duração na aba ""Registros"". Ao abrir esta aba do cronômetro, este é iniciado automaticamente.")
    cronometro_interroga_normal.Visible = True
    cronometro_interroga_selected.Visible = False
End Sub

Private Sub fundo_crono_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cronometro_interroga_normal.Visible = True
    cronometro_interroga_selected.Visible = False
End Sub

Private Sub cronometro_interroga_normal_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cronometro_interroga_normal.Visible = False
    cronometro_interroga_selected.Visible = True
End Sub

Public Sub continuar_Click()
    Time_Start = Now
    continuar.Enabled = False
    pausar.Enabled = True
    status = "Continuar"
    Call atualiza_registro_cronometro
    linha = linha + 1
    pausar.SetFocus
End Sub


Public Sub encerrar_Click()
    
    If status = "Pausar" Then
        Sub_End = 0
        Sub_Start = 0
    ElseIf status = "Continuar" Then
        Sub_End = Now
        Sub_Start = Time_Start
    ElseIf status = "Iniciar" Then
        Sub_End = Now
        Sub_Start = Time_Start
    End If
    Tempo = Tempo + (Sub_End - Sub_Start)
    status = "Encerrar"
    Call atualiza_registro_cronometro
    linha = linha + 1
    
    MsgBox ("Tempo Cronometrado: " & tempo_formatado)
    
    confirmacao_tempo.Show

    Call Unload(cronometro_janela) 'fechar esse form
End Sub
Sub fechar()
    Call atualiza_duracao
End Sub



Public Sub atualiza_duracao()
    new_min = Int(trab.caixa_minuto.Value) + Int(minutos)
    h_adc = 0
    If new_min >= 60 Then
        h_adc = Int(new_min / 60)
        new_min = Int(new_min Mod 60)
    End If
    new_h = Int(trab.caixa_hora.Value) + Int(horas) + h_adc

    trab.caixa_hora.Value = new_h
    trab.caixa_minuto.Value = new_min
    trab.data_fim_registro.Value = CDate(Date)
End Sub



Public Sub pausar_Click()
    Time_End = Now
    Tempo = Tempo + (Time_End - Time_Start)
    status = "Pausar"
    Call atualiza_registro_cronometro
    linha = linha + 1
    continuar.Enabled = True
    pausar.Enabled = False
    continuar.SetFocus
    
End Sub

Private Sub registro_cronometro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListScroll Me, Me.registro_cronometro
End Sub

Public Sub UserForm_Initialize()
    
    linha = 2
    registro_cronometro.ColumnCount = 4
    registro_cronometro.ColumnHeads = False
    registro_cronometro.ColumnWidths = "72;60;60;54"
    
    pausar.Enabled = True
    continuar.Enabled = False
    Tempo = 0
    Time_Start = Now
    
    status = "Iniciar"
    Call atualiza_registro_cronometro
    linha = linha + 1
    
End Sub

Sub atualiza_registro_cronometro()
    Sheets("Cronometro").Range("A" & linha) = Date
    Sheets("Cronometro").Range("B" & linha) = Format(Time, "hh:mm:ss")
    Call formata_tempo
    Sheets("Cronometro").Range("C" & linha) = tempo_formatado
    Sheets("Cronometro").Range("D" & linha) = status
    registro_cronometro.RowSource = "Cronometro!A2:D" & linha
End Sub

Public Function formata_tempo() As String
    horas = Tempo * 24
    minutos = (horas - Int(horas)) * 60
    segundos = (minutos - Int(minutos)) * 60
    tempo_formatado = Int(horas) & "h" & Int(minutos) & "min" & Int(segundos) & "s"
End Function

Private Sub UserForm_Terminate()
    Sheets("Cronometro").Range("A2:D" & linha).Clear
End Sub
