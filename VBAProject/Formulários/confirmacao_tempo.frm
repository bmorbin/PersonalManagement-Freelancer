VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirmacao_tempo 
   Caption         =   "Confirmar Tempo Cronometrado"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "confirmacao_tempo.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirmacao_tempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub nao_Click()
    Call Unload(confirmacao_tempo)
End Sub

Private Sub sim_Click()
    Call cronometro_janela.fechar
    Call Unload(confirmacao_tempo)
End Sub

