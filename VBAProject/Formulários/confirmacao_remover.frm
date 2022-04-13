VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirmacao_remover 
   Caption         =   "Confirmar Remoção"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "confirmacao_remover.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirmacao_remover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call trab.removendo
    trab.botton_remover.Enabled = False
    Call trab.botton_edit_Click
    Call Unload(confirmacao_remover)
End Sub

Private Sub CommandButton2_Click()
    Call Unload(confirmacao_remover)
End Sub
