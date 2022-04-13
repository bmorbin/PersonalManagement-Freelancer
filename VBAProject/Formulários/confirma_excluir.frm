VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirma_excluir 
   Caption         =   "Confimar Remoção"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "confirma_excluir.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirma_excluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelar_Click()
    Call Unload(confirma_excluir)
End Sub

Private Sub remover_Click()
    Call pagamento_janela.remover_pag
    Call Unload(confirma_excluir)
    Exit Sub
End Sub
