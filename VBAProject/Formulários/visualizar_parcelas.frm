VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} visualizar_parcelas 
   Caption         =   "Pré-Visualização Parcelas"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3225
   OleObjectBlob   =   "visualizar_parcelas.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "visualizar_parcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    visualizar_parcelas.visu_parcelas_list.ColumnCount = 3
    visualizar_parcelas.visu_parcelas_list.ColumnHeads = False
    visualizar_parcelas.visu_parcelas_list.ColumnWidths = "36;60;54"
    visualizar_parcelas.visu_parcelas_list.RowSource = "Gastos!Q1:T" & trab.parcelas_gasto.Value
End Sub

Private Sub visu_parcelas_list_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListScroll Me, Me.visu_parcelas_list
End Sub
