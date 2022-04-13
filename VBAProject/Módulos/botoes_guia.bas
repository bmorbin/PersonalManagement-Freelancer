Attribute VB_Name = "botoes_guia"

'Fun��o que retornar� o nome da classe e o nome do UserForm
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Fun��o que recupera as informa��es sobre o nome da classe e o estilo da janela do UserForm
Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

 


'Fun��o que altera o estilo da janela do UserForm
Private Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Sub que ir� obter o nome do UserForm (ObjForm)
Sub HabilitaBotoes(ObjForm As Object)

    'C�digo que atribui os bot�es minimizar e maximizar e possibilita redimensionar o UserForm
    SetWindowLong FindWindow("ThunderDFrame", ObjForm.Caption), -16, _
    GetWindowLong(FindWindow("ThunderDFrame", ObjForm.Caption), -16) Or &H20000 'Or &H10000 Or &H40000
    'H10000 serve para maximizar formulario
    'H20000 serve para minimizar formulario
    'H40000 serve para redimensionar formulario

End Sub

