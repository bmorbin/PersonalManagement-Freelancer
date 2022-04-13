Attribute VB_Name = "botoes_guia"

'Função que retornará o nome da classe e o nome do UserForm
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Função que recupera as informações sobre o nome da classe e o estilo da janela do UserForm
Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

 


'Função que altera o estilo da janela do UserForm
Private Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Sub que irá obter o nome do UserForm (ObjForm)
Sub HabilitaBotoes(ObjForm As Object)

    'Código que atribui os botões minimizar e maximizar e possibilita redimensionar o UserForm
    SetWindowLong FindWindow("ThunderDFrame", ObjForm.Caption), -16, _
    GetWindowLong(FindWindow("ThunderDFrame", ObjForm.Caption), -16) Or &H20000 'Or &H10000 Or &H40000
    'H10000 serve para maximizar formulario
    'H20000 serve para minimizar formulario
    'H40000 serve para redimensionar formulario

End Sub

