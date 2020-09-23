Attribute VB_Name = "mSubclass"
' APIs '
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const GWL_STYLE = (-16)
Const ES_PASSWORD = &H20&
Const EM_SETPASSWORDCHAR& = &HCC
Const RDW_INVALIDATE = &H1
Const WM_GETTEXT = &HD

Dim Buffer As String
Public retHwnd As Long
'
' Funcion para subclasficar la ventana
' Subclassification function
'
Function EnumWinProc(ByVal hwnd As Long, lpData As Long) As Long
    
    If GetWindowLong(hwnd, GWL_STYLE) Then
        ' Si la ventana es de estilo password...
        If (GetWindowLong(hwnd, GWL_STYLE) And ES_PASSWORD) Then
            ' Obtengo el texto de la ventana que contiene la contraseña
            SendMessage hwnd, WM_GETTEXT, Len(Buffer), ByVal Buffer
            Debug.Print "Password: " & Buffer
            DoEvents
            ' Envio el mensaje para quitar la mascara de "*" a la ventana
            SendMessage hwnd, EM_SETPASSWORDCHAR&, 0&, 0&
            ' Redibujo la ventana para que los cambios tengan efecto
            RedrawWindow hwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE
        End If
    End If

    ' Continua enumerando
    EnumWinProc = True
End Function

