Attribute VB_Name = "auto_click"
'Criado por Leonardo Mezavilla - leodecm3@gmail.com
'
'


Private Const SW_SHOWMAXIMIZED = 3
Private Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Type POINTAPI
X As Long
Y As Long
End Type
 
Public Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "User32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10



Sub rodaLinhaSelecionada()


If Cells(Selection.Row, 1) = "moveMouse" Then
    SetCursorPos Cells(Selection.Row, 2), Cells(Selection.Row, 3)
Else
    MsgBox "Botao secreto. seleciona uma linha que tenha o comando moveMouse." & Chr(10) & Chr(10) & Chr(10) & "esse botao eh util para quando vc se pergunta como diabos eu vou saber onde fica 234x453 "
End If


End Sub


Sub vai()
'Criado por Leonardo Mezavilla - leodecm3@gmail.com
'
'
'iniciaa variaveis
Dim pt As POINTAPI
fim = Sheets("Script").Range("a64000").End(xlUp).Row




'ActiveWindow.WindowState = xlMinimized



'cria o txt que se apagado a macro trava
Open Application.ActiveWorkbook.Path & "\delete para parar.txt" For Append As #1
Close #1
Dim fso
Dim file As String
file = Application.ActiveWorkbook.Path & "\delete para parar.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(file) Then Exit Sub

'pinta de branco a coluna D
Sheets("Script").Columns("D:D").Interior.Pattern = xlNone



For i = 10 To fim
    
    
    'verifica se o arquivo txt existe, se nao ele interrompe (para nao ficar rodando sem parara)
    If Not fso.FileExists(file) Then Exit Sub
    comando = Sheets("Script").Range("a1").Offset(i, 0).Value
    X = Sheets("Script").Range("b1").Offset(i, 0).Value
    Y = Sheets("Script").Range("c1").Offset(i, 0).Value
    
    
    'coloca im x na linha que ta rodando
    Sheets("Script").Columns("D:D").ClearContents
    Sheets("Script").Range("d1").Offset(i, 0).Value = "x"
    

    
    
    x_cor = Sheets("Script").Range("b1").Offset(i, 0).Value
    y_cor = Sheets("Script").Range("c1").Offset(i, 0).Value
    
    
    If comando = "" Then GoTo nextfor1 'nada
    
    
    If comando = "wait colour" Then
        tmp_espera_cor = espera_cor(X, Y)
        If tmp_espera_cor <> "ok" Then
            Sheets("Script").Range("d1").Offset(i, 0).Value = ".....Travou nessa linha porque nao encontrava a cor. eu pintei essa celula da cor que eu encontrei no lugar. cod." & tmp_espera_cor
            Sheets("Script").Range("d1").Offset(i, 0).Interior.Color = RGB(ExtractR(tmp_espera_cor), ExtractG(tmp_espera_cor), ExtractB(tmp_espera_cor))
            Exit Sub
        End If
        GoTo nextfor1
    End If
    
    
    
    If comando = "pause" Then 'pause
        If X > 60 Then
            MsgBox "comando Pause maior que 60 segundos, estou enterrompendo a macro"
            Exit Sub
        End If
        Application.Wait (Now() + TimeSerial(0, 0, X))
        GoTo nextfor1
    End If
    
    If comando = "fim" Then
        Exit Sub 'fim
    End If
    
    If comando = "shell" Then 'shell
        Shell (X)
        GoTo nextfor1
    End If
    
    If comando = "press" Then 'pressionar tecla
        Application.SendKeys X
        Application.Wait (Now() + TimeSerial(0, 0, 1)) ' aqui eu espero um pouco para nao travar, se vc tiver fazendo algo ilegal (bot de jogo) , coloque um tempo aleatorios
        GoTo nextfor1
    End If
    
    If comando = "click" Then 'clicar
        Application.Wait (Now() + TimeSerial(0, 0, 1))
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        GoTo nextfor1
    End If
    
    
    If comando = "moveMouse" Then  'mover cursor
        pt.X = X
        pt.Y = Y
        SetCursorPos pt.X, pt.Y
        GoTo nextfor1
    End If
    
    If comando = "LOOP-de-ate" Then 'volta para a linha 10
        'X = Sheets("Script").Range("b1").Offset(i, 0).Value
        'Y = Sheets("Script").Range("c1").Offset(i, 0).Value
        If X >= Y Then
            'vida que segue
            GoTo nextfor1
        Else
            'volta para o comeco
            Sheets("Script").Range("b1").Offset(i, 0).Value = Sheets("Script").Range("b1").Offset(i, 0).Value + 1
            i = 9
            GoTo nextfor1
        End If
    End If
    
    Stop ' deu merda, o parametro nao era esperado!!!
nextfor1:
Next

'Y = Nothing
'X = Nothing

End Sub







Sub MoveMouse(X As Single, Y As Single)
Dim pt As POINTAPI
pt.X = X
pt.Y = Y
SetCursorPos pt.X, pt.Y
End Sub




Private Sub DoubleClick()   'NAO TESTADO
  'Simulate a double click as a quick series of two clicks
  SetCursorPos 100, 100 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub RightClick()   'NAO TESTADO
  'Simulate a right click
  SetCursorPos 200, 200 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub






