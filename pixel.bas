Attribute VB_Name = "pixel"
'Criado por Leonardo Mezavilla - leodecm3@gmail.com
'
'

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As POINT) As Long
Private Declare Function GetWindowDC Lib "User32" (ByVal hwnd As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Long


Private Type POINT
    X As Long
    Y As Long
End Type





 
Function espera_cor(segundos, cor)
espera_cor = "ok"
Dim pLocation As POINT
Dim lColour, lDC As Long
lDC = GetWindowDC(0)
Call GetCursorPos(pLocation)
X = pLocation.X
Y = pLocation.Y


For i = 1 To segundos '  segundos de espera
    tmpCor = GetPixel(lDC, X, Y)
    If cor = tmpCor Then
        Exit Function
    End If
    Application.Wait Now + TimeValue("00:00:01")
    
    'Sleep 1000  'wait 0.5 seconds
Next

MsgBox "esperei por " & segundos & " segundos pela cor na posicao " & X & ":" & Y & "  e ela nao apareceu..."
espera_cor = tmpCor

End Function








Sub record()

segundos = InputBox("TUTORIAL: Aperte `x` ao inves de digitar ,  clique por + de 0.5seg." & Chr(10) & Chr(10) & Chr(10) & "Algumas perguntas: " & Chr(10) & Chr(10) & "->>> Gravo por quantos segundos? (MAXIMO 59)" & Chr(10) & "- Grava cor dos botoes? (sim/nao)", , 8)
If segundos >= 60 Or segundos <= 0 Then
    MsgBox "segundos entre 0 a 59 , " & segundos & " eh invalido. ESTOU SAINDO SEM FAZER NADA"
    Exit Sub
End If

cor = InputBox("TUTORIAL: Aperte `x` ao inves de digitar ,  clique por + de 0.5seg." & Chr(10) & Chr(10) & Chr(10) & "Algumas perguntas: " & Chr(10) & Chr(10) & "- Gravo por quantos segundos? (MAXIMO 59)" & Chr(10) & "->>> Grava cor dos botoes clicados? (sim/nao)", , "sim")
If cor = "sim" Then
    cor = True
Else
    cor = False
End If


temp_lColour = 0
temp_X = 0
temp_Y = 0




f = Now() + TimeSerial(0, 0, segundos)
fim = Sheets("Script").Range("a64000").End(xlUp).Row
Dim pLocation As POINT
Dim lColour, lDC As Long


'espera 1 seg antes de comecar
Application.Wait (Now() + TimeSerial(0, 0, 1))


'this loop will stop when the time above run out
Do While Now() < f

    'verifica se apertou X
    If (GetAsyncKeyState(vbKeyX)) Then
        If Sheets("Script").Range("a1").Offset(fim - 1, 0).Value = "press" Then
            'n faco nada, se nao ele vai passar 1 segundo colocando press em todas as linhas (1600 linhas)
        Else
            'na linha de baixo
            fim = fim + 1
            Sheets("Script").Range("a1").Offset(fim - 1, 0).Value = "press"
            Sheets("Script").Range("b1").Offset(fim - 1, 0).Value = "x"
            Sheets("Script").Range("c1").Offset(fim - 1, 0).Value = "-"
        End If
    End If


    'verifica se apertou mouse
    If (GetAsyncKeyState(vbKeyLButton)) Then
    
        lDC = GetWindowDC(0)
        Call GetCursorPos(pLocation)
        lColour = GetPixel(lDC, pLocation.X, pLocation.Y)
        X = pLocation.X
        Y = pLocation.Y
        
        'verifica para nao mandar o mesmo comando duas vezes seguidas
        variacao = Abs(temp_X - X) + Abs(temp_Y - Y) + Abs(temp_lColour - lColour)
            
        If variacao > 20 Then
            'verifica para nao mandar o mesmo comando duas vezes seguidas
            temp_lColour = lColour
            temp_X = X
            temp_Y = Y
            
            'na linha de baixo
            fim = fim + 1
            Sheets("Script").Range("a1").Offset(fim - 1, 0).Value = "moveMouse"
            Sheets("Script").Range("b1").Offset(fim - 1, 0).Value = X
            Sheets("Script").Range("c1").Offset(fim - 1, 0).Value = Y
            
            'se tiver com cor
            If cor = True Then
                'na linha de baixo
                fim = fim + 1
                Sheets("Script").Range("a1").Offset(fim - 1, 0).Value = "wait colour"
                Sheets("Script").Range("b1").Offset(fim - 1, 0).Value = 5
                Sheets("Script").Range("c1").Offset(fim - 1, 0).Value = lColour
                Sheets("Script").Range("c1").Offset(fim - 1, 0).Interior.Color = RGB(ExtractR(lColour), ExtractG(lColour), ExtractB(lColour))
            End If
            
            'na linha de baixo
            fim = fim + 1
            Sheets("Script").Range("a1").Offset(fim - 1, 0).Value = "click"
            Sheets("Script").Range("b1").Offset(fim - 1, 0).Value = "-"
            Sheets("Script").Range("c1").Offset(fim - 1, 0).Value = "-"


        End If
    End If


Loop



MsgBox "fim da gravacao"
Application.Visible = True
End Sub












Public Function ExtractR(ByVal CurrentColor As Long) As Byte
   ExtractR = CurrentColor And 255
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Byte
   ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Byte
   ExtractB = (CurrentColor \ 65536) And 255
End Function

