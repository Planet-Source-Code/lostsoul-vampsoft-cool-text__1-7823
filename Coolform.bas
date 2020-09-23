Attribute VB_Name = "Module1"
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function Sendformssage Lib "user32" Alias "SendformssageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Sub FormDrag(TheForm As Form)
       ReleaseCapture
       Call Sendformssage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Public Function FormEffect(Form As Form, xRed As Integer, xGreen As Integer, xBlue As Integer, X As Boolean, Min As Boolean) As Integer
With Form
    r2 = xRed + 50: If r2 > 255 Then r2 = 255
    g2 = xGreen + 50: If g2 > 255 Then g2 = 255
    b2 = xBlue + 50: If b2 > 255 Then b2 = 255
    r3 = xRed - 50: If r3 < 0 Then r3 = 0
    g3 = xGreen - 50: If g3 < 0 Then g3 = 0
    b3 = xBlue - 50: If b3 < 0 Then b3 = 0
    Form.BackColor = RGB(xRed, xGreen, xBlue)
    Form.Line (60, 60)-(.Width - 80, 60 + 210), RGB(r3, g3, b3), BF
    Form.Line (15, 15)-(Form.Width - 30, Form.Height - 30), RGB(r2, g2, b2), B
    Form.Line (15, Form.Height - 30)-(Form.Width - 30, Form.Height - 30), RGB(r3, g3, b3)
    Form.Line (Form.Width - 30, Form.Height - 30)-(Form.Width - 30, 30), RGB(r3, g3, b3)
    Form.Line (60, 60 + 255)-(Form.Width - 60, 60 + 255), RGB(r3, g3, b3)
    Form.Line (30, 60 + 270)-(Form.Width - 30, 60 + 270), RGB(r2, g2, b2)

    Form.CurrentX = 120
    Form.CurrentY = 60
    Form.ForeColor = RGB(r2, g2, b2)
    Form.Print Form.Caption
    
    If X = True Then
        Form.Line (.Width - 275, 70)-(.Width - 90, 255), RGB(r2, g2, b2), BF
        Form.CurrentX = .Width - 230
        Form.CurrentY = 60
        Form.ForeColor = RGB(r3, g3, b3)
        Form.Print "X"
    End If
    
    If Min = True Then
        Form.Line (.Width - 485, 70)-(.Width - 300, 255), RGB(r2, g2, b2), BF
        Form.CurrentX = .Width - 420
        Form.CurrentY = 60
        Form.ForeColor = RGB(r3, g3, b3)
        Form.Print "_"
    End If
    
    For i = 0 To 5
        Form.Label1(i).BackColor = RGB(xRed, xGreen, xBlue)
        Form.Label1(i).ForeColor = RGB(r3, g3, b3)
    Next i
End With
End Function

