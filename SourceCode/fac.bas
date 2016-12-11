Attribute VB_Name = "Module1"
Option Explicit







Public A, B, C, A1, B1, C1, M, P, S As Currency
Public Const PI = 3.141592654
Public Type Tri
    Bx As Currency
    cx As Currency
    cy As Currency
End Type
Public Type Fx
    x As Currency
    y As Currency
End Type
Public Function DTR(ByVal x As Currency) As Currency
DTR = x / 180 * PI
End Function
Public Function RTD(ByVal x As Currency) As Currency
RTD = x / PI * 180
End Function
Public Function Judge(ByVal D As Currency, ByVal E As Currency, ByVal F As Currency) As Boolean
If (D + E > F) And (D + F > E) And (E + F > D) And A > 0 And B > 0 And C > 0 Then
    Judge = True
Else
    Judge = False
End If
End Function
Public Function ASin(ByVal x As Currency) As Currency
Select Case x
  Case 1: ASin = PI / 2
  Case -1: ASin = -PI / 2
  Case Else: ASin = Atn(x / Sqr(1 - x * x))
End Select
End Function

Public Function ACos(ByVal x As Currency) As Currency
  Select Case x
  Case 1: ACos = 0
  Case -1: ACos = PI
  Case Else: ACos = PI / 2 - Atn(x / Sqr(1 - x * x))
End Select
End Function

Public Function Paint(ByVal Z As Currency, ByVal y As Currency, ByVal x As Currency) As Tri
On Error GoTo 1:
    Paint.Bx = Z
    Paint.cx = (y ^ 2 + Z ^ 2 - x ^ 2) / (2 * Z)
    Paint.cy = y * Sin(ACos((y ^ 2 + Z ^ 2 - x ^ 2) / (2 * y * Z)))
1:
Exit Function
End Function

Public Function DTR2(ByVal x As Single) As Single
DTR2 = x / 180 * PI
End Function
Public Function RTD2(ByVal x As Single) As Single
RTD2 = x / PI * 180
End Function

Public Function Style(ByVal x As Currency, ByVal y As Currency, ByVal Z As Currency) As Integer

If (x = 0) And (y = 0) And (Z = 0) Then
    Style = 0
   Exit Function
End If

If x = y Or y = Z Or Z = x Then
    If x = y And y = Z Then
        Style = 3
        Exit Function
    Else
        
        If (CCur(Format(x ^ 2 + y ^ 2, "0.##")) = CCur(Format(Z ^ 2, "0.##"))) Or (CCur(Format(y ^ 2 + Z ^ 2, "0.##")) = CCur(Format(x ^ 2, "0.##"))) Or (CCur(Format(x ^ 2 + Z ^ 2, "0.##")) = CCur(Format(y ^ 2, "0.##"))) Then
            Style = 1
            Exit Function
        Else
            Style = 4
            Exit Function
        End If
    End If
End If


If (CCur(Format(x ^ 2 + y ^ 2, "0.##")) = CCur(Format(Z ^ 2, "0.##"))) Or (CCur(Format(y ^ 2 + Z ^ 2, "0.##")) = CCur(Format(x ^ 2, "0.##"))) Or (CCur(Format(x ^ 2 + Z ^ 2, "0.##")) = CCur(Format(y ^ 2, "0.##"))) Then
        
       Style = 2
        Exit Function
Else
   If x <> 0 Or y <> 0 Or Z <> 0 Then
        Style = 5
        Exit Function
   End If
End If






End Function




