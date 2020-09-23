Attribute VB_Name = "modColorFunctions"
Option Explicit
'This code by KRYO
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function InvertRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
    End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
    End Type
    Private Const PS_SOLID = 0
    'This function will brighten or darken a
    '     color
    'Example: Picture1.BackColor = AdjustBri
    '     ghtness(Picture1.BackColor, -50)

Public Function AdjustBrightness(ByVal Color As Long, ByVal Amount As Single) As Long
    On Error Resume Next
    
    Dim R(1) As Integer, G(1) As Integer, B(1) As Integer
    
    'get red, green, and blue values
    GetRGB R(0), G(0), B(0), Color
    
    'add/subtract the amount to/from the ori
    '     ginal RGB values
    R(1) = SetBound(R(0) + Amount, 0, 255)
    G(1) = SetBound(G(0) + Amount, 0, 255)
    B(1) = SetBound(B(0) + Amount, 0, 255)
    
    'convert RGB back to Long value
    AdjustBrightness = RGB(R(1), G(1), B(1))
End Function
'This function will blend two colors tog
'     ether at any percentage 0 - 100
'Example: Picture1.BackColor = BlendColo
'     rs(vbRed, vbBlue, 50)

Public Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Percentage As Single) As Long
    On Error Resume Next
    
    Dim R(2) As Integer, G(2) As Integer, B(2) As Integer
    Dim fPercentage(2) As Single
    Dim DAmt(2) As Single
    
    'make sure Percentage is between 0 and 1
    '     00
    Percentage = SetBound(Percentage, 0, 100)
    
    'extract the RGB values from Color1 and
    '     Color2
    GetRGB R(0), G(0), B(0), Color1
    GetRGB R(1), G(1), B(1), Color2
    
    '1st part: get the positive or negative
    '     amount between the 2 colors
    '2nd part: calculate how much needs to b
    '     e added to Color1
    '(Difference divided by 100 multiplied b
    '     y the percentage)
    DAmt(0) = R(1) - R(0): fPercentage(0) = (DAmt(0) / 100) * Percentage
    DAmt(1) = G(1) - G(0): fPercentage(1) = (DAmt(1) / 100) * Percentage
    DAmt(2) = B(1) - B(0): fPercentage(2) = (DAmt(2) / 100) * Percentage
    
    'add/subtract each percentage to RGB val
    '     ues
    R(2) = R(0) + fPercentage(0)
    G(2) = G(0) + fPercentage(1)
    B(2) = B(0) + fPercentage(2)
    
    'convert RGB back to Long value
    BlendColors = RGB(R(2), G(2), B(2))
End Function
'This will draw Verticle/Horizontal grad
'     ient very quickly
'Example: DrawGradient Picture1.hDC, 0,
'     0, Picture1.ScaleWidth, Picture1.ScaleHe
'     ight, vbRed, vbBlue, True
'(note: if the picture1 is set to autore
'     draw it must be refreshed after this fun
'     ction)

Public Sub GetRGB(R As Integer, G As Integer, B As Integer, ByVal Color As Long)
    Dim TempValue As Long
    
    'First translate the color from a long v
    '     alue to a short value
    TranslateColor Color, 0, TempValue
    
    'Calculate the red, green, and blue valu
    '     es from the short value
    R = TempValue And &HFF&
    G = (TempValue And &HFF00&) / 2 ^ 8
    B = (TempValue And &HFF0000) / 2 ^ 16
End Sub
'Invert colors (Negative image)
'Example: InvertColor Picture1.hDC, 0, 0
'     , Picture1.ScaleWidth, Picture1.ScaleHei
'     ght

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single


    If Num < MinNum Then
        'if less that min value make it the min
        '     value
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        'if more than max value make it the max
        '     value
        SetBound = MaxNum
    Else
        'if between min and max then leave it al
        '     one
        SetBound = Num
    End If
End Function

