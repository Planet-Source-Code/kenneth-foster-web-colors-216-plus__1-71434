VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "216 Web Colors"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Make Your Own Color"
      Height          =   1185
      Left            =   9840
      TabIndex        =   29
      Top             =   7275
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Make Your Own"
      Height          =   2325
      Left            =   10590
      TabIndex        =   20
      Top             =   6150
      Visible         =   0   'False
      Width           =   1635
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RGB/HEX"
         Height          =   225
         Left            =   150
         TabIndex        =   30
         Top             =   2025
         Width           =   1350
      End
      Begin VB.VScrollBar vsBlue 
         Height          =   1500
         Left            =   75
         Max             =   255
         TabIndex        =   25
         Top             =   465
         Width           =   255
      End
      Begin VB.VScrollBar vsGreen 
         Height          =   1500
         Left            =   450
         Max             =   255
         TabIndex        =   23
         Top             =   465
         Width           =   255
      End
      Begin VB.VScrollBar vsRed 
         Height          =   1500
         Left            =   825
         Max             =   255
         TabIndex        =   21
         Top             =   465
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HEX"
         Height          =   210
         Left            =   1125
         TabIndex        =   31
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1425
         Left            =   1110
         TabIndex        =   27
         Top             =   540
         Width           =   465
      End
      Begin VB.Label lblBlue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   26
         Top             =   225
         Width           =   330
      End
      Begin VB.Label lblGreen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   420
         TabIndex        =   24
         Top             =   225
         Width           =   330
      End
      Begin VB.Label lblRed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   795
         TabIndex        =   22
         Top             =   225
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Blend Colors"
      Height          =   1200
      Left            =   6975
      TabIndex        =   11
      Top             =   7275
      Width           =   2835
      Begin VB.HScrollBar hsBlend 
         Height          =   195
         Left            =   90
         Max             =   100
         TabIndex        =   17
         Top             =   930
         Width           =   2130
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2265
         TabIndex        =   19
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RGB(0,0,0)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1245
         TabIndex        =   18
         Top             =   585
         Width           =   1485
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1245
         TabIndex        =   16
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Base2   Color"
         Height          =   360
         Left            =   720
         TabIndex        =   15
         Top             =   465
         Width           =   510
      End
      Begin VB.Label lblBase2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   690
         TabIndex        =   14
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Base1  Color"
         Height          =   390
         Left            =   150
         TabIndex        =   13
         Top             =   465
         Width           =   420
      End
      Begin VB.Label lblBase1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   12
         Top             =   210
         Width           =   510
      End
   End
   Begin VB.HScrollBar HS1 
      Height          =   210
      Left            =   2700
      Max             =   100
      Min             =   -100
      TabIndex        =   9
      Top             =   7530
      Width           =   2400
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Color/Value on Mouse Over"
      Height          =   270
      Left            =   4635
      TabIndex        =   8
      Top             =   8205
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Color Values"
      Height          =   255
      Left            =   4635
      TabIndex        =   7
      Top             =   7920
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "216 Web Colors"
      Enabled         =   0   'False
      Height          =   405
      Left            =   150
      TabIndex        =   6
      Top             =   8010
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RGB Base Colors"
      Height          =   405
      Left            =   2040
      TabIndex        =   5
      Top             =   8010
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send RGB to Clipboard"
      Height          =   570
      Left            =   5220
      TabIndex        =   1
      Top             =   7305
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   7335
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "For information Only."
      Height          =   240
      Left            =   7995
      TabIndex        =   28
      Top             =   6075
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      Height          =   6075
      Left            =   8010
      Top             =   0
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   8010
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   3390
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   8010
      Picture         =   "Form1.frx":2747
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Brigthness: 0"
      Height          =   210
      Left            =   3270
      TabIndex        =   10
      Top             =   7305
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   240
      Left            =   645
      Top             =   8565
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   4
      Top             =   7665
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1695
      TabIndex        =   3
      Top             =   7335
      Width           =   975
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Ken Foster 2008
'216 Web Colors plus

'Prevents flicker
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Dim R As Long
Dim G As Long
Dim B As Long
Dim toggle As Boolean      'keeps track of which color chart is visible
Dim MYO As Boolean         'determines which color to brighten
Dim TemCol As String        'stores a non-chart color
Dim colID As Integer         'stores index of chosen chart color

Private Sub Form_Load()
   'size form
   Form1.Width = 12525
   Form1.Height = 9010
   
   'build first color chart
   BuildMatrix216
   DrawColors216
   
   Shape1.top = 3480
   Shape1.left = 30
   Shape1.Width = 7500
   Shape1.Height = 225
   Shape1.BorderWidth = 2
   Shape1.Visible = False
   
   Check1.Value = 1           'show color values
   toggle = False                'chart 1 showing
End Sub

Private Sub BuildMatrix216()
   Dim x As Integer
   Dim y As Integer
   
   'set parameters for label on form
   lblColor(0).left = 195
   lblColor(0).top = 225
   lblColor(0).Alignment = 2
   lblColor(0).Width = 975
   lblColor(0).Height = 350
   
   For x = 1 To 215
      Load lblColor(x)   'create the control array
      'load the first half of labels
      If x < 108 Then
         With lblColor(x)
            .top = lblColor(x - 1).top
            .Width = lblColor(x - 1).Width
            .Height = lblColor(x - 1).Height
            .left = (lblColor(x - 1).left + lblColor(x - 1).Width) + 25
            .Visible = True
         End With
         
         y = y + 1
         If y = 6 Then  'start a new row
         lblColor(x).top = (lblColor(x - 1).top + lblColor(x - 1).Height) + 25
         lblColor(x).left = lblColor(0).left
         y = 0
      End If
   End If
   'set postion of first label on right side
   If x = 108 Then
      With lblColor(108)
         .top = 225
         .Width = lblColor(x - 1).Width
         .Height = lblColor(x - 1).Height
         .left = 6255
         .Visible = True
      End With
      y = 0
   End If
   'load the last half of the labels
   If x > 108 Then
      With lblColor(x)
         .top = lblColor(x - 1).top
         .Width = lblColor(x - 1).Width
         .Height = lblColor(x - 1).Height
         .left = (lblColor(x - 1).left + lblColor(x - 1).Width) + 25
         .Visible = True
      End With
      
      y = y + 1
      If y = 6 Then    'start a new row
      lblColor(x).top = (lblColor(x - 1).top + lblColor(x - 1).Height) + 25
      lblColor(x).left = lblColor(108).left
      y = 0
   End If
End If
Next x
End Sub

Private Sub DrawColors216()
   Dim x As Integer
   Dim y As Integer
   
   Dim R As String
   Dim G As String
   Dim B As String
   Dim gt As Integer
   
   gt = 255
   R = 255
   G = 255
   B = 255
   y = 0
   
   For x = 0 To 215
      lblColor(x).Alignment = 2   'center
      lblColor(x).ForeColor = vbBlack
      'red value
      If x = 36 Then R = 204
      If x = 72 Then R = 153
      If x = 108 Then R = 102
      If x > 107 Then lblColor(x).ForeColor = vbWhite   'change font color on last half of labels
      If x = 144 Then R = 51
      If x = 180 Then R = 0
      
      If y = 5 Then
         gt = gt - 51
         If gt < 0 Then gt = 255
         B = 255 - (y * 51)   'blue value for last label in row
         y = 0   'reset y counter
      Else
         G = gt      'green value
         B = 255 - (y * 51)  'blue value
         y = y + 1
      End If
      
      lblColor(x).BackColor = RGB(R, G, B)   'set label backcolor
      If Check1.Value = 1 Then
         lblColor(x).Caption = R & "," & G & "," & B   'show color value
      Else
         lblColor(x).Caption = ""
      End If
   Next x
End Sub

Private Sub BuildMatrixRGB()
   Dim x As Integer
   Dim y As Integer
   Dim z As Integer
   
   'set parameters for label on form
   lblColor(0).left = 195
   lblColor(0).top = 225
   lblColor(0).Alignment = 2
   lblColor(0).Width = 1000
   lblColor(0).Height = 225
   z = 0
   For x = 1 To 195
      Load lblColor(x)   'create the control array
      
      With lblColor(x)
         .top = (lblColor(x - 1).top + lblColor(x - 1).Height) + 25
         .Width = lblColor(x - 1).Width
         .Height = lblColor(x - 1).Height
         .left = lblColor(x - 1).left
         .Caption = ""
         .Visible = True
      End With
      
      y = y + 1
      
      If x = (z * 28) + 28 Then
         y = 0
         With lblColor((z * 28) + 28)
            .top = lblColor(z * 28).top
            .Width = lblColor(x - 1).Width
            .Height = lblColor(x - 1).Height
            .left = lblColor((z * 28) + 28).left + lblColor(x - 1).Width + 25
            .Visible = True
         End With
         z = z + 1
      End If
   Next x
   
End Sub

Private Sub DrawColorsRGB()
   Dim x As Integer
   Dim y As Integer
   Dim z As Integer
   
   For x = 0 To 196
   
      'red
      If x < 15 Then
         lblColor(x).BackColor = RGB(34 + (y * 17), 0, 0)
         lblColor(13).BackColor = RGB(255, 0, 0)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 13 And x < 28 Then
         lblColor(x).BackColor = RGB(255, 17 + (y * 17), 17 + (y * 17))
         lblColor(x).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      
      'blue
      If x > 27 And x < 43 Then
         lblColor(x).BackColor = RGB(0, 0, 34 + (y * 17))
         lblColor(41).BackColor = RGB(0, 0, 255)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 41 And x < 56 Then
         lblColor(x).BackColor = RGB(17 + (y * 17), 17 + (y * 17), 255)
         lblColor(x).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      
      'green
      If x > 55 And x < 70 Then
         lblColor(x).BackColor = RGB(0, 34 + (y * 17), 0)
         lblColor(69).BackColor = RGB(0, 255, 0)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 69 And x < 84 Then
         lblColor(x).BackColor = RGB(17 + (y * 17), 255, 17 + (y * 17))
         lblColor(x).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      
      'yellow
      If x > 83 And x < 98 Then
         lblColor(x).BackColor = RGB(34 + (y * 17), 34 + (y * 17), 0)
         lblColor(97).BackColor = RGB(255, 255, 0)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 97 And x < 112 Then
         lblColor(x).BackColor = RGB(255, 255, 17 + (y * 17))
         lblColor(x).ForeColor = vbBlack
         lblColor(97).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      
      'turquise
      If x > 111 And x < 126 Then
         lblColor(x).BackColor = RGB(0, 34 + (y * 17), 34 + (y * 17))
         lblColor(125).BackColor = RGB(0, 255, 255)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 125 And x < 140 Then
         lblColor(x).BackColor = RGB(17 + (y * 17), 255, 255)
         lblColor(x).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      
      'purple
      If x > 139 And x < 154 Then
         lblColor(x).BackColor = RGB(34 + (y * 17), 0, 34 + (y * 17))
         lblColor(153).BackColor = RGB(255, 0, 255)
         lblColor(x).ForeColor = vbWhite
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      If x > 153 And x < 168 Then
         lblColor(x).BackColor = RGB(255, 17 + (y * 17), 255)
         lblColor(x).ForeColor = vbBlack
         If Check1.Value = 1 Then
            lblColor(x).Caption = lblColor(x).BackColor
         Else
            lblColor(x).Caption = ""
         End If
      End If
      y = y + 1
      If y = 14 Then y = 0
      
   Next x
   y = 0
   
   'black/white
   For z = 168 To 195
      lblColor(z).ForeColor = vbWhite
      lblColor(168).BackColor = RGB(0, 0, 0)
      lblColor(z).BackColor = RGB(9 + (y * 9), 9 + (y * 9), 9 + (y * 9))
      lblColor(195).BackColor = RGB(255, 255, 255)
      If z > 181 Then lblColor(z).ForeColor = vbBlack
      If Check1.Value = 1 Then
         lblColor(z).Caption = lblColor(z).BackColor
      Else
         lblColor(z).Caption = ""
      End If
      y = y + 1
   Next z
End Sub

Private Sub HS1_Change()
   HS1_Scroll
End Sub

Private Sub HS1_Scroll()
   If MYO = False Then
      Label1.BackColor = AdjustBrightness(lblColor(colID).BackColor, HS1.Value)
   Else
       Label1.BackColor = AdjustBrightness(TemCol, HS1.Value)
   End If
   
   Label3.Caption = "Brightness: " & HS1.Value
   GetRGB Label1.BackColor, R, G, B
   Text1.Text = "RGB(" & R & "," & G & "," & B & ")"
   Label2.Caption = Hex(Label1.BackColor)
   Label2.Caption = String(6 - Len(Label2.Caption), "0") & Label2.Caption
End Sub

Private Sub hsBlend_Change()
   hsBlend_Scroll
End Sub

Private Sub hsBlend_Scroll()
    Label6.BackColor = BlendColors(lblBase1.BackColor, lblBase2.BackColor, hsBlend.Value)
    GetRGB Label6.BackColor, R, G, B
    Label7.Caption = "RGB(" & R & "," & G & "," & B & ")"
    Label8.Caption = hsBlend.Value
End Sub

Private Sub Label10_Click()
   MYO = True
   Label1.BackColor = Label10.BackColor
   TemCol = Label10.BackColor
   HS1.Value = 0
   Label3.Caption = "Brightness: 0"
   GetRGB Label1.BackColor, R, G, B
   Text1.Text = "RGB(" & R & "," & G & "," & B & ")"
   Label2.Caption = Hex(Label1.BackColor)
   Label2.Caption = String(6 - Len(Label2.Caption), "0") & Label2.Caption
End Sub

Private Sub Label6_Click()
   MYO = True
   Label1.BackColor = Label6.BackColor
   TemCol = Label6.BackColor
   HS1.Value = 0
   Label3.Caption = "Brightness: 0"
   GetRGB Label1.BackColor, R, G, B
   Text1.Text = "RGB(" & R & "," & G & "," & B & ")"
   Label2.Caption = Hex(Label1.BackColor)
   Label2.Caption = String(6 - Len(Label2.Caption), "0") & Label2.Caption
End Sub

Private Sub lblBase1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then lblBase1.BackColor = Label1.BackColor
   hsBlend.Value = 0
End Sub

Private Sub lblBase2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then lblBase2.BackColor = Label1.BackColor
   hsBlend.Value = 0
End Sub

Private Sub lblColor_Click(Index As Integer)
   MYO = False
   Label1.BackColor = lblColor(Index).BackColor
   Label2.Caption = Hex(lblColor(Index).BackColor)
   Label2.Caption = String(6 - Len(Label2.Caption), "0") & Label2.Caption
   GetRGB lblColor(Index).BackColor, R, G, B
   Text1.Text = "RGB(" & R & "," & G & "," & B & ")"
   colID = Index
   TemCol = lblColor(Index).BackColor
   HS1.Value = 0
   Label3.Caption = "Brightness: 0"
   If toggle = True Then
      DrawColorsRGB
   Else
      DrawColors216
   End If
      If R > 204 And G < 153 And B <= 204 Then
         lblColor(Index).ForeColor = vbBlue
      Else
         lblColor(Index).ForeColor = vbRed
      End If
End Sub

Private Sub lblColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Check2.Value = 1 Then                        'if Color/Value on Mouse Over is checked
      Label1.BackColor = lblColor(Index).BackColor
      Label2.Caption = Hex(lblColor(Index).BackColor)
      Label2.Caption = String(6 - Len(Label2.Caption), "0") & Label2.Caption
      GetRGB lblColor(Index).BackColor, R, G, B
      Text1.Text = "RGB(" & R & "," & G & "," & B & ")"
   End If
End Sub

Private Sub Check1_Click()
   If toggle = False Then
      DrawColors216
   Else
      DrawColorsRGB
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 0 Then
      lblBlue.left = 45
      vsBlue.left = 75
      lblRed.left = 795
      vsRed.left = 825
      lblBlue.Caption = Hex(vsBlue.Value)
      lblGreen.Caption = Hex(vsGreen.Value)
      lblRed.Caption = Hex(vsRed.Value)
      Label11.Caption = "HEX"
   Else
      lblBlue.left = 795
      vsBlue.left = 825
      lblRed.left = 45
      vsRed.left = 75
      lblBlue.Caption = vsBlue.Value
      lblGreen.Caption = vsGreen.Value
      lblRed.Caption = vsRed.Value
      Label11.Caption = "RGB"
   End If
End Sub

Private Sub Command1_Click()
   Clipboard.Clear
   Clipboard.SetText Text1.Text
End Sub

Private Sub Command2_Click()
   Dim x As Integer
   LockWindowUpdate Form1.hWnd

   For x = 1 To 215
      Unload lblColor(x)
   Next x
   BuildMatrixRGB
   DrawColorsRGB
   colID = 195
   TemCol = vbWhite
   Command3.Enabled = True
   Command2.Enabled = False
   Shape1.Visible = True
   Text1.Text = ""
   Label2.Caption = ""
   Label1.BackColor = vbWhite
   Form1.Caption = "RGB Based Colors"
   toggle = True
   HS1.Value = 0
   Label3.Caption = "Brightness: 0"
   Image1.Visible = True
   Image2.Visible = True
   Label9.Visible = True
   Shape2.Visible = True
   LockWindowUpdate 0&
   MYO = True
End Sub

Private Sub Command3_Click()
   Dim x As Integer
   LockWindowUpdate Form1.hWnd
   
   For x = 1 To 195
      Unload lblColor(x)
   Next x
   BuildMatrix216
   DrawColors216
   colID = 0
   TemCol = vbWhite
   Command3.Enabled = False
   Command2.Enabled = True
   Shape1.Visible = False
   Text1.Text = ""
   Label2.Caption = ""
   Label1.BackColor = vbWhite
   Form1.Caption = "216 Web Colors"
   toggle = False
   HS1.Value = 0
   Label3.Caption = "Brightness: 0"
   Image1.Visible = False
   Image2.Visible = False
   Label9.Visible = False
   Shape2.Visible = False
   LockWindowUpdate 0&
End Sub

Private Sub Command5_Click()
   Frame2.Visible = Not Frame2.Visible
End Sub

Private Sub GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
   R = LngCol Mod 256    'Red
   G = (LngCol And vbGreen) / 256 'Green
   B = (LngCol And vbBlue) / 65536 'Blue
End Sub

Private Sub vsBlue_Change()
   vsBlue_Scroll
End Sub

Private Sub vsBlue_Scroll()
   If Check3.Value = 0 Then
      lblBlue.Caption = Hex(vsBlue.Value)
   Else
      lblBlue.Caption = vsBlue.Value
   End If
   Label10.BackColor = RGB(vsRed.Value, vsGreen.Value, vsBlue.Value)
End Sub

Private Sub vsGreen_Change()
   vsGreen_Scroll
End Sub

Private Sub vsGreen_Scroll()
   If Check3.Value = 0 Then
      lblGreen.Caption = Hex(vsGreen.Value)
   Else
      lblGreen.Caption = vsGreen.Value
   End If
   Label10.BackColor = RGB(vsRed.Value, vsGreen.Value, vsBlue.Value)
End Sub

Private Sub vsRed_Change()
   vsRed_Scroll
End Sub

Private Sub vsRed_Scroll()
   If Check3.Value = 0 Then
      lblRed.Caption = Hex(vsRed.Value)
   Else
      lblRed.Caption = vsRed.Value
   End If
   Label10.BackColor = RGB(vsRed.Value, vsGreen.Value, vsBlue.Value)
End Sub
