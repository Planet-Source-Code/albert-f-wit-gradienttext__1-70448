VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   270
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   270
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call GradientText(Picture1, "ColoredTextDemo", RGB(255, 128, 0), RGB(0, 128, 255))
End Sub

Private Sub GradientText(Picture As PictureBox, ByVal Text$, ByVal tStart&, ByVal tEnd&)
    
    Dim Number%, I%, Red1%, Green1%, Blue1%, Red2%, Green2%, Blue2%, dRed!, dGreen!, dBlue!
    Number = Len(Text)
    getRGB tStart, Red1, Green1, Blue1
    getRGB tEnd, Red2, Green2, Blue2
    dRed = (Red2 - Red1) / Number
    dGreen = (Green2 - Green1) / Number
    dBlue = (Blue2 - Blue1) / Number
    Picture.Cls
    Picture.CurrentX = 0
    Picture.CurrentY = 0
    Picture.Width = Picture.TextWidth(Text & "   ")
    Picture.Height = Picture.TextHeight(Text)
    For I = 1 To Number
        Picture.ForeColor = RGB(Red1, Green1, Blue1)
        Picture.Print Mid(Text, I, 1);
        Red1 = Red1 + dRed
        Green1 = Green1 + dGreen
        Blue1 = Blue1 + dBlue
    Next I
End Sub

Private Sub getRGB(ByVal gColor&, Red%, Green%, Blue%)
    
    Red = gColor And &HFF
    Green = (gColor \ &H100) And &HFF
    Blue = (gColor \ &H10000) And &HFF
End Sub


