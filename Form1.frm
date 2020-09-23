VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skyler Lyon's Imaging Effects 2.0"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   840
      ScaleHeight     =   7215
      ScaleWidth      =   5535
      TabIndex        =   39
      Top             =   8200
      Width           =   5535
   End
   Begin VB.CommandButton cmdPointalism 
      Caption         =   "Pointalism"
      Height          =   375
      Left            =   8160
      TabIndex        =   38
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPixilated 
      Caption         =   "Pixilated"
      Height          =   375
      Left            =   6960
      TabIndex        =   37
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmboss 
      Caption         =   "Emboss"
      Height          =   375
      Left            =   5760
      TabIndex        =   36
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSilhuette 
      Caption         =   "Silhuette"
      Height          =   375
      Left            =   8160
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlatten2 
      Caption         =   "Flatten2"
      Height          =   375
      Left            =   6960
      TabIndex        =   34
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncoherence 
      Caption         =   "Incoherence"
      Height          =   375
      Left            =   5760
      TabIndex        =   33
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdVividNeg 
      Caption         =   "Vivid Negation"
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNegative 
      Caption         =   "Negative"
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNeonNegation 
      Caption         =   "Neon Negation"
      Height          =   375
      Left            =   5760
      TabIndex        =   30
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdWarped 
      Caption         =   "Warped"
      Height          =   375
      Left            =   8160
      TabIndex        =   29
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetMain 
      Caption         =   "Set To Main Image"
      Height          =   375
      Left            =   9480
      TabIndex        =   26
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   13920
      TabIndex        =   25
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   7440
      Width           =   1095
   End
   Begin VB.PictureBox imgNo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   6000
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   23
      Top             =   8040
      Width           =   135
   End
   Begin VB.CommandButton cmdSilk 
      Caption         =   "Silk"
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlatten 
      Caption         =   "Flatten"
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdVividSilk 
      Caption         =   "VividSilk"
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSunShinnyDay 
      Caption         =   "SunShinnyDay"
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGlowInTheDark 
      Caption         =   "GlowInDark"
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseRGB 
      Caption         =   "DecreaseRGB"
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseBlue 
      Caption         =   "DecreaseBlue"
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseGreen 
      Caption         =   "DecreaseG"
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseRed 
      Caption         =   "DecreaseRed"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseRGB 
      Caption         =   "IncreaseRGB"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseBlue 
      Caption         =   "IncreaseBlue"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseGreen 
      Caption         =   "IncreaseGreen"
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseRed 
      Caption         =   "IncreaseRed"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoGreen 
      Caption         =   "NoGreen"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoBlue 
      Caption         =   "NoBlue"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoRed 
      Caption         =   "NoRed"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlur2 
      Caption         =   "Blur More"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlur1 
      Caption         =   "Blur"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDarken 
      Caption         =   "Darken"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLighten 
      Caption         =   "Lighten"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreyScale 
      Caption         =   "GrayScale"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   9480
      ScaleHeight     =   7185
      ScaleWidth      =   5505
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7185
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   6600
      TabIndex        =   28
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label lblT 
      Caption         =   "Time (MS):"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   7440
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:           frmMain
'Version:         2.0
'Date:            10/07/2008
'Author:          Skyler Lyon
'Copyright:       © 2008 Skyler Lyon
'Description:     Main window.

Option Explicit

Private Sub cmdBlur1_Click()
lblTime.Caption = Blur(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdBlur2_Click()
lblTime.Caption = BlurMore(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdDarken_Click()
lblTime.Caption = Darken(Picture1, picTemp, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdDecreaseBlue_Click()
lblTime.Caption = IncreaseBlue(Picture1, picTemp, -50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdDecreaseGreen_Click()
lblTime.Caption = IncreaseGreen(Picture1, picTemp, -50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdDecreaseRed_Click()
lblTime.Caption = IncreaseRed(Picture1, picTemp, -50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdDecreaseRGB_Click()
lblTime.Caption = IncreaseRGB(Picture1, picTemp, -50, -100, -50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdEmboss_Click()
lblTime.Caption = Emboss(Picture1, picTemp, 100)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdFlatten_Click()
lblTime.Caption = Flatten(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdFlatten2_Click()
lblTime.Caption = Flatten2(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdGlowInTheDark_Click()
lblTime.Caption = GlowInTheDark(Picture1, picTemp, 1)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdGreyScale_Click()
lblTime.Caption = GrayScale(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdIncoherence_Click()
lblTime.Caption = Incoherence(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdIncreaseBlue_Click()
lblTime.Caption = IncreaseBlue(Picture1, picTemp, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdIncreaseGreen_Click()
lblTime.Caption = IncreaseGreen(Picture1, picTemp, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdIncreaseRed_Click()
lblTime.Caption = IncreaseRed(Picture1, picTemp, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdIncreaseRGB_Click()
lblTime.Caption = IncreaseRGB(Picture1, picTemp, 50, 100, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdLighten_Click()
lblTime.Caption = Lighten(Picture1, picTemp, 50)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdLoadImage_Click()
Dim FileName As String

FileName = fncGetFileNametoOpen("Open Image File", "All Files|*.*", "")
If IsValidFile(FileName) = False Then
    Exit Sub
End If

On Error GoTo No_Open

Open FileName For Input As #1
Picture1.Picture = LoadPicture(FileName)
Close 1
Exit Sub
No_Open:
Resume ExitLine
ExitLine:
Exit Sub
End Sub

Private Sub cmdNegative_Click()
lblTime.Caption = Negative(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdNeonNegation_Click()
lblTime.Caption = NeonNegation(Picture1, picTemp, 255)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdNoBlue_Click()
lblTime.Caption = NoBlue(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdNoGreen_Click()
lblTime.Caption = NoGreen(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdNoRed_Click()
lblTime.Caption = NoRed(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdPixilated_Click()
lblTime.Caption = Pixilated(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdPointalism_Click()
lblTime.Caption = Pointalism(Picture1, picTemp, 2)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdSave_Click()
Dim FileName As String

FileName = fncGetFileNametoSave("Bitmap Files|*.bmp", "", "Save Bitmap Image File")
If IsValidFile(FileName) = False Then
    Exit Sub
End If

On Error GoTo No_Save

Open FileName For Output As #2
SavePicture Picture2.Picture, FileName
Close 2
Exit Sub
No_Save:
Resume ExitLine
ExitLine:
Exit Sub
End Sub

Private Sub cmdSetMain_Click()
Set Picture1.Picture = Picture2.Picture
End Sub

Private Sub cmdSilhuette_Click()
lblTime.Caption = Silhuette(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdSilk_Click()
lblTime.Caption = Silk(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdSunShinnyDay_Click()
lblTime.Caption = GlowInTheDark(Picture1, picTemp, 10)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdVividNeg_Click()
lblTime.Caption = VividNegation(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdVividSilk_Click()
lblTime.Caption = VividSilk(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub

Private Sub cmdWarped_Click()
lblTime.Caption = Warped(Picture1, picTemp)
Picture2.Picture = picTemp.Image
End Sub
