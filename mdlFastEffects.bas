Attribute VB_Name = "mdlFastEffects"
'Title:           mdlFastEffects
'Version:         2.0
'Date:            10/07/2008
'Author:          Skyler Lyon
'Copyright:       © 2008 Skyler Lyon
'Description:     Module for unique and more complex image effects.

Option Explicit

Public Function Setup(PicSrc As PictureBox, PicDest As PictureBox)
Dim m_curPerformanceFrequency As Currency

'Obtain speed of system
Call QueryPerformanceFrequency(m_curPerformanceFrequency)
m_dblPerformanceFrequency = CDbl(m_curPerformanceFrequency)

'Make a copy
m_p32Original() = GetPictureArrayInv(PicSrc.Picture)
m_p32Output() = m_p32Original()

'Clear Picture
Set PicDest.Picture = Nothing

'Obtain Width and height
m_lngWidth = UBound(m_p32Original, 1) + 1
m_lngHeight = UBound(m_p32Original, 2) + 1

'Set Dest picture to right size (use this is if your using a temp image like I did
'   otherwise comment it out)
PicDest.Width = m_lngWidth * 15
PicDest.Height = m_lngHeight * 15

'Zero output
ReDim m_p32Output(0 To m_lngWidth - 1, 0 To m_lngHeight - 1)

'Begin timer
m_lngStartTime = GetTimeMS
End Function

Public Function ReturnPicAndTime(PicDest As PictureBox) As Long
'Return Time
m_lngEndTime = GetTimeMS
ReturnPicAndTime = m_lngEndTime - m_lngStartTime

'Paint Image
CopyPixelsToDC PicDest.hDC, m_p32Output
End Function

Public Function Setup2(PicSrc As PictureBox, PicDest As PictureBox)
Dim m_curPerformanceFrequency As Currency

'Obtain speed of system
Call QueryPerformanceFrequency(m_curPerformanceFrequency)
m_dblPerformanceFrequency = CDbl(m_curPerformanceFrequency)

'Clear Picture
Set PicDest.Picture = Nothing

'Begin timer
m_lngStartTime = GetTimeMS
End Function

Public Function ReturnTime() As Long
'Return Time
m_lngEndTime = GetTimeMS
ReturnTime = m_lngEndTime - m_lngStartTime
End Function

Public Function GrayScale(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim intGrayScale As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        intGrayScale = (R + G + B) / 3
        m_p32Output(X, Y).Red = intGrayScale
        m_p32Output(X, Y).Green = intGrayScale
        m_p32Output(X, Y).Blue = intGrayScale
    Next X
Next Y

GrayScale = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Negative(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        m_p32Output(X, Y).Red = 255 - m_p32Original(X, Y).Red
        m_p32Output(X, Y).Green = 255 - m_p32Original(X, Y).Green
        m_p32Output(X, Y).Blue = 255 - m_p32Original(X, Y).Blue
    Next X
Next Y

Negative = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Incoherence(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        m_p32Output(X, Y).Red = m_p32Original(X, Y).Blue
        m_p32Output(X, Y).Green = m_p32Original(X, Y).Green
        m_p32Output(X, Y).Blue = m_p32Original(X, Y).Red
    Next X
Next Y

Incoherence = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Lighten(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Long) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        R = R + Magnitude
        G = G + Magnitude
        B = B + Magnitude
        
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Lighten = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Darken(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Long) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        R = R - Magnitude
        G = G - Magnitude
        B = B - Magnitude
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Darken = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Blur(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        If Y = 0 Then
            R2 = R1
            G2 = G1
            B2 = B1
        End If
        
        R = (R1 + R2) / 2
        G = (G1 + G2) / 2
        B = (B1 + B2) / 2
        
        R2 = R1
        G2 = G1
        B2 = B1
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Blur = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function BlurMore(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim R3 As Integer, G3 As Integer, B3 As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        If Y = 0 Then
            R3 = R1
            G3 = G1
            B3 = B1
            
            R2 = R1
            G2 = G1
            B2 = B1
        End If
        
        R = (((R1 + R2) / 2) + R3) / 2
        G = (((G1 + G2) / 2) + G3) / 2
        B = (((B1 + B2) / 2) + B3) / 2
        
        R3 = R2
        G3 = G2
        B3 = B2
        R2 = R1
        G2 = G1
        B2 = B1
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

BlurMore = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

'Mag: 1.0 - 10.0
Public Function GlowInTheDark(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Single) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        G2 = Abs(((255 * 3) / G)) * Magnitude
        B2 = Abs(((255 * 3) / B)) * Magnitude
        R2 = Abs(((255 * 3) / R)) * Magnitude
        
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0
        If R2 > 255 Then R2 = 255
        If G2 > 255 Then G2 = 255
        If B2 > 255 Then B2 = 255
        
        m_p32Output(X, Y).Red = R2
        m_p32Output(X, Y).Green = G2
        m_p32Output(X, Y).Blue = B2
    Next X
Next Y

GlowInTheDark = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Silk(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        If Y = 0 Then
            R2 = R1
            G2 = G1
            B2 = B1
        End If

        R = Abs(255 - ((R1 * R2) / 10))
        G = Abs(255 - ((G1 * G2) / 10))
        B = Abs(255 - ((B1 * B2) / 10))
        
        R2 = R1
        G2 = G1
        B2 = B1
        
        If R > 255 Then
            m_p32Output(X, Y).Red = 255
        Else
            If R < 0 Then
                m_p32Output(X, Y).Red = 0
            Else
                m_p32Output(X, Y).Red = R
            End If
        End If
        
        If B > 255 Then
            m_p32Output(X, Y).Blue = 255
        Else
            If B < 0 Then
                m_p32Output(X, Y).Blue = 0
            Else
                m_p32Output(X, Y).Blue = B
            End If
        End If
        
        If G > 255 Then
            m_p32Output(X, Y).Green = 255
        Else
            If G < 0 Then
                m_p32Output(X, Y).Green = 0
            Else
                m_p32Output(X, Y).Green = G
            End If
        End If
        
        'm_p32Output(X, Y).Red = R
        'm_p32Output(X, Y).Blue = B
        'm_p32Output(X, Y).Green = G
    Next X
Next Y

Silk = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function VividSilk(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

'On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        'old version (103ms)
'        R2 = (Abs((255 - (R + (R * 2))) * 2) / 5) * 3
'        G2 = (Abs((255 - (G + (G * 2))) * 2) / 5) * 3
'        B2 = (Abs((255 - (B + (B * 2))) * 2) / 5) * 3
        
        'fastest version (94ms)
        R2 = Abs((255 - (R + R + R)) * 2) * 0.6
        G2 = Abs((255 - (G + G + G)) * 2) * 0.6
        B2 = Abs((255 - (B + B + B)) * 2) * 0.6
        
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0
        If R2 > 255 Then R2 = 255
        If G2 > 255 Then G2 = 255
        If B2 > 255 Then B2 = 255
        
        m_p32Output(X, Y).Red = R2
        m_p32Output(X, Y).Green = G2
        m_p32Output(X, Y).Blue = B2
    Next X
Next Y

VividSilk = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Flatten(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim R3 As Integer, G3 As Integer, B3 As Integer
Dim R4 As Integer, G4 As Integer, B4 As Integer

IsProcessing = True

'On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 4 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y - 3).Red
        G1 = m_p32Original(X, Y - 3).Green
        B1 = m_p32Original(X, Y - 3).Blue
        R2 = m_p32Original(X, Y - 2).Red
        G2 = m_p32Original(X, Y - 2).Green
        B2 = m_p32Original(X, Y - 2).Blue
        R3 = m_p32Original(X, Y - 1).Red
        G3 = m_p32Original(X, Y - 1).Green
        B3 = m_p32Original(X, Y - 1).Blue
        R4 = m_p32Original(X, Y).Red
        G4 = m_p32Original(X, Y).Green
        B4 = m_p32Original(X, Y).Blue
        
        R1 = (R1 - R3) * 2.5
        R2 = (R2 - R2) * 2.5
        R3 = (R3 - R1) * 2.5
        R4 = (R4 - R4) * 2.5
        G1 = (G1 - G3) * 2.5
        G2 = (G2 - G2) * 2.5
        G3 = (G3 - G1) * 2.5
        G4 = (G4 - G4) * 2.5
        B1 = (B1 - B3) * 2.5
        B2 = (B2 - B2) * 2.5
        B3 = (B3 - B1) * 2.5
        B4 = (B4 - B4) * 2.5
        
        If R1 < 0 Then R1 = 0
        If G1 < 0 Then G1 = 0
        If B1 < 0 Then B1 = 0
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0
        If R3 < 0 Then R3 = 0
        If G3 < 0 Then G3 = 0
        If B3 < 0 Then B3 = 0
        If R4 < 0 Then R4 = 0
        If G4 < 0 Then G4 = 0
        If B4 < 0 Then B4 = 0
        If R1 > 255 Then R1 = 255
        If G1 > 255 Then G1 = 255
        If B1 > 255 Then B1 = 255
        If R2 > 255 Then R2 = 255
        If G2 > 255 Then G2 = 255
        If B2 > 255 Then B2 = 255
        If R3 > 255 Then R3 = 255
        If G3 > 255 Then G3 = 255
        If B3 > 255 Then B3 = 255
        If R4 > 255 Then R4 = 255
        If G4 > 255 Then G4 = 255
        If B4 > 255 Then B4 = 255
        
        m_p32Output(X, Y - 3).Red = Fix(R1)
        m_p32Output(X, Y - 3).Green = Fix(G1)
        m_p32Output(X, Y - 3).Blue = Fix(B1)
        m_p32Output(X, Y - 2).Red = Fix(R2)
        m_p32Output(X, Y - 2).Green = Fix(G2)
        m_p32Output(X, Y - 2).Blue = Fix(B2)
        m_p32Output(X, Y - 1).Red = Fix(R3)
        m_p32Output(X, Y - 1).Green = Fix(G3)
        m_p32Output(X, Y - 1).Blue = Fix(B3)
        m_p32Output(X, Y).Red = Fix(R4)
        m_p32Output(X, Y).Green = Fix(G4)
        m_p32Output(X, Y).Blue = Fix(B4)
    Next X
Next Y

Flatten = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Flatten2(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        R2 = m_p32Original(X + 1, Y + 1).Red
        G2 = m_p32Original(X + 1, Y + 1).Green
        B2 = m_p32Original(X + 1, Y + 1).Blue
        
        R = Abs((R1 - R2) / 2) + 100
        G = Abs((G1 - G2) / 2) + 100
        B = Abs((B1 - B2) / 2) + 100
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Flatten2 = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Silhuette(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        If Y = 0 Then
            R2 = R1
            G2 = G1
            B2 = B1
        End If
        
        R = ((Abs((R - 100) / 2)) + R1 - R2) / 2
        G = ((Abs((G - 100) / 2)) + G1 - G2) / 2
        B = ((Abs((B - 100) / 2)) + B1 - B2) / 2
        
        R2 = R1
        G2 = G1
        B2 = B1
        
        If R < 0 Then R = 0
            If G < 0 Then G = 0
            If B < 0 Then B = 0
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Silhuette = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Emboss(PicSrc As PictureBox, PicDest As PictureBox, Factor As Integer) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        R2 = m_p32Original(X + 1, Y + 1).Red
        G2 = m_p32Original(X + 1, Y + 1).Green
        B2 = m_p32Original(X + 1, Y + 1).Blue
        
        R = Abs(R1 - R2 - Factor)
        G = Abs(G1 - G2 - Factor)
        B = Abs(B1 - B2 - Factor)
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Emboss = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Pixilated(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim R3 As Integer, G3 As Integer, B3 As Integer
Dim R4 As Integer, G4 As Integer, B4 As Integer
Dim R5 As Integer, G5 As Integer, B5 As Integer
Dim R6 As Integer, G6 As Integer, B6 As Integer
Dim R7 As Integer, G7 As Integer, B7 As Integer
Dim R8 As Integer, G8 As Integer, B8 As Integer
Dim R9 As Integer, G9 As Integer, B9 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1 Step 2
    For X = 0 To m_lngWidth - 1 Step 2
        R1 = m_p32Original(X, Y).Red
        G1 = m_p32Original(X, Y).Green
        B1 = m_p32Original(X, Y).Blue
        
        R2 = m_p32Original(X - 1, Y - 1).Red
        G2 = m_p32Original(X - 1, Y - 1).Green
        B2 = m_p32Original(X - 1, Y - 1).Blue
        
        R3 = m_p32Original(X, Y - 1).Red
        G3 = m_p32Original(X, Y - 1).Green
        B3 = m_p32Original(X, Y - 1).Blue
        
        R4 = m_p32Original(X + 1, Y - 1).Red
        G4 = m_p32Original(X + 1, Y - 1).Green
        B4 = m_p32Original(X + 1, Y - 1).Blue
        
        R5 = m_p32Original(X - 1, Y).Red
        G5 = m_p32Original(X - 1, Y).Green
        B5 = m_p32Original(X - 1, Y).Blue
        
        R6 = m_p32Original(X + 1, Y).Red
        G6 = m_p32Original(X + 1, Y).Green
        B6 = m_p32Original(X + 1, Y).Blue
        
        R7 = m_p32Original(X - 1, Y - 1).Red
        G7 = m_p32Original(X - 1, Y - 1).Green
        B7 = m_p32Original(X - 1, Y - 1).Blue
        
        R8 = m_p32Original(X, Y - 1).Red
        G8 = m_p32Original(X, Y - 1).Green
        B8 = m_p32Original(X, Y - 1).Blue
        
        R9 = m_p32Original(X + 1, Y - 1).Red
        G9 = m_p32Original(X + 1, Y - 1).Green
        B9 = m_p32Original(X + 1, Y - 1).Blue
        
        R2 = (R1 + R2) / 2
        G2 = (G1 + G2) / 2
        B2 = (B1 + G2) / 2
        
        R3 = (R1 + G3) / 2
        G3 = (G1 + G3) / 2
        B3 = (B1 + B3) / 2
        
        R4 = (R1 + R4) / 2
        G4 = (G1 + G4) / 2
        B4 = (B1 + B4) / 2
        
        R5 = (R1 + R5) / 2
        G5 = (R1 + G5) / 2
        B5 = (B1 + B5) / 2
        
        R6 = (R1 + R6) / 2
        G6 = (G1 + G6) / 2
        B6 = (B1 + B6) / 2
        
        R7 = (R1 + R7) / 2
        G7 = (G1 + G7) / 2
        B7 = (B1 + B7) / 2
        
        R8 = (R1 + R8) / 2
        G8 = (G1 + G8) / 2
        B8 = (B1 + B8) / 2
        
        R9 = (R1 + R9) / 2
        G9 = (G1 + G9) / 2
        B9 = (B1 + B9) / 2
        
        If R1 < 0 Then R1 = 0
        If G1 < 0 Then G1 = 0
        If B1 < 0 Then B1 = 0
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0
        If R3 < 0 Then R3 = 0
        If G3 < 0 Then G3 = 0
        If B3 < 0 Then B3 = 0
        If R4 < 0 Then R4 = 0
        If G4 < 0 Then G4 = 0
        If B4 < 0 Then B4 = 0
        If R5 < 0 Then R5 = 0
        If G5 < 0 Then G5 = 0
        If B5 < 0 Then B5 = 0
        If R6 < 0 Then R6 = 0
        If G6 < 0 Then G6 = 0
        If B6 < 0 Then B6 = 0
        If R7 < 0 Then R7 = 0
        If G7 < 0 Then G7 = 0
        If B7 < 0 Then B7 = 0
        If R8 < 0 Then R8 = 0
        If G8 < 0 Then G8 = 0
        If B8 < 0 Then B8 = 0
        If R9 < 0 Then R9 = 0
        If G9 < 0 Then G9 = 0
        If B9 < 0 Then B9 = 0
        
        m_p32Output(X, Y).Red = R1
        m_p32Output(X, Y).Green = G1
        m_p32Output(X, Y).Blue = B1
        
        m_p32Output(X - 1, Y - 1).Red = R2
        m_p32Output(X - 1, Y - 1).Green = G2
        m_p32Output(X - 1, Y - 1).Blue = B2
        
        m_p32Output(X, Y - 1).Red = R3
        m_p32Output(X, Y - 1).Green = G3
        m_p32Output(X, Y - 1).Blue = B3
        
        m_p32Output(X + 1, Y - 1).Red = R4
        m_p32Output(X + 1, Y - 1).Green = G4
        m_p32Output(X + 1, Y - 1).Blue = B4
        
        m_p32Output(X - 1, Y).Red = R5
        m_p32Output(X - 1, Y).Green = G5
        m_p32Output(X - 1, Y).Blue = B5
        
        m_p32Output(X + 1, Y).Red = R6
        m_p32Output(X + 1, Y).Green = G6
        m_p32Output(X + 1, Y).Blue = B6
        
        m_p32Output(X - 1, Y - 1).Red = R7
        m_p32Output(X - 1, Y - 1).Green = G7
        m_p32Output(X - 1, Y - 1).Blue = B7
        
        m_p32Output(X, Y - 1).Red = R8
        m_p32Output(X, Y - 1).Green = G8
        m_p32Output(X, Y - 1).Blue = B8
        
        m_p32Output(X + 1, Y - 1).Red = R9
        m_p32Output(X + 1, Y - 1).Green = G9
        m_p32Output(X + 1, Y - 1).Blue = B9
    Next X
Next Y

Pixilated = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Pointalism(PicSrc As PictureBox, PicDest As PictureBox, Radius As Integer) As Long
Dim X As Long, Y As Long, Z As Long, Z2 As Long, u As Long
Dim R As Integer, G As Integer, B As Integer
Dim pix
X = Radius
Z = PicDest.Width / 15
Y = Radius
Z2 = PicDest.Height / 15

Setup2 PicSrc, PicDest

IsProcessing = True
PicDest.FillStyle = 0
Set PicDest.Picture = frmMain.imgNo.Image

Do Until X >= Z
    
    Y = Radius
    
    Do Until Y >= Z2
        u = u + 1: If u = 4000 Then DoEvents: u = 0
        
        pix = PicSrc.Point(X * 15, Y * 15)
        UnRGB pix, R%, G%, B%
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        PicDest.FillColor = RGB(R, G, B)
        PicDest.Circle (X * 15, Y * 15), Radius * 15, RGB(R, G, B)
        
    Y = Y + (Radius * 2)
    Loop

X = X + (Radius * 2)
Loop

Pointalism = ReturnTime()

IsProcessing = True
End Function

Public Function VividNegation(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
'Dim R1 As Integer, G1 As Integer, B1 As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        'Basic concept: Color = Abs(255 - Color) * 3
        'Times it by 1.2 to gain proper light/dark scale
        G2 = Abs(((G + G + G)) / 1.2)
        B2 = Abs(((B + B + B)) / 1.2)
        R2 = Abs(((R + R + R)) / 1.2)
        
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0
        
        If R2 > 255 Then R2 = 255
        If G2 > 255 Then G2 = 255
        If B2 > 255 Then B2 = 255
        
        G2 = 255 - G2
        B2 = 255 - B2
        R2 = 255 - R2
        
        m_p32Output(X, Y).Red = R2
        m_p32Output(X, Y).Green = G2
        m_p32Output(X, Y).Blue = B2
    Next X
Next Y

VividNegation = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function NeonNegation(PicSrc As PictureBox, PicDest As PictureBox, Modifyer As Long) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

IsProcessing = True

On Error Resume Next

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        '117 (old)
        'G2 = (Abs((255 - (G + (G * 2))) * 2) / 5) * 3
        'B2 = (Abs((255 - (B + (B * 2))) * 2) / 5) * 3
        'R2 = (Abs((255 - (R + (R * 2))) * 2) / 5) * 3
        
        '102 faster version that is mathematically summed up
        'G2 = Abs((255 - (G + G + G)) * 2) * 0.6
        'B2 = Abs((255 - (B + B + B)) * 2) * 0.6
        'R2 = Abs((255 - (R + R + R)) * 2) * 0.6
        
        '98ms, fastest version
        'Basic concept: Color = Abs(255 - Color) * 3
        'Times it by 1.2 to gain proper light/dark scale
        G2 = Abs((Modifyer - (G + G + G)) * 1.2)
        B2 = Abs((Modifyer - (B + B + B)) * 1.2)
        R2 = Abs((Modifyer - (R + R + R)) * 1.2)
        
        If R2 < 0 Then R2 = 0
        If G2 < 0 Then G2 = 0
        If B2 < 0 Then B2 = 0

        If R2 > 255 Then R2 = 255
        If G2 > 255 Then G2 = 255
        If B2 > 255 Then B2 = 255
        
        m_p32Output(X, Y).Red = 255 - R2
        m_p32Output(X, Y).Green = 255 - G2
        m_p32Output(X, Y).Blue = 255 - B2
    Next X
Next Y

NeonNegation = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function
