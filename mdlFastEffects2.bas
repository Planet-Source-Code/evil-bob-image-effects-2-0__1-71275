Attribute VB_Name = "mdlFastEffects2"
'Title:           mdlFastEffects2
'Version:         2.0
'Date:            10/07/2008
'Author:          Skyler Lyon
'Copyright:       © 2008 Skyler Lyon
'Description:     Module for extra basic image effects.

Option Explicit

Public Function NoRed(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim intGrayScale As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        m_p32Output(X, Y).Red = 0
        m_p32Output(X, Y).Green = m_p32Original(X, Y).Green
        m_p32Output(X, Y).Blue = m_p32Original(X, Y).Blue
    Next X
Next Y

NoRed = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function NoGreen(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim intGrayScale As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        m_p32Output(X, Y).Green = 0
        m_p32Output(X, Y).Red = m_p32Original(X, Y).Red
        m_p32Output(X, Y).Blue = m_p32Original(X, Y).Blue
    Next X
Next Y

NoGreen = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function NoBlue(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim intGrayScale As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        m_p32Output(X, Y).Blue = 0
        m_p32Output(X, Y).Green = m_p32Original(X, Y).Green
        m_p32Output(X, Y).Red = m_p32Original(X, Y).Red
    Next X
Next Y

NoBlue = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function Warped(PicSrc As PictureBox, PicDest As PictureBox) As Long
Dim X As Long, Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim p As Integer

IsProcessing = True

Setup PicSrc, PicDest

'Loop
For Y = 0 To m_lngHeight - 1
    For X = 0 To m_lngWidth - 1
        R = m_p32Original(X, Y).Red
        G = m_p32Original(X, Y).Green
        B = m_p32Original(X, Y).Blue
        
        p = RandomNumber(9)
            
        R = (R * p) / 2
        G = (G * p) / 2
        B = (B * p) / 2
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

Warped = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function IncreaseRed(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Integer) As Long
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
        
        R = (R + Magnitude)
        
        If R < 0 Then R = 0
        If R > 255 Then R = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

IncreaseRed = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function IncreaseGreen(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Integer) As Long
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
        
        G = (G + Magnitude)
        
        If G < 0 Then G = 0
        If G > 255 Then G = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

IncreaseGreen = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function IncreaseBlue(PicSrc As PictureBox, PicDest As PictureBox, Magnitude As Integer) As Long
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
        
        B = (B + Magnitude)
        
        If B < 0 Then B = 0
        If B > 255 Then B = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

IncreaseBlue = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function

Public Function IncreaseRGB(PicSrc As PictureBox, PicDest As PictureBox, Red As Integer, Green As Integer, Blue As Integer) As Long
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
        
        R = R + Red
        G = G + Green
        B = B + Blue
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        
        m_p32Output(X, Y).Red = R
        m_p32Output(X, Y).Green = G
        m_p32Output(X, Y).Blue = B
    Next X
Next Y

IncreaseRGB = ReturnPicAndTime(PicDest)

IsProcessing = False
End Function
