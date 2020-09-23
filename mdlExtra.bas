Attribute VB_Name = "mdlExtra"
'Title:           mdlExtra
'Version:         2.0
'Date:            10/07/2008
'Author:          Skyler Lyon
'Copyright:       © 2008 Skyler Lyon
'Description:     Module for image effect and time handling.

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public m_p32Original() As Pixel32
Public m_p32Output() As Pixel32

Public m_lngWidth As Long, m_lngHeight As Long

Public m_lngTimeElapsed As Long
Public m_lngStartTime As Long, m_lngEndTime As Long
Public m_dblPerformanceFrequency As Double

Public IsProcessing As Boolean

Public Function GetTimeMS() As Long
Dim m_curTime As Currency
    Call QueryPerformanceCounter(m_curTime)
    GetTimeMS = CLng((CDbl(m_curTime) / m_dblPerformanceFrequency) * 1000)
End Function

Public Function RandomNumber(finished)
    Randomize
    RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Public Sub UnRGB(ByVal Color As OLE_COLOR, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    B = Color \ 65536
    G = (Color \ 256) Mod 256
    R = Color Mod 256
End Sub
