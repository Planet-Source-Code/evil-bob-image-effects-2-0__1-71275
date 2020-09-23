Attribute VB_Name = "mdlOldImageEffects"
'Title:           mdlOldImageEffects
'Version:         1.0
'Date:            10/07/2008
'Author:          Skyler Lyon
'Copyright:       © 2008 Skyler Lyon
'Description:     Old module for image effects.

'Public Sub GrayScale(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, b%
'            GC = (R + G + b) / 3
'            PicDest.PSet (X * 15, Y * 15), RGB(GC, GC, GC)
'    Next Y, X
'End Sub

'Public Sub Negative(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, b%
'            R = 255 - R
'            G = 255 - G
'            b = 255 - b
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, b)
'        Next Y, X
'End Sub

'Public Sub Lighten(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'            R = R + Magnitude
'            G = G + Magnitude
'            B = B + Magnitude
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Sub

'Public Sub Darken(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'            R = R - Magnitude
'            G = G - Magnitude
'            B = B - Magnitude
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Sub

'Public Function Blur(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R1%, G1%, B1%
'
'            If Y = 0 Then
'            R2 = R1
'            G2 = G1
'            B2 = B1
'            End If
'
'            R = (R1 + R2) / 2
'            G = (G1 + G2) / 2
'            B = (B1 + B2) / 2
'
'            R2 = R1
'            G2 = G1
'            B2 = B1
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Function

'Public Sub Blur2(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R1%, G1%, B1%
'
'            If Y = 0 Then
'            R3 = R1
'            G3 = R1
'            B3 = R1
'
'            R2 = R1
'            G2 = G1
'            B2 = B1
'            End If
'
'            R = (((R1 + R2) / 2) + R3) / 2
'            G = (((G1 + G2) / 2) + G3) / 2
'            B = (((B1 + B2) / 2) + B3) / 2
'
'            R3 = R2
'            G3 = G2
'            B3 = B2
'
'            R2 = R1
'            G2 = G1
'            B2 = B1
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Sub
'
'Public Sub NoRed(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = 0
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub NoGreen(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            G = 0
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub NoBlue(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            B = 0
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub

'Public Sub Warped(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            p = RandomNumber(9)
'
'            R = (R * p) / 2
'            G = (G * p) / 2
'            B = (B * p) / 2
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub IncreaseRed(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = (R + Magnitude)
'            G = G
'            B = B
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub IncreaseGreen(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R
'            G = (G + Magnitude)
'            B = B
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub IncreaseBlue(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R
'            G = G
'            B = (B + Magnitude)
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub IncreaseRGB(PicSrc As PictureBox, PicDest As PictureBox, Red, Green, Blue)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R + Red
'            G = G + Green
'            B = B + Blue
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub

'Public Sub SandBlasting(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            p = RandomNumber(5)
'
'            If p = 0 Then
'            R = 255 + R
'            G = 128 + G
'            B = 64 + B
'            End If
'
'            If p = 1 Then
'            R = 135 + R
'            G = 45 + G
'            B = B
'            End If
'
'            If p = 2 Then
'            R = 220 + R
'            G = 220 + G
'            B = 220 + B
'            End If
'
'            If p = 3 Then
'            R = 128 = R
'            G = 64 + G
'            B = B
'            End If
'
'            If p = 4 Then
'            R = 200 + R
'            G = 180 + G
'            B = 190 + B
'            End If
'
'            If p = 5 Then
'            R = 90 + R
'            G = 45 + G
'            B = B
'            End If
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub

'Public Sub Incoherence(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = B
'            G = G
'            B = R
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Sub

'Public Sub DecreaseRed(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = (R - Magnitude)
'            G = G
'            B = B
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub DecreaseGreen(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R
'            G = (G - Magnitude)
'            B = B
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub DecreaseBlue(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R
'            G = G
'            B = (B - Magnitude)
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub DecreaseRGB(PicSrc As PictureBox, PicDest As PictureBox, Red, Green, Blue)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            R = R - Red
'            G = G - Green
'            B = B - Blue
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub

'Public Sub GlowInTheDark(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15  'Mag: 1 - 10
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            On Error Resume Next
'
'            G2 = Abs(((255 * 3) / G)) * Magnitude
'            B2 = Abs(((255 * 3) / B)) * Magnitude
'            R2 = Abs(((255 * 3) / R)) * Magnitude
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R2, G2, B2)
'        Next Y, X
'End Sub

'Public Sub SunShinnyDay(PicSrc As PictureBox, PicDest As PictureBox, Magnitude)
'    For X = 0 To PicDest.Width / 15 'Mag 1 - 40
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            G2 = Abs((128 * G) / Magnitude)
'            B2 = Abs((128 * B) / Magnitude)
'            R2 = Abs((128 * R) / Magnitude)
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R2, G2, B2)
'        Next Y, X
'End Sub

'Public Sub AmphibiousSilk(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            On Error Resume Next
'
'            G2 = (Abs((255 - (G + (G * 2))) * 2) / 5) * 3
'            B2 = (Abs((255 - (B + (B * 2))) * 2) / 5) * 3
'            R2 = (Abs((255 - (R + (R * 2))) * 2) / 5) * 3
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R2, G2, B2)
'        Next Y, X
'End Sub

'Public Function Flatten(PicSrc As PictureBox, PicTarg As PictureBox)
'    For X = 0 To PicSrc.Width / 15
'
'        For Y = 0 To PicTarg.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'
'            p1 = PicSrc.Point(X * 15, (Y - 3) * 15)
'            If p1 < 0 Then p1 = 0
'            UnRGB p1, R1%, G1%, B1%
'
'            p2 = PicSrc.Point(X * 15, (Y - 2) * 15)
'            If p2 < 0 Then p2 = 0
'            UnRGB p2, R2%, G2%, B2%
'
'            p3 = PicSrc.Point(X * 15, (Y - 1) * 15)
'            If p3 < 0 Then p3 = 0
'            UnRGB p3, R3%, G3%, B3%
'
'            p4 = PicSrc.Point(X * 15, Y * 15)
'            If p4 < 0 Then p4 = 0
'            UnRGB p4, R4%, G4%, B4%
'
'            R1 = (R1 - R3) * 2.5
'            R2 = (R2 - R2) * 2.5
'            R3 = (R3 - R1) * 2.5
'            R4 = (R4 - R4) * 2.5
'            G1 = (G1 - G3) * 2.5
'            G2 = (G2 - G2) * 2.5
'            G3 = (G3 - G1) * 2.5
'            G4 = (G4 - G4) * 2.5
'            B1 = (B1 - B3) * 2.5
'            B2 = (B2 - B2) * 2.5
'            B3 = (B3 - B1) * 2.5
'            B4 = (B4 - B4) * 2.5
'
'            If R1 < 0 Then R1 = 0
'            If G1 < 0 Then G1 = 0
'            If B1 < 0 Then B1 = 0
'            If R2 < 0 Then R2 = 0
'            If G2 < 0 Then G2 = 0
'            If B2 < 0 Then B2 = 0
'            If R3 < 0 Then R3 = 0
'            If G3 < 0 Then G3 = 0
'            If B3 < 0 Then B3 = 0
'            If R4 < 0 Then R4 = 0
'            If G4 < 0 Then G4 = 0
'            If B4 < 0 Then B4 = 0
'
'            PicTarg.PSet (X * 15, (Y - 3) * 15), RGB(Fix(R1%), Fix(G1%), Fix(B1%))
'            PicTarg.PSet (X * 15, (Y - 2) * 15), RGB(Fix(R2%), Fix(G2%), Fix(B2%))
'            PicTarg.PSet (X * 15, (Y - 1) * 15), RGB(Fix(R3%), Fix(G3%), Fix(B3%))
'            PicTarg.PSet (X * 15, Y * 15), RGB(Fix(R4%), Fix(G4%), Fix(B4%))
'        Next Y, X
'End Function

'Public Function Silhuette(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R1%, G1%, B1%
'
'            If Y = 0 Then
'            R2 = R1
'            G2 = G1
'            B2 = B1
'            End If
'
'            R = ((Abs((R - 100) / 2)) + R1 - R2) / 2
'            G = ((Abs((G - 100) / 2)) + G1 - G2) / 2
'            B = ((Abs((B - 100) / 2)) + B1 - B2) / 2
'
'            R2 = R1
'            G2 = G1
'            B2 = B1
'
'            If R < 0 Then R = 0
'            If G < 0 Then G = 0
'            If B < 0 Then B = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Function

'Public Sub Silk2(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R1%, G1%, B1%
'
'            On Error Resume Next
'
'            If Y = 0 Then
'            R2 = R1
'            G2 = G1
'            B2 = B1
'            End If
'
'            R = Abs(255 - ((R1 * R2) / 10))
'            G = Abs(255 - ((G1 * G2) / 10))
'            B = Abs(255 - ((B1 * B2) / 10))
'
'            R2 = R1
'            G2 = G1
'            B2 = B1
'
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'        Next Y, X
'End Sub

'Public Sub Emboss(PicSrc As PictureBox, PicDest As PictureBox, Factor)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix1 = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix1, R1%, G1%, B1%
'
'            pix1 = PicSrc.Point((X + 1) * 15, (Y + 1) * 15)
'            UnRGB pix1, R2%, G2%, B2%
'
'            R = Abs(R1 - R2 - Factor)
'            G = Abs(G1 - G2 - Factor)
'            B = Abs(B1 - B2 - Factor)
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub
'
'Public Sub Pixilated(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15 Step 2
'        For Y = 0 To PicDest.Height / 15 Step 2
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R1%, G1%, B1%
'
'            pix = PicSrc.Point((X - 1) * 15, (Y - 1) * 15)
'            UnRGB pix, R2%, G2%, B2%
'
'            pix = PicSrc.Point(X * 15, (Y - 1) * 15)
'            UnRGB pix, R3%, G3%, B3%
'
'            pix = PicSrc.Point((X + 1) * 15, (Y - 1) * 15)
'            UnRGB pix, R4%, G4%, B4%
'
'            pix = PicSrc.Point((X - 1) * 15, Y * 15)
'            UnRGB pix, R5%, G5%, B5%
'
'            pix = PicSrc.Point((X + 1) * 15, Y * 15)
'            UnRGB pix, R6%, G6%, B6%
'
'            pix = PicSrc.Point((X - 1) * 15, (Y - 1) * 15)
'            UnRGB pix, R7%, G7%, B7%
'
'            pix = PicSrc.Point(X * 15, (Y - 1) * 15)
'            UnRGB pix, R8%, G8%, B8%
'
'            pix = PicSrc.Point((X + 1) * 15, (Y - 1) * 15)
'            UnRGB pix, R9%, G9%, B9%
'
'            R2 = (R1 + R2) / 2
'            G2 = (G1 + G2) / 2
'            B2 = (B1 + G2) / 2
'
'            R3 = (R1 + G3) / 2
'            G3 = (G1 + G3) / 2
'            B3 = (B1 + B3) / 2
'
'            R4 = (R1 + R4) / 2
'            G4 = (G1 + G4) / 2
'            B4 = (B1 + B4) / 2
'
'            R5 = (R1 + R5) / 2
'            G5 = (R1 + G5) / 2
'            B5 = (B1 + B5) / 2
'
'            R6 = (R1 + R6) / 2
'            G6 = (G1 + G6) / 2
'            B6 = (B1 + B6) / 2
'
'            R7 = (R1 + R7) / 2
'            G7 = (G1 + G7) / 2
'            B7 = (B1 + B7) / 2
'
'            R8 = (R1 + R8) / 2
'            G8 = (G1 + G8) / 2
'            B8 = (B1 + B8) / 2
'
'            R9 = (R1 + R9) / 2
'            G9 = (G1 + G9) / 2
'            B9 = (B1 + B9) / 2
'
'            If R1 < 0 Then R1 = 0
'            If G1 < 0 Then G1 = 0
'            If B1 < 0 Then B1 = 0
'            If R2 < 0 Then R2 = 0
'            If G2 < 0 Then G2 = 0
'            If B2 < 0 Then B2 = 0
'            If R3 < 0 Then R3 = 0
'            If G3 < 0 Then G3 = 0
'            If B3 < 0 Then B3 = 0
'            If R4 < 0 Then R4 = 0
'            If G4 < 0 Then G4 = 0
'            If B4 < 0 Then B4 = 0
'            If R5 < 0 Then R5 = 0
'            If G5 < 0 Then G5 = 0
'            If B5 < 0 Then B5 = 0
'            If R6 < 0 Then R6 = 0
'            If G6 < 0 Then G6 = 0
'            If B6 < 0 Then B6 = 0
'            If R7 < 0 Then R7 = 0
'            If G7 < 0 Then G7 = 0
'            If B7 < 0 Then B7 = 0
'            If R8 < 0 Then R8 = 0
'            If G8 < 0 Then G8 = 0
'            If B8 < 0 Then B8 = 0
'            If R9 < 0 Then R9 = 0
'            If G9 < 0 Then G9 = 0
'            If B9 < 0 Then B9 = 0
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R1, G1, B1)
'            PicDest.PSet ((X - 1) * 15, (Y - 1) * 15), RGB(R2, G2, B2)
'            PicDest.PSet (X * 15, (Y - 1) * 15), RGB(R3, G3, B3)
'            PicDest.PSet ((X + 1) * 15, (Y - 1) * 15), RGB(R4, G4, B4)
'            PicDest.PSet ((X - 1) * 15, Y * 15), RGB(R5, G5, B5)
'            PicDest.PSet ((X + 1) * 15, Y * 15), RGB(R6, G6, B6)
'            PicDest.PSet ((X - 1) * 15, (Y - 1) * 15), RGB(R7, G7, B7)
'            PicDest.PSet (X * 15, (Y - 1) * 15), RGB(R8, G8, B8)
'            PicDest.PSet ((X + 1) * 15, (Y - 1) * 15), RGB(R9, G9, B9)
'        Next Y, X
'End Sub

'Public Sub Flatten2(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix1 = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix1, R1%, G1%, B1%
'
'            pix1 = PicSrc.Point((X + 1) * 15, (Y + 1) * 15)
'            UnRGB pix1, R2%, G2%, B2%
'
'            On Error Resume Next
'            R = Abs((R1 - R2) / 2) + 100
'            G = Abs((G1 - G2) / 2) + 100
'            B = Abs((B1 - B2) / 2) + 100
'            PicDest.PSet (X * 15, Y * 15), RGB(R, G, B)
'    Next Y, X
'End Sub

'Public Sub RGBFilm(PicSrc As PictureBox, PicDest As PictureBox)
'    For X = 0 To PicDest.Width / 15
'        For Y = 0 To PicDest.Height / 15
'            u = u + 1: If u = 4000 Then DoEvents: u = 0
'            pix = PicSrc.Point(X * 15, Y * 15)
'            UnRGB pix, R%, G%, B%
'
'            On Error Resume Next
'
'            G2 = (Abs((255 - (G + (G * 2))) * 2) / 5) * 3
'            B2 = (Abs((255 - (B + (B * 2))) * 2) / 5) * 3
'            R2 = (Abs((255 - (R + (R * 2))) * 2) / 5) * 3
'
'            If R2 < 0 Then R2 = 0
'            If G2 < 0 Then G2 = 0
'            If B2 < 0 Then B2 = 0
'
'            If R2 > 255 Then R2 = 255
'            If G2 > 255 Then G2 = 255
'            If B2 > 255 Then B2 = 255
'
'            G2 = 255 - G2
'            B2 = 255 - B2
'            R2 = 255 - R2
'
'            PicDest.PSet (X * 15, Y * 15), RGB(R2, G2, B2)
'        Next Y, X
'End Sub

'Public Sub Pointalism(PicSrc As PictureBox, PicDest As PictureBox, Radius As Integer)
'Dim X As Long, Y As Long, Z As Long, Z2 As Long
'X = Radius
'Z = PicDest.Width / 15
'Y = Radius
'Z2 = PicDest.Height / 15
'
'PicDest.FillStyle = 0
'Set PicDest.Picture = imgNo.Image
'
'Do Until X >= Z
'
'    Y = Radius
'
'    Do Until Y >= Z2
'        u = u + 1: If u = 4000 Then DoEvents: u = 0
'
'        pix = PicSrc.Point(X * 15, Y * 15)
'        UnRGB pix, R%, G%, B%
'
'        If R < 0 Then R = 0
'        If G < 0 Then G = 0
'        If B < 0 Then B = 0
'
'        PicDest.FillColor = RGB(R, G, B)
'        PicDest.Circle (X * 15, Y * 15), Radius * 15, RGB(R, G, B)
'
'    Y = Y + (Radius * 2)
'    Loop
'
'X = X + (Radius * 2)
'Loop
'End Sub


'Public Sub UnRGB(ByVal Color As OLE_COLOR, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
'    B = Color \ 65536
'    G = (Color \ 256) Mod 256
'    R = Color Mod 256
'End Sub

'Public Function RandomNumber(finished)
'    Randomize
'    RandomNumber = Int((Val(finished) * Rnd) + 1)
'End Function
