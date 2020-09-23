Attribute VB_Name = "modFilterScale"
Option Explicit
' vb implementation of a 2-pass resample filter using a generic interface template
' for the filter and inheritance to derive a bunch of coClasses from the master
' class with the Implements keyword - this allows us to pass different filter
' classes into the scale function using a data type of the master interface class ...

Public Const FILTER_PI As Double = (3.14159265358979)
Public Const FILTER_2PI As Double = (2# * 3.14159265358979)
Public Const FILTER_4PI As Double = (4# * 3.14159265358979)

Public Type ContributionType
   Weights() As Double
   Left As Integer
   Right As Integer
End Type

Public Type LineContribType
   ContribRow() As ContributionType
   WindowSize As Integer
   LineLength As Integer
End Type

Dim CurFilter As IGenericFilter

Private Function AllocContributions(ByVal uLineLength As Integer, ByVal uWindowSize As Integer) As LineContribType
    Dim res As LineContribType
    Dim i As Integer
    
    res.WindowSize = uWindowSize
    res.LineLength = uLineLength
    ReDim res.ContribRow(uLineLength)
    For i = 0 To uLineLength - 1
        ReDim res.ContribRow(i).Weights(uWindowSize)
    Next i
    AllocContributions = res
End Function

Private Sub FreeContributions(p As LineContribType)
    Dim i As Integer
    
    For i = 0 To p.LineLength - 1
        Erase p.ContribRow(i).Weights
    Next i
    Erase p.ContribRow
End Sub
 
 Private Function Max(ByVal i As Double, ByVal j As Double) As Double
    If i > j Then Max = i Else Max = j
 End Function
 
 Private Function Min(ByVal i As Double, ByVal j As Double) As Double
    If i < j Then Min = i Else Min = j
 End Function
 
Private Function CalcContributions(ByVal uLineSize As Integer, ByVal uSrcSize As Integer, ByVal dScale As Double) As LineContribType
    Dim res As LineContribType
    Dim dWidth As Double, dFScale As Double, dFilterWidth As Double
    Dim iWindowSize As Integer, u As Integer, dCenter As Double, iLeft As Integer
    Dim iRight As Integer, dTotalWeight As Double, iSrc As Integer

    dFScale = 1#
    dFilterWidth = CurFilter.Width

    If (dScale < 1#) Then
        ' Minification
        dWidth = dFilterWidth / dScale
        dFScale = dScale
    Else
        ' Magnification
        dWidth = dFilterWidth
    End If

    ' Window size is the number of sampled pixels
    iWindowSize = 2 * Int(dWidth) + 1

    ' Allocate a new line contributions strucutre
    res = AllocContributions(uLineSize, iWindowSize)

    For u = 0 To uLineSize - 1
        ' Scan through line of contributions
        dCenter = CDbl(u / dScale) ' Reverse mapping
        ' Find the significant edge points that affect the pixel
        iLeft = Max(0, CInt(Fix(dCenter - dWidth)))
        iRight = Min(Int(dCenter + dWidth), Int(uSrcSize) - 1)

        ' Cut edge points to fit in filter window in case of spill-off
        If (iRight - iLeft + 1 > iWindowSize) Then
            iLeft = iLeft + 1
        Else
            iRight = iRight + 1
        End If
        ' ets+++ adjusted ileft and iright values not stored in contrib array
        res.ContribRow(u).Left = iLeft
        res.ContribRow(u).Right = iRight

        ' ets
        dTotalWeight = 0# ' Zero sum of weights
        For iSrc = iLeft To iRight
            ' Calculate weights
            res.ContribRow(u).Weights(iSrc - iLeft) = _
                (dFScale * CurFilter.Filter(dFScale * (dCenter - CDbl(iSrc))))
            dTotalWeight = dTotalWeight + (res.ContribRow(u).Weights(iSrc - iLeft))
        Next iSrc
        If (dTotalWeight > 0#) Then
           ' Normalize weight of neighbouring points
            For iSrc = iLeft To iRight
                ' Normalize point
                res.ContribRow(u).Weights(iSrc - iLeft) = _
                    res.ContribRow(u).Weights(iSrc - iLeft) / dTotalWeight
            Next iSrc
        End If
    Next u
    CalcContributions = res
End Function

Private Sub ScaleRow(pSrc() As RGBQUAD, ByVal uSrcWidth As Integer, pRes() As RGBQUAD, ByVal uResWidth As Integer, ByVal uRow As Integer, Contrib As LineContribType)
    Dim x As Integer, iLeft As Integer, iRight As Integer
    Dim i As Integer, r As Long, g As Long, b As Long

    For x = 0 To uResWidth - 1
        r = 0: g = 0: b = 0
        ' Loop through row
        iLeft = Contrib.ContribRow(x).Left
        iRight = Contrib.ContribRow(x).Right
        For i = iLeft To iRight ' Scan between boundries
            ' Accumulate weighted effect of each neighboring pixel
On Error Resume Next
            r = r + CByte(Contrib.ContribRow(x).Weights(i - iLeft) * CDbl(pSrc(i, uRow).rgbRed))
            g = g + CByte(Contrib.ContribRow(x).Weights(i - iLeft) * CDbl(pSrc(i, uRow).rgbGreen))
            b = b + CByte(Contrib.ContribRow(x).Weights(i - iLeft) * CDbl(pSrc(i, uRow).rgbBlue))
On Error GoTo 0
        Next i
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        If g < 0 Then g = 0
        If g > 255 Then g = 255
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        pRes(x, uRow).rgbRed = r
        pRes(x, uRow).rgbGreen = g
        pRes(x, uRow).rgbBlue = b
    Next x
End Sub

Private Sub HorizScale(pSrc() As RGBQUAD, ByVal uSrcWidth As Integer, ByVal uSrcHeight As Integer, pDst() As RGBQUAD, ByVal uResWidth As Integer, ByVal uResHeight As Integer)
    Dim Contrib As LineContribType
    Dim i As Integer
    
    If (uResWidth = uSrcWidth) Then
        ReDim pDst(0 To uSrcWidth - 1, 0 To uSrcHeight - 1) As RGBQUAD
        pDst = CloneBits(pSrc)
        Exit Sub
    End If
    
    Contrib = CalcContributions(uResWidth, uSrcWidth, CDbl(uResWidth) / CDbl(uSrcWidth))
    For i = 0 To uResHeight - 1
        ScaleRow pSrc, uSrcWidth, pDst, uResWidth, i, Contrib
    Next i
    FreeContributions Contrib
End Sub
    
Private Sub ScaleCol(pSrc() As RGBQUAD, ByVal uSrcWidth As Integer, pRes() As RGBQUAD, ByVal uResWidth As Integer, ByVal uResHeight As Integer, ByVal uCol As Integer, Contrib As LineContribType)
    Dim y As Integer, iLeft As Integer, iRight As Integer, pCurSrc As RGBQUAD
    Dim i As Integer, r As Long, g As Long, b As Long

    For y = 0 To uResHeight - 1
        r = 0: g = 0: b = 0
        ' Loop through row
        iLeft = Contrib.ContribRow(y).Left
        iRight = Contrib.ContribRow(y).Right
        For i = iLeft To iRight ' Scan between boundries
            ' Accumulate weighted effect of each neighboring pixel
On Error Resume Next
            pCurSrc = pSrc(uCol, i)
            r = r + CByte(Contrib.ContribRow(y).Weights(i - iLeft) * CDbl(pCurSrc.rgbRed))
            g = g + CByte(Contrib.ContribRow(y).Weights(i - iLeft) * CDbl(pCurSrc.rgbGreen))
            b = b + CByte(Contrib.ContribRow(y).Weights(i - iLeft) * CDbl(pCurSrc.rgbBlue))
 On Error GoTo 0
       Next i
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        If g < 0 Then g = 0
        If g > 255 Then g = 255
        If b < 0 Then b = 0
        If b > 255 Then b = 255
        pRes(uCol, y).rgbRed = r
        pRes(uCol, y).rgbGreen = g
        pRes(uCol, y).rgbBlue = b
    Next y
End Sub

Private Sub VertScale(pSrc() As RGBQUAD, ByVal uSrcWidth As Integer, ByVal uSrcHeight As Integer, pDst() As RGBQUAD, ByVal uResWidth As Integer, ByVal uResHeight As Integer)
    Dim Contrib As LineContribType
    Dim i As Integer

    If (uResHeight = uSrcHeight) Then
        ReDim pDst(0 To uSrcWidth - 1, 0 To uSrcHeight - 1) As RGBQUAD
        pDst = CloneBits(pSrc)
        Exit Sub
    End If
    
    Contrib = CalcContributions(uResHeight, uSrcHeight, CDbl(uResHeight) / CDbl(uSrcHeight))

    For i = 0 To uResWidth - 1
        ScaleCol pSrc, uSrcWidth, pDst, uResWidth, uResHeight, i, Contrib
    Next i
    FreeContributions Contrib
End Sub
                             
Public Function AllocAndScale(pOrigImage() As RGBQUAD, ByVal uOrigWidth As Integer, ByVal uOrigHeight As Integer, ByVal uNewWidth As Integer, ByVal uNewHeight As Integer, ByVal FilterIndex As Integer) As RGBQUAD()
    Dim pTemp() As RGBQUAD
    Dim pRes() As RGBQUAD
    
    ' set up the current filter
    Select Case FilterIndex
        Case 1: Set CurFilter = New ClsBoxFilter
        Case 2: Set CurFilter = New ClsBilinearFilter
        Case 3: Set CurFilter = New ClsBlackmanFilter
        Case 4: Set CurFilter = New ClsGaussianFilter
        Case 5: Set CurFilter = New ClsHammingFilter
        Case 6: Set CurFilter = New ClsQuadraticFilter
        Case 7: Set CurFilter = New ClsHanningFilter
        Case 8: Set CurFilter = New ClsMitchellFilter
        Case 9: Set CurFilter = New ClsLanczosFilter
        Case 10: Set CurFilter = New ClsHermiteFilter
        Case 11: Set CurFilter = New ClsCubicFilter
        Case 12: Set CurFilter = New ClsCatromFilter
        Case 13: Set CurFilter = New ClsSplineFilter
        Case 14: Set CurFilter = New ClsBellFilter
        Case 15: Set CurFilter = New ClsTriangleFilter
    End Select
    
    ReDim pTemp(0 To uNewWidth - 1, 0 To uOrigHeight - 1) As RGBQUAD
    HorizScale pOrigImage, uOrigWidth, uOrigHeight, pTemp, uNewWidth, uOrigHeight
    ReDim pRes(0 To uNewWidth - 1, 0 To uNewHeight - 1) As RGBQUAD
    VertScale pTemp, uNewWidth, uOrigHeight, pRes, uNewWidth, uNewHeight
    Erase pTemp
    AllocAndScale = pRes
End Function

Function StdResize(img() As RGBQUAD, ByVal w As Integer, ByVal h As Integer, ByVal w2 As Integer, ByVal h2 As Integer) As RGBQUAD()
    Dim ret() As RGBQUAD, x As Integer, y As Integer
    Dim CalcY As Integer, DstY As Single, yScale As Single, IntrplY As Single
    Dim CalcX As Integer, DstX As Single, xScale As Single, IntrplX As Single
    'The red and green values the we use to interpolate the new pixel
    Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
    Dim g As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
    Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
    'The interpolated red, green, and blue
    Dim Ir1 As Long, Ig1 As Long, Ib1 As Long
    Dim Ir2 As Long, Ig2 As Long, Ib2 As Long
    
    ReDim ret(0 To w2 - 1, 0 To h2 - 1) As RGBQUAD
    
    xScale = (w - 1) / w2
    yScale = (h - 1) / h2
'Draw each pixel in the new image
    For y = 0 To h2 - 1
'Generate the y calculation variables
        DstY = y * yScale
        IntrplY = Int(DstY)
        CalcY = DstY - IntrplY
        For x = 0 To w2 - 1
'Generate the x calculation variables
            DstX = x * xScale
            IntrplX = Int(DstX)
            CalcX = DstX - IntrplX
            
            'Get the 4 pixels around the interpolated one
            r1 = img(IntrplX, IntrplY).rgbRed
            g1 = img(IntrplX, IntrplY).rgbGreen
            b1 = img(IntrplX, IntrplY).rgbBlue
            
            r2 = img(IntrplX + 1, IntrplY).rgbRed
            g2 = img(IntrplX + 1, IntrplY).rgbGreen
            b2 = img(IntrplX + 1, IntrplY).rgbBlue
            
            r3 = img(IntrplX, IntrplY + 1).rgbRed
            g3 = img(IntrplX, IntrplY + 1).rgbGreen
            b3 = img(IntrplX, IntrplY + 1).rgbBlue

            r4 = img(IntrplX + 1, IntrplY + 1).rgbRed
            g4 = img(IntrplX + 1, IntrplY + 1).rgbGreen
            b4 = img(IntrplX + 1, IntrplY + 1).rgbBlue

            'Interpolate the R,G,B values in the X direction
            Ir1 = r1 * (1 - CalcY) + r3 * CalcY
            Ig1 = g1 * (1 - CalcY) + g3 * CalcY
            Ib1 = b1 * (1 - CalcY) + b3 * CalcY
            Ir2 = r2 * (1 - CalcY) + r4 * CalcY
            Ig2 = g2 * (1 - CalcY) + g4 * CalcY
            Ib2 = b2 * (1 - CalcY) + b4 * CalcY
            'Intepolate the R,G,B values in the Y direction
            r = Ir1 * (1 - CalcX) + Ir2 * CalcX
            g = Ig1 * (1 - CalcX) + Ig2 * CalcX
            b = Ib1 * (1 - CalcX) + Ib2 * CalcX
            
            'Make sure that the values are in the acceptable range
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            'Set this pixel onto the new picture box
            ret(x, y).rgbRed = r
            ret(x, y).rgbGreen = g
            ret(x, y).rgbBlue = b
        Next x
    Next y
    StdResize = ret
End Function

' this on uses the average of all neighbours ...
Function StdResize2(img() As RGBQUAD, ByVal w As Integer, ByVal h As Integer, ByVal w2 As Integer, ByVal h2 As Integer) As RGBQUAD()
    Dim ret() As RGBQUAD, x As Integer, y As Integer
    Dim CalcY As Integer, DstY As Single, yScale As Single, IntrplY As Single
    Dim CalcX As Integer, DstX As Single, xScale As Single, IntrplX As Single
    'The red and green values the we use to interpolate the new pixel
    Dim r As Long, r1 As Long, r2 As Long, r3 As Long, r4 As Long
    Dim r5 As Long, r6 As Long, r7 As Long, r8 As Long, r9 As Long
    Dim g As Long, g1 As Long, g2 As Long, g3 As Long, g4 As Long
    Dim g5 As Long, g6 As Long, g7 As Long, g8 As Long, g9 As Long
    Dim b As Long, b1 As Long, b2 As Long, b3 As Long, b4 As Long
    Dim b5 As Long, b6 As Long, b7 As Long, b8 As Long, b9 As Long
    'The interpolated red, green, and blue
    Dim Ir1 As Long, Ig1 As Long, Ib1 As Long
    Dim Ir2 As Long, Ig2 As Long, Ib2 As Long
    
    ReDim ret(0 To w2 - 1, 0 To h2 - 1) As RGBQUAD
    
    xScale = (w - 1) / w2
    yScale = (h - 1) / h2
'Draw each pixel in the new image
On Error Resume Next
    For y = 1 To h2 - 2
'Generate the y calculation variables
        DstY = y * yScale
        IntrplY = Int(DstY)
        CalcY = DstY - IntrplY
        For x = 1 To w2 - 2
'Generate the x calculation variables
            DstX = x * xScale
            IntrplX = Int(DstX)
            CalcX = DstX - IntrplX
            
            'Get the 4 pixels around the interpolated one
            r1 = img(IntrplX - 1, IntrplY - 1).rgbRed
            g1 = img(IntrplX - 1, IntrplY - 1).rgbGreen
            b1 = img(IntrplX - 1, IntrplY - 1).rgbBlue
            
            r2 = img(IntrplX, IntrplY - 1).rgbRed
            g2 = img(IntrplX, IntrplY - 1).rgbGreen
            b2 = img(IntrplX, IntrplY - 1).rgbBlue
            
            r3 = img(IntrplX + 1, IntrplY - 1).rgbRed
            g3 = img(IntrplX + 1, IntrplY - 1).rgbGreen
            b3 = img(IntrplX + 1, IntrplY - 1).rgbBlue

            r4 = img(IntrplX - 1, IntrplY).rgbRed
            g4 = img(IntrplX - 1, IntrplY).rgbGreen
            b4 = img(IntrplX - 1, IntrplY).rgbBlue

            r5 = img(IntrplX, IntrplY).rgbRed
            g5 = img(IntrplX, IntrplY).rgbGreen
            b5 = img(IntrplX, IntrplY).rgbBlue
            
            r6 = img(IntrplX + 1, IntrplY).rgbRed
            g6 = img(IntrplX + 1, IntrplY).rgbGreen
            b6 = img(IntrplX + 1, IntrplY).rgbBlue
            
            r7 = img(IntrplX - 1, IntrplY + 1).rgbRed
            g7 = img(IntrplX - 1, IntrplY + 1).rgbGreen
            b7 = img(IntrplX - 1, IntrplY + 1).rgbBlue
            
            r8 = img(IntrplX, IntrplY + 1).rgbRed
            g8 = img(IntrplX, IntrplY + 1).rgbGreen
            b8 = img(IntrplX, IntrplY + 1).rgbBlue
            
            r9 = img(IntrplX + 1, IntrplY + 1).rgbRed
            g9 = img(IntrplX + 1, IntrplY + 1).rgbGreen
            b9 = img(IntrplX + 1, IntrplY + 1).rgbBlue
            r = (r1 + r2 + r3 + r4 + r5 + r7 + r8 + r9) \ 8
            g = (g1 + g2 + g3 + g4 + g5 + g7 + g8 + g9) \ 8
            b = (b1 + b2 + b3 + b4 + b5 + b7 + b8 + b9) \ 8
            'Make sure that the values are in the acceptable range
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            'Set this pixel onto the new picture box
            ret(x, y).rgbRed = r
            ret(x, y).rgbGreen = g
            ret(x, y).rgbBlue = b
        Next x
    Next y
    StdResize2 = ret
On Error GoTo 0
End Function


