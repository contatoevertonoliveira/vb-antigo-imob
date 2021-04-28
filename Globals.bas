Attribute VB_Name = "Globals"
Option Explicit

Public AlertCount As Integer

Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Public Sub DrawAngle(picDraw As PictureBox, ByVal fAngle As Single)

Dim iSize       As Integer
Dim iFillStyle  As Integer
Dim lFillColor  As Long
Dim lForeColor  As Long
Dim lRet        As Long
Dim uaPts(3)    As PointAPI

    'Size arrow to best fit picDraw at any angle
    iSize = IIf(picDraw.ScaleHeight < picDraw.ScaleWidth, Int(picDraw.ScaleHeight / PI), Int(picDraw.ScaleWidth / PI))
    
    'Setup the 4 points of the arrow using the first point
    'as the center and the other points offset from the center.
    uaPts(0).X = picDraw.ScaleWidth / 2
    uaPts(0).Y = picDraw.ScaleHeight / 2
    uaPts(1).X = uaPts(0).X - iSize
    uaPts(1).Y = uaPts(0).Y - iSize
    uaPts(2).X = uaPts(0).X + iSize
    uaPts(2).Y = uaPts(0).Y
    uaPts(3).X = uaPts(0).X - iSize
    uaPts(3).Y = uaPts(0).Y + iSize
    
    'Rotate the arrow to the correct angle
    Call RotatePoints(uaPts(0), uaPts, fAngle)
    
    'Save picDraw settings
    iFillStyle = picDraw.FillStyle
    lFillColor = picDraw.FillColor
    lForeColor = picDraw.ForeColor
    
    'Setup picDraw to fill the arrow
    picDraw.FillStyle = vbFSSolid   'Solid Fill
    picDraw.FillColor = &HFFFFFF    'Inside = White
    picDraw.ForeColor = &H0&        'Border = Black
    
    'Draw the filled arrow
    lRet = Polygon(picDraw.hDC, uaPts(0), 4)
    
    'Restore picDraw settings
    picDraw.FillStyle = iFillStyle
    picDraw.FillColor = lFillColor
    picDraw.ForeColor = lForeColor

    'Free the memory
    Erase uaPts
    
End Sub


Private Sub RotatePoints(uAxisPt As PointAPI, uRotatePts() As PointAPI, fDegrees As Single)

'Rotates an array of PointAPI points around a center point by fDegrees

Dim lIdx        As Long
Dim fDX         As Single
Dim fDY         As Single
Dim fRadians    As Single

    fRadians = fDegrees * RADS
    
    For lIdx = 0 To UBound(uRotatePts)
        fDX = uRotatePts(lIdx).X - uAxisPt.X
        fDY = uRotatePts(lIdx).Y - uAxisPt.Y
        uRotatePts(lIdx).X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        uRotatePts(lIdx).Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    Next lIdx
    
End Sub


