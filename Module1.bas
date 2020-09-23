Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Type VEC3D
    X As Single
    Y As Single
    Z As Single
End Type

Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Type TRI3D
    Coord3(0 To 2) As VEC3D
    Coord2(0 To 2) As POINTAPI
    Color As RGBQUAD
End Type

'Camera position
Global Const CAMX As Single = 0
Global Const CAMY As Single = 0
Global Const CAMZ As Single = 300

'Scale of screen position, not space
Global Const SCALESIZE As Single = 4
    
Global Const PI As Single = 3.14159 'Mmmm... PI

'Location of camera and screen midpoints
Global Camera As VEC3D
Global MidX As Single
Global MidY As Single

'Because I didn't want to manualy create a triangle every time when creating this
Public Function CreateTri(X1 As Single, Y1 As Single, Z1 As Single, _
                            X2 As Single, Y2 As Single, Z2 As Single, _
                            X3 As Single, Y3 As Single, Z3 As Single) As TRI3D
    'Pass a set of 3 points via coordinates, to produce a triangle
    With CreateTri.Coord3(0)
        .X = X1
        .Y = Y1
        .Z = Z1
    End With
    
    With CreateTri.Coord3(1)
        .X = X2
        .Y = Y2
        .Z = Z2
    End With
    
    With CreateTri.Coord3(2)
        .X = X3
        .Y = Y3
        .Z = Z3
    End With
    
End Function

'Rotate a triangle in 3D space around the Y axis
Public Function RotateTri(tri As TRI3D, Angle As Single) As TRI3D
    Dim X As Single, Y As Single, Z As Single
    
    '=====Around Y
    X = Cos(Angle) * tri.Coord3(0).X - Sin(Angle) * tri.Coord3(0).Z
    Z = Sin(Angle) * tri.Coord3(0).X + Cos(Angle) * tri.Coord3(0).Z
    RotateTri.Coord3(0).X = X
    RotateTri.Coord3(0).Y = tri.Coord3(0).Y
    RotateTri.Coord3(0).Z = Z
    
    X = Cos(Angle) * tri.Coord3(1).X - Sin(Angle) * tri.Coord3(1).Z
    Z = Sin(Angle) * tri.Coord3(1).X + Cos(Angle) * tri.Coord3(1).Z
    RotateTri.Coord3(1).X = X
    RotateTri.Coord3(1).Y = tri.Coord3(1).Y
    RotateTri.Coord3(1).Z = Z
    
    X = Cos(Angle) * tri.Coord3(2).X - Sin(Angle) * tri.Coord3(2).Z
    Z = Sin(Angle) * tri.Coord3(2).X + Cos(Angle) * tri.Coord3(2).Z
    RotateTri.Coord3(2).X = X
    RotateTri.Coord3(2).Y = tri.Coord3(2).Y
    RotateTri.Coord3(2).Z = Z
    
End Function

'Project world space onto screen coordinates
Public Sub TriToScreen(tri As TRI3D, Center As Byte)
    Dim Zshift As Single
    
    'If you want to center the rendering around the middle of the picturebox then
    If Center = 1 Then
    
        With tri
            
            Zshift = (CAMZ - .Coord3(0).Z) / CAMZ
            .Coord2(0).X = MidX + .Coord3(0).X * Zshift * SCALESIZE
            .Coord2(0).Y = Form1.pbxDraw.ScaleHeight - MidY + CAMY + .Coord3(0).Y * Zshift * SCALESIZE
            
            Zshift = (CAMZ - .Coord3(1).Z) / CAMZ
            .Coord2(1).X = MidX + .Coord3(1).X * Zshift * SCALESIZE
            .Coord2(1).Y = Form1.pbxDraw.ScaleHeight - MidY + CAMY + .Coord3(1).Y * Zshift * SCALESIZE
            
            Zshift = (CAMZ - .Coord3(2).Z) / CAMZ
            .Coord2(2).X = MidX + .Coord3(2).X * Zshift * SCALESIZE
            .Coord2(2).Y = Form1.pbxDraw.ScaleHeight - MidY + CAMY + .Coord3(2).Y * Zshift * SCALESIZE
            
        End With
        
    Else
    
        With tri
            
            Zshift = (CAMZ - .Coord3(0).Z) / CAMZ
            .Coord2(0).X = .Coord3(0).X * Zshift * SCALESIZE
            .Coord2(0).Y = CAMY + .Coord3(0).Y * Zshift * SCALESIZE
            
            Zshift = (CAMZ - .Coord3(1).Z) / CAMZ
            .Coord2(1).X = .Coord3(1).X * Zshift * SCALESIZE
            .Coord2(1).Y = CAMY + .Coord3(1).Y * Zshift * SCALESIZE
            
            Zshift = (CAMZ - .Coord3(2).Z) / CAMZ
            .Coord2(2).X = .Coord3(2).X * Zshift * SCALESIZE
            .Coord2(2).Y = CAMY + .Coord3(2).Y * Zshift * SCALESIZE
            
        End With
    
    End If
End Sub

'Returns a vector perpendicular to the input vectors
Public Function CrossProduct(V As VEC3D, W As VEC3D) As VEC3D

    CrossProduct.X = V.Y * W.Z - W.Y * V.Z
    CrossProduct.Y = V.Z * W.X - W.Z * V.X
    CrossProduct.Z = V.X * W.Y - W.X * V.Y
    
End Function

'Returns a scalar value representing the Cosine of the angle between two vectors
Public Function DotProduct(V As VEC3D, W As VEC3D) As Single
    
    DotProduct = (V.X * W.X) + (V.Y * W.Y) + (V.Z * W.Z)
    
End Function

'Length of a vector
Public Function Mag(V As VEC3D) As Double
    
    Mag = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
    
End Function

'Makes a vector's length = 1
Public Function Normalize(Vect As VEC3D) As VEC3D
    Dim m As Double
    
    m = Mag(Vect)
    If m = 0 Then m = 1
    
    Normalize.X = (Vect.X / m)
    Normalize.Y = (Vect.Y / m)
    Normalize.Z = (Vect.Z / m)
    
End Function

