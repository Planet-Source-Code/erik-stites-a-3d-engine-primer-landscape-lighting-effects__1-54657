VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Try 3D"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLiteZ 
      Height          =   330
      Left            =   8595
      TabIndex        =   13
      Text            =   "50"
      Top             =   5445
      Width           =   600
   End
   Begin VB.TextBox txtLiteY 
      Height          =   330
      Left            =   7965
      TabIndex        =   11
      Text            =   "50"
      Top             =   5445
      Width           =   600
   End
   Begin VB.TextBox txtLiteX 
      Height          =   330
      Left            =   7335
      TabIndex        =   8
      Text            =   "50"
      Top             =   5445
      Width           =   600
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "Render Terrain"
      Height          =   510
      Left            =   7830
      TabIndex        =   7
      Top             =   4590
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Load 3D Coordinates"
      Height          =   510
      Left            =   6345
      TabIndex        =   6
      Top             =   4590
      Width           =   1320
   End
   Begin VB.PictureBox pbxMap 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4950
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   4590
      Width           =   1020
   End
   Begin VB.PictureBox pbxDebug 
      Height          =   4380
      Left            =   4860
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   282
      TabIndex        =   4
      Top             =   90
      Width           =   4290
   End
   Begin VB.PictureBox pbxDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   4425
      Left            =   90
      ScaleHeight     =   291
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   90
      Width           =   4335
   End
   Begin VB.CommandButton cmdDrawTri 
      Caption         =   "Draw 3D Cube"
      Height          =   510
      Left            =   2880
      TabIndex        =   1
      Top             =   5085
      Width           =   1545
   End
   Begin VB.HScrollBar hscrAngle 
      Height          =   420
      Left            =   900
      Max             =   -179
      Min             =   179
      TabIndex        =   0
      Top             =   4635
      Width           =   3525
   End
   Begin VB.Label Label5 
      Caption         =   "Z"
      Height          =   195
      Left            =   8685
      TabIndex        =   14
      Top             =   5220
      Width           =   330
   End
   Begin VB.Label Label4 
      Caption         =   "Y"
      Height          =   195
      Left            =   8055
      TabIndex        =   12
      Top             =   5220
      Width           =   330
   End
   Begin VB.Label Label3 
      Caption         =   "X"
      Height          =   195
      Left            =   7425
      TabIndex        =   10
      Top             =   5220
      Width           =   330
   End
   Begin VB.Label Label2 
      Caption         =   "Terrain Light Position"
      Height          =   555
      Left            =   6255
      TabIndex        =   9
      Top             =   5220
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Rotate Around Y"
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   4635
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MapW As Integer = 64
Const MapH As Integer = 64

Dim Cube(0 To 11) As TRI3D
Dim PLite As VEC3D 'Point light
Dim TerrLite As VEC3D 'Terrain Lighting
Dim Map(0 To MapW - 1, 0 To MapH - 1) As VEC3D

Private Sub cmdBuild_Click()
    Dim X As Single, Y As Single
    
    For Y = 0 To MapH - 1
    
        For X = 0 To MapW - 1
            
            'Added scaling to make it look a little better
            'the values are lowered to appear smoother
            Map(X, Y).X = X
            Map(X, Y).Y = Y
            Map(X, Y).Z = (pbxMap.Point(X, Y) And 255) / 10
        
        Next
        
    Next
    
End Sub

Private Sub cmdDrawTri_Click()
    Dim RotTri As TRI3D 'Used for rotations
    Dim tmpTri As TRI3D
    Dim tmpVec1 As VEC3D 'used to calc normals
    Dim tmpVec2 As VEC3D '---^
    Dim Norm As VEC3D
    Dim Brightness As VEC3D
    Dim B As Single 'temporary multiplier for brightness
    Dim col As Long
    Dim Vis As Single
    Dim Perc As Single
    Dim i As Integer
    
    pbxDraw.Picture = Nothing
    pbxDebug.Cls
    
    '======
    For i = 0 To 11
    
        Perc = hscrAngle.Value / 179
        RotTri = RotateTri(Cube(i), Perc * PI)
        
        'get the vectors for the plane
        'here we are getting two vectors that are flat against the triangle
        'we use the crossproduct to find a vector perpendicular to
        'both of these vectors. This resulting vector is the normal of this triangle
        tmpVec1.X = RotTri.Coord3(0).X - RotTri.Coord3(1).X
        tmpVec1.Y = RotTri.Coord3(0).Y - RotTri.Coord3(1).Y
        tmpVec1.Z = RotTri.Coord3(0).Z - RotTri.Coord3(1).Z
        
        tmpVec2.X = RotTri.Coord3(2).X - RotTri.Coord3(1).X
        tmpVec2.Y = RotTri.Coord3(2).Y - RotTri.Coord3(1).Y
        tmpVec2.Z = RotTri.Coord3(2).Z - RotTri.Coord3(1).Z
        
        
        'get the normal of the plane
        Norm = CrossProduct(tmpVec1, tmpVec2)
        Norm = Normalize(Norm)
        
        
        'Get the normal of the camera, based on a point on the triangle
        Camera.X = CAMX - RotTri.Coord3(0).X
        Camera.Y = CAMY - RotTri.Coord3(0).Y
        Camera.Z = CAMZ - RotTri.Coord3(0).Z
        Camera = Normalize(Camera)
        
        'Used for BackFace test
        Vis = DotProduct(Norm, Camera)
        
        'Shows the resulting dot product
        pbxDebug.Print Format(Vis, "0.000")
        
        
        'If the angle between the camera and the normal of the current face
        'is less than 90 degrees then we draw the face
        'Remember that the returned value of the DotProduct is the Cosine of
        'the angle between vectors, to get the angle, the inverse cosine function
        'would be used. For most 3D calculations, the actual angle is not needed
        '
        '      +Y
        '       |
        'vis<=0 | vis>0
        '     \ | /
        '-Z----------- Cam(+Z)
        '       |
        '       |
        '      -Y
        If Vis > 0 Then 'Only render non-backfaces
            
            'get vector from light to polygon
            'basically, this is a visibility test from the light source instead
            'of the camera, and our shaded color is B (value between 0 and 1)
            'times the original color
            Brightness.X = PLite.X - RotTri.Coord3(0).X
            Brightness.Y = PLite.Y - RotTri.Coord3(0).Y
            Brightness.Z = PLite.Z - RotTri.Coord3(0).Z
            Brightness = Normalize(Brightness)
            B = DotProduct(Norm, Brightness)
            If B < 0 Then
                B = 0
            End If
            
            With Cube(i).Color
                col = RGB(.R * B, .G * B, .B * B)
            End With
            'col = 255 *Vis 'assumes light is on the camera
        
            TriToScreen RotTri, 1
            
            pbxDraw.FillColor = col
        
            Polygon pbxDraw.hdc, RotTri.Coord2(0), 3
            
        End If
    
    Next
        
    pbxDraw.Refresh
End Sub

Private Sub cmdRender_Click()
    Dim X As Single, Y As Single
    Dim TmpTri1 As TRI3D
    Dim TmpTri2 As TRI3D
    Dim tmpVec1 As VEC3D
    Dim tmpVec2 As VEC3D
    Dim Norm As VEC3D
    Dim Bright As VEC3D
    Dim B As Single
    Dim col As Long
    
    For Y = 0 To MapH - 2
    
        For X = 0 To MapW - 2
            
            'Going to leave the map green
            'Assume all polys of terrain are visible because I did not change the view
            'to add this, use the same type of visibility detection as the cube
            '2 triangles for each 4x4 pixel block
            'This is going to render with Y increasing downward
            'Why? Because it was simpler...
            
            With TmpTri1
            
                'Same basic steps as with the cube, this time
                'we are drawing 2 triangles at once and our
                'light source is different
                .Coord3(0) = Map(X, Y)
                .Coord3(1) = Map(X, Y + 1)
                .Coord3(2) = Map(X + 1, Y)
                
                tmpVec1.X = .Coord3(0).X - .Coord3(1).X
                tmpVec1.Y = .Coord3(0).Y - .Coord3(1).Y
                tmpVec1.Z = .Coord3(0).Z - .Coord3(1).Z
                
                tmpVec2.X = .Coord3(2).X - .Coord3(1).X
                tmpVec2.Y = .Coord3(2).Y - .Coord3(1).Y
                tmpVec2.Z = .Coord3(2).Z - .Coord3(1).Z
                
                Norm = CrossProduct(tmpVec1, tmpVec2)
                Norm = Normalize(Norm)
                
                
                Bright.X = TerrLite.X - .Coord3(0).X
                Bright.Y = TerrLite.Y - .Coord3(0).Y
                Bright.Z = TerrLite.Z - .Coord3(0).Z
                Bright = Normalize(Bright)
                B = DotProduct(Norm, Bright)
                If B < 0 Then
                    B = 0
                End If
                
                col = B * 255
                .Color.R = 0
                .Color.G = col
                .Color.B = 0
                
                pbxDraw.FillColor = RGB(.Color.R, .Color.G, .Color.B)
                
            End With
            
            TriToScreen TmpTri1, 0
            Polygon pbxDraw.hdc, TmpTri1.Coord2(0), 3
            
            
            With TmpTri2
            
                .Coord3(0) = Map(X + 1, Y)
                .Coord3(1) = Map(X, Y + 1)
                .Coord3(2) = Map(X + 1, Y + 1)
                
                tmpVec1.X = .Coord3(0).X - .Coord3(1).X
                tmpVec1.Y = .Coord3(0).Y - .Coord3(1).Y
                tmpVec1.Z = .Coord3(0).Z - .Coord3(1).Z
                
                tmpVec2.X = .Coord3(2).X - .Coord3(1).X
                tmpVec2.Y = .Coord3(2).Y - .Coord3(1).Y
                tmpVec2.Z = .Coord3(2).Z - .Coord3(1).Z
                
                Norm = CrossProduct(tmpVec1, tmpVec2)
                Norm = Normalize(Norm)
                
                
                Bright.X = TerrLite.X - .Coord3(0).X
                Bright.Y = TerrLite.Y - .Coord3(0).Y
                Bright.Z = TerrLite.Z - .Coord3(0).Z
                Bright = Normalize(Bright)
                B = DotProduct(Norm, Bright)
                If B < 0 Then
                    B = 0
                End If
                
                col = B * 255
                .Color.R = 0
                .Color.G = col
                .Color.B = 0
                
                pbxDraw.FillColor = RGB(.Color.R, .Color.G, .Color.B)
                
            End With
            
            TriToScreen TmpTri2, 0
            Polygon pbxDraw.hdc, TmpTri2.Coord2(0), 3
        
        Next
        
    Next
    
    pbxDraw.Refresh
End Sub

Private Sub Form_Load()
    MidX = pbxDraw.ScaleWidth \ 2
    MidY = pbxDraw.ScaleHeight \ 2
    
    
    'White lights, colors are dependant on the object
    PLite.X = 10
    PLite.Y = 30
    PLite.Z = 50
    
    TerrLite.X = 50
    TerrLite.Y = 50
    TerrLite.Z = 50
    
    
    'Six different sides for our cube
    'Each side has a different color
    Cube(0) = CreateTri(-5, -5, -5, 5, -5, -5, -5, -5, 5)
    Cube(0).Color.R = 255: Cube(0).Color.G = 0: Cube(0).Color.B = 0
    Cube(1) = CreateTri(5, -5, 5, -5, -5, 5, 5, -5, -5)
    Cube(1).Color.R = 255: Cube(1).Color.G = 0: Cube(1).Color.B = 0
    
    
    Cube(2) = CreateTri(-5, 5, -5, -5, 5, 5, 5, 5, -5)
    Cube(2).Color.R = 0: Cube(2).Color.G = 255: Cube(2).Color.B = 0
    Cube(3) = CreateTri(5, 5, 5, 5, 5, -5, -5, 5, 5)
    Cube(3).Color.R = 0: Cube(3).Color.G = 255: Cube(3).Color.B = 0
    
    
    Cube(4) = CreateTri(5, 5, -5, 5, -5, -5, 5, 5, 5)
    Cube(4).Color.R = 0: Cube(4).Color.G = 0: Cube(4).Color.B = 255
    Cube(5) = CreateTri(5, -5, 5, 5, 5, 5, 5, -5, -5)
    Cube(5).Color.R = 0: Cube(5).Color.G = 0: Cube(5).Color.B = 255
    
    
    Cube(6) = CreateTri(-5, 5, -5, -5, 5, 5, -5, -5, -5)
    Cube(6).Color.R = 0: Cube(6).Color.G = 255: Cube(6).Color.B = 255
    Cube(7) = CreateTri(-5, -5, 5, -5, -5, -5, -5, 5, 5)
    Cube(7).Color.R = 0: Cube(7).Color.G = 255: Cube(7).Color.B = 255
    
    
    Cube(8) = CreateTri(-5, 5, -5, -5, -5, -5, 5, 5, -5)
    Cube(8).Color.R = 255: Cube(8).Color.G = 0: Cube(8).Color.B = 255
    Cube(9) = CreateTri(5, -5, -5, 5, 5, -5, -5, -5, -5)
    Cube(9).Color.R = 255: Cube(9).Color.G = 0: Cube(9).Color.B = 255
    
    
    Cube(10) = CreateTri(-5, 5, 5, 5, 5, 5, -5, -5, 5)
    Cube(10).Color.R = 255: Cube(10).Color.G = 255: Cube(10).Color.B = 0
    Cube(11) = CreateTri(5, -5, 5, -5, -5, 5, 5, 5, 5)
    Cube(11).Color.R = 255: Cube(11).Color.G = 255: Cube(11).Color.B = 0

End Sub

Private Sub hscrAngle_Change()
    cmdDrawTri_Click
End Sub

Private Sub hscrAngle_Scroll()
    cmdDrawTri_Click
End Sub

'Change Terrain Lighting
Private Sub txtLiteX_Change()
    If Not txtLiteX.Text = "" Then
        If IsNumeric(txtLiteX.Text) Then
        
            TerrLite.X = txtLiteX.Text
            
        End If
    End If
End Sub

Private Sub txtLiteY_Change()
    If Not txtLiteY.Text = "" Then
        If IsNumeric(txtLiteY.Text) Then
        
            TerrLite.Y = txtLiteY.Text
            
        End If
    End If
End Sub

Private Sub txtLiteZ_Change()
    If Not txtLiteZ.Text = "" Then
        If IsNumeric(txtLiteZ.Text) Then
        
            TerrLite.Z = txtLiteZ.Text
            
        End If
    End If
End Sub
