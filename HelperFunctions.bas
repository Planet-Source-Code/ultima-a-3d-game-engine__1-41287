Attribute VB_Name = "HelperFunctions"
Public Type VERTEX
    position As D3DVECTOR
    color As Long
    tu As Single
    tv As Single
End Type

Public Type Material
    Texture As Direct3DTexture8
End Type

Public Type Light
    position As D3DVECTOR
    ColR As Byte
    ColG As Byte
    ColB As Byte
    Range As Single
End Type

Public Type BBox
    MinX As Single
    MinY As Single
    MinZ As Single
    MaxX As Single
    MaxY As Single
    MaxZ As Single
End Type

Public Type Tile
    Vertices(1 To 6) As VERTEX
    Mat As Integer
    Normal As D3DVECTOR
    Box As BBox
End Type

Public Const Pi As Single = 3.14159275180032
Public Const D3DFVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const THei As Single = 5
Public Const BreathSpeed As Single = 2
Public Const Breath As Single = 0.005
Public Const MoveSpeed As Single = 0.5
Public Const ShootRange As Single = 300

Public Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Function AddLight(X As Single, Y As Single, Z As Single, R As Byte, G As Byte, B As Byte, Range As Single)
    NumLights = NumLights + 1
    ReDim Preserve Lights(1 To NumLights) As Light
    With Lights(NumLights)
        .ColB = B
        .ColG = G
        .ColR = R
        .position = MakeVector(X, Y, Z)
        .Range = Range
    End With
End Function

Public Function AddMaterial(FPath As String)
    NumMat = NumMat + 1
    ReDim Preserve Materials(1 To NumMat) As Material
    Set Materials(NumMat).Texture = D3DX.CreateTextureFromFile(D3DDevice, FPath)
End Function

Public Function AddTile2(P1 As VERTEX, P2 As Single, P3 As VERTEX, P4 As Single, Optional Mat As Integer = 0)
    NumTiles = NumTiles + 1
    ReDim Preserve Tiles(1 To NumTiles) As Tile
    With Tiles(NumTiles)
        .Vertices(1) = P1
        .Vertices(4) = P1
        .Vertices(2) = MakeVertex(P1.position.X, P2, P3.position.Z, P3.tu, P1.tv)
        .Vertices(3) = P3
        .Vertices(5) = P3
        .Vertices(6) = MakeVertex(P3.position.X, P4, P1.position.Z, P1.tu, P3.tv)
        .Mat = Mat
        GetBox NumTiles
        GetNormal NumTiles
    End With
End Function

Public Function AddTile(P1 As VERTEX, P2 As Single, P3 As VERTEX, P4 As Single, Optional Mat As Integer = 0)
    NumTiles = NumTiles + 1
    ReDim Preserve Tiles(1 To NumTiles) As Tile
    With Tiles(NumTiles)
        .Vertices(1) = P1
        .Vertices(4) = P1
        .Vertices(2) = MakeVertex(P3.position.X, P2, P1.position.Z, P3.tu, P1.tv)
        .Vertices(3) = P3
        .Vertices(5) = P3
        .Vertices(6) = MakeVertex(P1.position.X, P4, P3.position.Z, P1.tu, P3.tv)
        .Mat = Mat
        GetBox NumTiles
        GetNormal NumTiles
    End With
End Function

Public Function MakeVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single) As VERTEX
    With MakeVertex
        .position.X = X
        .position.Y = Y
        .position.Z = Z
        .tu = U
        .tv = V
    End With
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function Min(A As Single, B As Single) As Single
    If A > B Then
        Min = B
    Else
        Min = A
    End If
End Function

Public Function Max(A As Single, B As Single) As Single
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function

Public Function GetBox(TNum As Integer)
    With Tiles(TNum)
        .Box.MinX = Min(.Vertices(1).position.X, .Vertices(2).position.X)
        .Box.MinY = Min(.Vertices(1).position.Y, Min(.Vertices(3).position.Y, Min(.Vertices(4).position.Y, .Vertices(2).position.Y)))
        .Box.MinZ = Min(.Vertices(1).position.Z, .Vertices(3).position.Z)
        .Box.MaxX = Max(.Vertices(1).position.X, .Vertices(2).position.X)
        .Box.MaxY = Max(.Vertices(1).position.Y, Max(.Vertices(3).position.Y, Max(.Vertices(4).position.Y, .Vertices(2).position.Y)))
        .Box.MaxZ = Max(.Vertices(1).position.Z, .Vertices(3).position.Z)
    End With
End Function

Public Function GetNormal(TNum As Integer)
    With Tiles(TNum)
        Dim VV1 As D3DVECTOR
        Dim VV2 As D3DVECTOR
        VV1 = VectorSubtract(.Vertices(1).position, .Vertices(2).position)
        VV2 = VectorSubtract(.Vertices(3).position, .Vertices(2).position)
        .Normal = Normalise(CrossProduct(VV1, VV2))
    End With
End Function

Public Function Normalise(V As D3DVECTOR, Optional CARRY As Single = 0, Optional Qual As Single = 1) As D3DVECTOR
    Dim Leng As Single
    With Normalise
        Leng = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z) / Qual
        If Leng > 0 Then
            .X = V.X / Leng
            .Y = V.Y / Leng
            .Z = V.Z / Leng
        End If
        CARRY = Leng
    End With
End Function

Public Function VectorSubtract(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With VectorSubtract
        .X = V1.X - V2.X
        .Y = V1.Y - V2.Y
        .Z = V1.Z - V2.Z
    End With
End Function

Public Function CrossProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With CrossProduct
        .X = V1.Y * V2.Z - V1.Z * V2.Y
        .Y = V1.Z * V2.X - V1.X * V2.Z
        .Z = V1.X * V2.Y - V1.Y * V2.X
    End With
End Function

Public Function DotProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function

Public Function Dist(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    Dist = Sqr((V2.X - V1.X) * (V2.X - V1.X) + (V2.Y - V1.Y) * (V2.Y - V1.Y) + (V2.Z - V1.Z) * (V2.Z - V1.Z))
End Function


