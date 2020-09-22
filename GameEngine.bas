Attribute VB_Name = "GameEngine"
Public WalkS As New MSound
Public ShootS As New MSound
Public WindS As New MSound
Public ZoomS As New MSound

Public TSky As New Sky

Public InC As Boolean
Public CtrDown As Boolean
Public ToFly As Single
Public EyePos As D3DVECTOR
Public EyeDir As D3DVECTOR
Public StandOn As Integer
Public CHeight As Single

Public NumTiles As Integer
Public Tiles() As Tile
Public NumMat As Integer
Public Materials() As Material

Public mVB As Direct3DVertexBuffer8
Public VertexSizeInBytes As Long

Public Bullet(1 To 3) As VERTEX
Public BulletTex As Direct3DTexture8
Public BVB As Direct3DVertexBuffer8
Public BulVis As Boolean
Public Life As Single

Public Moving As Boolean

Public Const MQual = 1

Public Function LoadSounds()
    WalkS.Create App.Path + "\move.wav", 0
    ShootS.Create App.Path + "\shot.wav", 0
    WindS.Create App.Path + "\wind.wav", -2500
    ZoomS.Create App.Path + "\zoom.wav", 0
End Function

Public Function InitGame(Lighting As Boolean)
    Dim LS() As Single
    Dim Wid As Integer
    Dim Hei As Integer
   
    Randomize Timer
    
    Wid = 20
    Hei = 20
    
    ReDim LS(1 To Wid, 1 To Hei) As Single
    
    For k = 1 To Wid
        For t = 1 To Hei
            LS(k, t) = Rnd * 10
        Next t
    Next k
    
    AddOutsideWall
    TSky.tCreate App.Path + "\sky.jpg"
    
    If Lighting = True Then
        AddLight 0, 50, 0, 255, 255, 255, 2000
    End If
    
    LoadLandscape LS, Wid, Hei, 20, 2
    
    AddMaterial App.Path + "\gras.jpg"
    
    LoadSounds
    
    WindS.PlaySound True
End Function

Public Function MoveBullet(StartP As D3DVECTOR, EndP As D3DVECTOR)
    Bullet(1).position = StartP
    Bullet(2).position = StartP
    Bullet(3).position = EndP
    Bullet(1).position.X = Bullet(1).position.X + 0.1
    Bullet(2).position.X = Bullet(2).position.X - 0.1
    Bullet(1).position.Z = Bullet(1).position.Z + 0.1
    Bullet(2).position.Z = Bullet(2).position.Z - 0.1
    Life = 2
End Function

Public Function BulletUp()
    If Life > 0 Then
        Life = Life - 1
        If Life = 0 Then
            BulVis = False
        End If
    End If
End Function

Public Function DrawCross()
    For k = 390 To 410
        SetPixelV frmMain.HDC, k, 300, RGB(255, 255, 255)
    Next k
    For k = 290 To 310
        SetPixelV frmMain.HDC, 400, k, RGB(255, 255, 255)
    Next k
End Function

Public Function Shoot()
    Dim EndP As D3DVECTOR
    Dim Hit As D3DVECTOR
    Dim HT As Integer
    Dim CC As Single
    If InC = True Then
        CC = 1
    Else
        CC = 3
    End If
    EndP = MakeVector(EyePos.X + (EyeDir.X * ShootRange), EyePos.Y + 3 + (EyeDir.Y * ShootRange), EyePos.Z + (EyeDir.Z * ShootRange))
    HT = RayCollision(MakeVector(EyePos.X, EyePos.Y + CC, EyePos.Z), EndP, Hit, 0, MQual)
    ShootS.PlaySound False
    MoveBullet MakeVector(EyePos.X + EyeDir.X * 2, EyePos.Y + CC - 0.5 + EyeDir.Y * 2, EyePos.Z + EyeDir.Z * 2), Hit
    BulVis = True
    If HT <> 0 Then
        For k = 1 To 6
            With Tiles(HT).Vertices(k)
                .color = RGB(0, 255, 0)
            End With
        Next k
    End If
End Function

Public Function Gravity()
    StandOn = RayCollision2(EyePos, MakeVector(EyePos.X, EyePos.Y - 0.25, EyePos.Z), EyePos, 0, 0.05)
    If EyePos.Y <= 0 Then EyePos.Y = 1: EyePos.X = EyePos.X + 1
End Function

Public Function Jump()
    If ToFly > 0 Then
        EyePos.Y = EyePos.Y + 1
        ToFly = ToFly - 1
    End If
End Function

Public Function Crouch()
    If CtrDown = True Then
        CHeight = 1
        InC = True
    ElseIf InC = True Then
        InC = False
        CHeight = 3
        EyePos.Y = EyePos.Y + 2
    End If
End Function

Public Function Breathe()
    CHeight = CHeight + Sin(Timer * BreathSpeed) * Breath
    EyePos.Y = EyePos.Y + Sin(Timer * BreathSpeed) * Breath
End Function

Public Function Zoom(InZ As Boolean)
    Dim matProj As D3DMATRIX
    If InZ = True Then
        D3DXMatrixPerspectiveFovLH matProj, Pi / 10, 1, 1, 1000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    Else
        D3DXMatrixPerspectiveFovLH matProj, Pi / 3, 1, 1, 1000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    End If
    ZoomS.PlaySound False
End Function

Public Function AddOutsideWall()
    AddMaterial App.Path + "\wall.jpg"
    AddTile MakeVertex(-160, -10, -160, 0, 0), -10, MakeVertex(200, 50, -160, 1, 1), 50, 1
    AddTile MakeVertex(-160, -10, 200, 0, 0), -10, MakeVertex(200, 50, 200, 1, 1), 50, 1
    AddTile2 MakeVertex(-160, -10, 200, 1, 0), -10, MakeVertex(-160, 50, -160, 0, 1), 50, 1
    AddTile2 MakeVertex(200, -10, 200, 1, 0), -10, MakeVertex(200, 50, -160, 0, 1), 50, 1
End Function

Public Function MLoop()
    CheckKey
    Breathe
    Crouch
    Jump
    Gravity
    BulletUp
    TSky.UpDate
    RenderAll
    DrawCross
End Function

Public Function InitBullet()
    Set BulletTex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\bullet.jpg")
    Set BVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
    BulVis = False
    Bullet(2).tu = 1
    Bullet(3).tu = 0.5
    Bullet(3).tv = 1
    Bullet(1).color = RGB(255, 255, 255)
    Bullet(2).color = RGB(255, 255, 255)
    Bullet(3).color = RGB(255, 255, 255)
End Function

Public Function LoadLandscape(P() As Single, Wid As Integer, Hei As Integer, Optional CSize As Single = 2, Optional Mat As Integer = 0)
    Dim V1 As VERTEX, V2 As VERTEX
    For k = 1 To Wid
        For t = 1 To Hei
            V1.position.X = (k * CSize) - (Wid * CSize / 2)
            V1.position.Z = t * CSize - (Hei * CSize / 2)
            V1.position.Y = P(k, t)
            V1.tu = 0
            V1.tv = 0
            V2.position.X = (k + 1) * CSize - (Wid * CSize / 2)
            V2.position.Z = (t + 1) * CSize - (Hei * CSize / 2)
            V2.tu = 1
            V2.tv = 1
            If Wid > k Then
                If Hei > t Then
                    V2.position.Y = P(k + 1, t + 1)
                    AddTile V1, P(k + 1, t), V2, P(k, t + 1), Mat
                Else
                    V2.position.Y = 0
                    AddTile V1, P(k + 1, t), V2, 0, Mat
                End If
            Else
                V2.position.Y = 0
                If Hei > t Then
                    AddTile V1, 0, V2, P(k, t + 1), Mat
                Else
                    AddTile V1, 0, V2, 0, Mat
                End If
            End If
        Next t
    Next k
    DoLighting
End Function
