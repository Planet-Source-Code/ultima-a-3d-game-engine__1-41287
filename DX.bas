Attribute VB_Name = "DX"
Public DX As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public DS As DirectSound8
Public D3DDevice As Direct3DDevice8

Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public DIState As DIKEYBOARDSTATE

Public Function Init() As Boolean

    CHeight = 3
    
    Set D3D = DX.Direct3DCreate()
    If D3D Is Nothing Then Exit Function
    
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    Dim D3DPP As D3DPRESENT_PARAMETERS
    
    With D3DPP
        .Windowed = 0
        .BackBufferHeight = 600
        .BackBufferWidth = 800
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = frmMain.HWND
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.HWND, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    If D3DDevice Is Nothing Then Exit Function
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    HelpInit
    SetupMatrices
    InitBullet
    
    Set mVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 6, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
    
    InitDI
    InitDS
    
    Init = True
End Function

Public Function SetupMatrices()
    Dim matView As D3DMATRIX
    
    EyePos = MakeVector(0, 10, 0)
    EyeDir = MakeVector(0, 0, 1)
    EyeDir = Normalise(EyeDir)
    D3DXMatrixLookAtLH matView, EyePos, MakeVector(EyeDir.X + EyePos.X, EyeDir.Y + EyePos.Y, EyePos.Z + EyeDir.Z), MakeVector(0, 1, 0)
    
    D3DDevice.SetTransform D3DTS_VIEW, matView

    Dim matProj As D3DMATRIX
    
    D3DXMatrixPerspectiveFovLH matProj, Pi / 3, 1, 1, 1000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Function

Public Function RenderAll()
    Dim matView As D3DMATRIX
    
    D3DXMatrixLookAtLH matView, MakeVector(EyePos.X, EyePos.Y + CHeight, EyePos.Z), MakeVector(EyeDir.X + EyePos.X, EyeDir.Y + EyePos.Y + CHeight, EyePos.Z + EyeDir.Z), MakeVector(0, 1, 0)
    
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    
    D3DDevice.BeginScene
        
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
        D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
        
        For k = 1 To NumTiles
            D3DVertexBuffer8SetData mVB, 0, VertexSizeInBytes * 6, 0, Tiles(k).Vertices(1)
            D3DDevice.SetStreamSource 0, mVB, VertexSizeInBytes
            D3DDevice.SetVertexShader D3DFVF_VERTEX
            If Tiles(k).Mat <> 0 Then
                D3DDevice.SetTexture 0, Materials(Tiles(k).Mat).Texture
            End If
            D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        Next k
        
        If BulVis = True Then
            D3DVertexBuffer8SetData BVB, 0, VertexSizeInBytes * 3, 0, Bullet(1)
            D3DDevice.SetStreamSource 0, BVB, VertexSizeInBytes
            D3DDevice.SetVertexShader D3DFVF_VERTEX
            D3DDevice.SetTexture 0, BulletTex
            D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
        End If
        
        TSky.Render
        
    D3DDevice.EndScene
    
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function HelpInit()
    Dim TempV As VERTEX
    
    VertexSizeInBytes = Len(TempV)
End Function

Public Function CheckKey()
    DIDevice.GetDeviceStateKeyboard DIState

    Dim TMov As Boolean, TTMov As Boolean

    TMov = Moving
    TTMov = False
    
    With DIState
        If .Key(DIK_W) <> 0 Then
            MoveCamera MoveSpeed, 0
            Moving = True
            TTMov = True
        End If
        If .Key(DIK_S) <> 0 Then
            MoveCamera -MoveSpeed, 0
            Moving = True
            TTMov = True
        End If
        If .Key(DIK_D) <> 0 Then
            MoveCamera 0, MoveSpeed
            Moving = True
            TTMov = True
        End If
        If .Key(DIK_A) <> 0 Then
            MoveCamera 0, -MoveSpeed
            Moving = True
            TTMov = True
        End If
        If .Key(DIK_LCONTROL) <> 0 Then
            CtrDown = True
        Else
            CtrDown = False
        End If
        If .Key(DIK_SPACE) <> 0 Then
            If StandOn <> 0 Then
                ToFly = 10
            End If
        End If
        If .Key(DIK_ESCAPE) <> 0 Then
            End
        End If
        If .Key(DIK_P) <> 0 Then
            For k = 1 To 800
                For t = 1 To 600
                    frmMain.SS.PSet (k, t), frmMain.Point(k, t)
                Next t
            Next k
            frmMain.SS.Refresh
            SavePicture frmMain.SS.Image, App.Path + "\test.bmp"
        End If
    End With
    
    If TTMov = False Then Moving = False
    If TMov = False And Moving = True Then WalkS.PlaySound True
    If Moving = False Then WalkS.StopSound

    
    If EyePos.X >= 198 Then EyePos.X = 198
    If EyePos.Z >= 198 Then EyePos.Z = 198
    If EyePos.X <= -158 Then EyePos.X = -158
    If EyePos.Z <= -158 Then EyePos.Z = -158
End Function

Public Function InitDS()
    Set DS = DX.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel frmMain.HWND, DSSCL_PRIORITY
End Function

Public Function InitDI()
    Set DI = DX.DirectInputCreate()
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")
    
    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel frmMain.HWND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    DIDevice.Acquire
End Function

Public Function KillApp()
    Set DVB = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    WindS.StopSound
    WalkS.StopSound
    Set WalkS = Nothing
    Set ShootS = Nothing
    Set ZoomS = Nothing
    Set WindS = Nothing
    Set DS = Nothing
    DIDevice.Unacquire
End Function

