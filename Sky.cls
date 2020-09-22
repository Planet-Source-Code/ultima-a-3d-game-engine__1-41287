VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WVB As Direct3DVertexBuffer8
Private Texture As Direct3DTexture8
Private mVerts() As VERTEX
Private NumVerts As Integer
Private LastEyePos As D3DVECTOR

Public Function Render()
    D3DVertexBuffer8SetData WVB, 0, VertexSizeInBytes * NumVerts, 0, mVerts(1)
    D3DDevice.SetStreamSource 0, WVB, VertexSizeInBytes
    D3DDevice.SetVertexShader D3DFVF_VERTEX
    D3DDevice.SetTexture 0, Texture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, NumVerts / 3
End Function

Public Function UpDate()
    For k = 1 To 30
        With mVerts(k).position
            .X = .X + (EyePos.X - LastEyePos.X)
            .Y = .Y + (EyePos.Y - LastEyePos.Y)
            .Z = .Z + (EyePos.Z - LastEyePos.Z)
        End With
    Next k
    LastEyePos = EyePos
End Function

Public Function tCreate(FPath As String)
    NumVerts = 30
    ReDim mVerts(1 To NumVerts) As VERTEX
    
    mVerts(1) = MakeVertex(-500, 500, 500, 0, 1)
    mVerts(2) = MakeVertex(500, 500, 500, 1, 1)
    mVerts(3) = MakeVertex(-500, 0, 500, 0, 0)
    mVerts(4) = mVerts(2)
    mVerts(5) = MakeVertex(500, 0, 500, 1, 0)
    mVerts(6) = mVerts(3)
    mVerts(7) = mVerts(2)
    mVerts(8) = mVerts(5)
    mVerts(9) = MakeVertex(500, 500, -500, 0, 1)
    mVerts(10) = mVerts(5)
    mVerts(11) = mVerts(9)
    mVerts(12) = MakeVertex(500, 0, -500, 0, 0)
    mVerts(13) = mVerts(9)
    mVerts(14) = mVerts(12)
    mVerts(15) = MakeVertex(-500, 500, -500, 1, 1)
    mVerts(16) = mVerts(15)
    mVerts(17) = mVerts(12)
    mVerts(18) = MakeVertex(-500, 0, -500, 1, 0)
    mVerts(19) = mVerts(18)
    mVerts(20) = mVerts(15)
    mVerts(21) = mVerts(1)
    mVerts(22) = mVerts(18)
    mVerts(23) = mVerts(3)
    mVerts(24) = mVerts(1)
    mVerts(25) = mVerts(1)
    mVerts(26) = mVerts(2)
    mVerts(27) = mVerts(9)
    mVerts(28) = mVerts(1)
    mVerts(29) = mVerts(9)
    mVerts(30) = mVerts(15)
    
    For k = 25 To 30
        mVerts(k).tu = 0.5
        mVerts(k).tv = 1
    Next k
    
    For k = 1 To 30
        mVerts(k).color = RGB(200, 200, 200)
    Next k
    
    Set WVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 30, 0, D3DFVF_VERTEX, D3DPOOL_DEFAULT)
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, FPath)
End Function
