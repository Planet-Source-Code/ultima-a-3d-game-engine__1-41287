Attribute VB_Name = "Movement"
Public Function MoveCamera(mFor As Single, mSide As Single)
    Dim VPos As D3DVECTOR
    If mFor <> 0 Then
        EyePos.Y = EyePos.Y + THei
        VPos = MakeVector(EyePos.X + (EyeDir.X * mFor), EyePos.Y, EyePos.Z + (EyeDir.Z * mFor))
        If RayCollision2(EyePos, VPos, EyePos, 0, 0.05) <> 0 Then
            EyePos.Y = EyePos.Y - THei
        Else
            EyePos = VPos
            VPos.Y = VPos.Y - THei
            RayCollision2 EyePos, VPos, EyePos, 0, 0.05
        End If
    End If
    If mSide <> 0 Then
        EyePos.Y = EyePos.Y + THei
        VPos = MakeVector(EyePos.X + (EyeDir.Z * mSide), EyePos.Y, EyePos.Z + (-EyeDir.X * mSide))
        If RayCollision2(EyePos, VPos, EyePos, 0, 0.05) <> 0 Then
            EyePos.Y = EyePos.Y - THei
        Else
            EyePos = VPos
            VPos.Y = VPos.Y - THei
            RayCollision2 EyePos, VPos, EyePos, 0, 0.05
        End If
    End If
End Function

Public Function RotateCamera(Alpha As Single, Beta As Single)
    Dim MatWorld2 As D3DMATRIX, MatWorld As D3DMATRIX
    Dim yy As Single
    D3DXMatrixRotationAxis MatWorld2, MakeVector(EyeDir.Z, 0, -EyeDir.X), Sin(Beta)
    D3DXMatrixRotationY MatWorld, Sin(Alpha)
    D3DXMatrixMultiply MatWorld, MatWorld, MatWorld2
    With EyeDir
        yy = .X * MatWorld.m12 + .Y * MatWorld.m22 + .Z * MatWorld.m32
        If yy < -0.9 Then
            yy = -0.9
        ElseIf yy > 0.9 Then
            yy = 0.9
        End If
        EyeDir = MakeVector(.X * MatWorld.m11 + .Y * MatWorld.m21 + .Z * MatWorld.m31, _
                                yy, _
                                .X * MatWorld.m13 + .Y * MatWorld.m23 + .Z * MatWorld.m33)
        EyeDir = Normalise(EyeDir)
    End With
End Function

