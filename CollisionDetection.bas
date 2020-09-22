Attribute VB_Name = "CollisionDetection"
Public Const AccBox As Single = 1
Public Const Acc As Single = 0.5
Public Const Acc2 As Single = 1

Public Function RayCollision2(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR, Optional NotT As Integer = 0, Optional Qual As Single = 1) As Integer
    Dim Dir As D3DVECTOR
    Dim t As Single
    Dir = Normalise(VectorSubtract(PEnd, PStart), t, Qual)
    POut = PStart
    For k = 0 To t
        RayCollision2 = CheckCollision2(POut.X, POut.Y, POut.Z, NotT)
        If RayCollision2 <> 0 Then
            Exit Function
        End If
        POut.X = POut.X + Dir.X
        POut.Y = POut.Y + Dir.Y
        POut.Z = POut.Z + Dir.Z
    Next k
    POut = PEnd
End Function

Public Function CheckCollision2(X As Single, Y As Single, Z As Single, Optional NotT As Integer = 0) As Integer
    CheckCollision2 = 0
    For k = 1 To NumTiles
        If NotT <> k Then
            With Tiles(k)
                If X + AccBox >= .Box.MinX And X - AccBox <= .Box.MaxX Then
                    If Y + AccBox >= .Box.MinY And Y - AccBox <= .Box.MaxY Then
                        If Z + AccBox >= .Box.MinZ And Z - AccBox <= .Box.MaxZ Then
                            If Abs(DotProduct(MakeVector(X, Y, Z), .Normal) - DotProduct(.Vertices(1).position, .Normal)) <= Acc2 Then
                                CheckCollision2 = k
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        End If
    Next k
End Function

Public Function RayCollision(PStart As D3DVECTOR, PEnd As D3DVECTOR, POut As D3DVECTOR, Optional NotT As Integer = 0, Optional Qual As Single = 1) As Integer
    Dim Dir As D3DVECTOR
    Dim t As Single
    Dir = Normalise(VectorSubtract(PEnd, PStart), t, Qual)
    POut = PStart
    For k = 0 To t
        RayCollision = CheckCollision(POut.X, POut.Y, POut.Z, NotT)
        If RayCollision <> 0 Then
            Exit Function
        End If
        POut.X = POut.X + Dir.X
        POut.Y = POut.Y + Dir.Y
        POut.Z = POut.Z + Dir.Z
    Next k
    POut = PEnd
End Function

Public Function CheckCollision(X As Single, Y As Single, Z As Single, Optional NotT As Integer = 0) As Integer
    CheckCollision = 0
    For k = 1 To NumTiles
        If NotT <> k Then
            With Tiles(k)
                If X > .Box.MinX And X < .Box.MaxX Then
                    If Y > .Box.MinY And Y < .Box.MaxY Then
                        If Z > .Box.MinZ And Z < .Box.MaxZ Then
                            If Abs(DotProduct(MakeVector(X, Y, Z), .Normal) - DotProduct(.Vertices(1).position, .Normal)) <= Acc Then
                                CheckCollision = k
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        End If
    Next k
End Function


