Attribute VB_Name = "Lighting"
Public NumLights As Integer
Public Lights() As Light

Public Function DoLighting()
    Dim DD As Single
    Dim CP As D3DVECTOR
    Dim Curr(1 To 3) As Single, NC As Long
    If NumLights > 0 Then
        For t = 1 To NumTiles
            For l = 1 To 6
                Curr(1) = 0
                Curr(2) = 0
                Curr(3) = 0
                NC = 0
                For k = 1 To NumLights
                    NC = NC + 1
                    With Tiles(t)
                        DD = Dist(.Vertices(l).position, Lights(k).position)
                        If DD <= Lights(k).Range Then
                            If RayCollision(.Vertices(l).position, Lights(k).position, CP, CInt(t)) = 0 Then
                                DD = 1 - (DD / Lights(k).Range)
                                DD = DD * Abs(DotProduct(.Normal, Normalise(VectorSubtract(.Vertices(l).position, Lights(k).position))))
                                Curr(1) = Lights(k).ColR * DD + Curr(1)
                                Curr(2) = Lights(k).ColG * DD + Curr(2)
                                Curr(3) = Lights(k).ColB * DD + Curr(3)
                            Else
                                Curr(1) = 10 + Curr(1)
                                Curr(2) = 10 + Curr(2)
                                Curr(3) = 10 + Curr(3)
                            End If
                        Else
                            Curr(1) = 10 + Curr(1)
                            Curr(2) = 10 + Curr(2)
                            Curr(3) = 10 + Curr(3)
                        End If
                    End With
                Next k
                Curr(1) = Curr(1) / NC
                Curr(2) = Curr(2) / NC
                Curr(3) = Curr(3) / NC
                Tiles(t).Vertices(l).color = RGB(Curr(1), Curr(2), Curr(3))
            Next l
        Next t
    Else
        For k = 1 To NumTiles
            For t = 1 To 6
                With Tiles(k).Vertices(t)
                    .color = RGB(200, 200, 200)
                End With
            Next t
        Next k
    End If
End Function

