Option Strict On
Option Explicit On

' W07 7-2 · 기하 계산 · 출력 서식 도우미 (AxisMatching.vb에서 이동)
Friend Module GeometryHelpers

    ' A 축 위의 점에서 B 축 직선까지 거리 · |(PA - PB) × uB| (uB는 단위벡터) · skew line 일반 공식은 쓰지 않는다
    Friend Function GetPointToLineDistance(ByVal objAxisA As CylinderAxis, ByVal objAxisB As CylinderAxis) As Double
        Dim dblDeltaX As Double = objAxisA.dblPointX - objAxisB.dblPointX
        Dim dblDeltaY As Double = objAxisA.dblPointY - objAxisB.dblPointY
        Dim dblDeltaZ As Double = objAxisA.dblPointZ - objAxisB.dblPointZ

        Dim dblCrossX As Double = dblDeltaY * objAxisB.dblDirectionZ - dblDeltaZ * objAxisB.dblDirectionY
        Dim dblCrossY As Double = dblDeltaZ * objAxisB.dblDirectionX - dblDeltaX * objAxisB.dblDirectionZ
        Dim dblCrossZ As Double = dblDeltaX * objAxisB.dblDirectionY - dblDeltaY * objAxisB.dblDirectionX

        Return System.Math.Sqrt(dblCrossX * dblCrossX + dblCrossY * dblCrossY + dblCrossZ * dblCrossZ)
    End Function

    Friend Function JoinTab(ParamArray strValues As String()) As String
        Return String.Join(System.Convert.ToChar(9).ToString(), strValues)
    End Function

    Friend Function FormatDouble(ByVal dblValue As Double) As String
        Return dblValue.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    ' W07 7-4 · 단위 변환 · API cm ↔ 내부 mm (W04 결정) · 이번 단계에서는 호출하지 않는다
    Friend Function ConvertCentimeterToMillimeter(ByVal dblValue As Double) As Double
        Return dblValue * dblCentimeterToMillimeter
    End Function

    Friend Function ConvertMillimeterToCentimeter(ByVal dblValue As Double) As Double
        Return dblValue / dblCentimeterToMillimeter
    End Function

    ' W07 7-4 · 두 점 사이 3D 거리 · mm 입력 · mm 출력 · 이번 단계에서는 호출하지 않는다
    Friend Function GetDistance3D(ByVal dblXA As Double, ByVal dblYA As Double, ByVal dblZA As Double, ByVal dblXB As Double, ByVal dblYB As Double, ByVal dblZB As Double) As Double
        Dim dblDeltaX As Double = dblXA - dblXB
        Dim dblDeltaY As Double = dblYA - dblYB
        Dim dblDeltaZ As Double = dblZA - dblZB
        Return System.Math.Sqrt(dblDeltaX * dblDeltaX + dblDeltaY * dblDeltaY + dblDeltaZ * dblDeltaZ)
    End Function

End Module
