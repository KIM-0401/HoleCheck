Option Strict On
Option Explicit On

' W07 · 기하 계산 · 출력 서식 도우미 (7-2 에서 AxisMatching.vb 에서 이동 · 7-6 에서 축 점-직선 거리 제거)
Friend Module GeometryHelpers

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
