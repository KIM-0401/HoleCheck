Option Strict On
Option Explicit On

' W07 · 결과·중간 데이터 클래스 (7-2 에서 AxisMatching.vb 에서 이동 · 7-6 에서 축용 클래스 제거)

' W07 7-4 · EdgeLoop 수집용 데이터 모델 · 숫자와 문자열만 담는다 (W05 결정) · 이번 단계에서는 호출하지 않는다

' 평면 face의 안쪽 원형 EdgeLoop 하나 · 길이는 mm, 좌표는 조립품 좌표
Public Class LoopCircleInfo

    Public Property intLeafIndex As Integer
    Public Property strOccurrenceName As String = ""
    Public Property dblCenterX As Double
    Public Property dblCenterY As Double
    Public Property dblCenterZ As Double
    Public Property dblDiameter As Double
    Public Property intLoopEdgeCount As Integer

End Class

' 같은 리프 안 동심 원 묶음 · 중심은 첫 원 중심(mm · 조립품 좌표) · 지름은 오름차순 리스트
Public Class HoleCluster

    Public Property intLeafIndex As Integer
    Public Property strOccurrenceName As String = ""
    Public Property dblCenterX As Double
    Public Property dblCenterY As Double
    Public Property dblCenterZ As Double
    Public ReadOnly Property colDiameters As New System.Collections.Generic.List(Of Double)()
    Public Property intCircleCount As Integer

    ' 지름을 오름차순 위치에 넣고 원 개수를 하나 늘린다 · 같은 지름도 버리지 않는다
    Public Sub AddDiameter(ByVal dblDiameter As Double)
        Dim intInsertIndex As Integer = colDiameters.BinarySearch(dblDiameter)
        If intInsertIndex < 0 Then
            intInsertIndex = Not intInsertIndex
        End If
        colDiameters.Insert(intInsertIndex, dblDiameter)
        intCircleCount += 1
    End Sub

End Class

' 서로 다른 리프의 묶음 한 쌍 · W08 에서 사용 · 이번 주는 정의만
Public Class MatchPair

    Public Property strOccurrenceNameA As String = ""
    Public Property strOccurrenceNameB As String = ""
    Public Property intClusterIndexA As Integer
    Public Property intClusterIndexB As Integer
    Public Property dblCenterDistance As Double
    Public Property strJudgement As String = ""

End Class

' W07 7-5 · EdgeLoop 수집 리프 하나의 요약 · 진단 출력용 (기존 LeafSummary 와 별도)
Friend Class LoopLeafSummary
    Public strOccurrenceName As String
    Public strDefinitionName As String
    Public intAcceptedCircleCount As Integer
    Public intClusterCount As Integer
    Public intMultiDiameterClusterCount As Integer
    ' W07 7-6 · 리프 단위 탈락 수 · LoopCounters 의 같은 값을 리프마다 나눠 기록 (수집 로직은 그대로)
    Public intRejectedArcCount As Integer
    Public intRejectedOtherCurveCount As Integer
    Public intReadFailedCount As Integer
    Public strLeafError As String = ""
End Class

' W07 7-5 · EdgeLoop 수집 카운터 · 진단 출력용 (기존 MatchCounters 와 별도)
Friend Class LoopCounters
    Public intLeafCount As Integer
    Public intPlaneFaceCount As Integer
    Public intInnerLoopCount As Integer
    Public intAcceptedLoopCount As Integer
    Public intRejectedArcCount As Integer
    Public intRejectedOtherCurveCount As Integer
    Public intReadFailedCount As Integer
    Public intMultiEdgeAcceptedLoopCount As Integer
    Public intClusterCount As Integer
    Public intMultiDiameterClusterCount As Integer
End Class
