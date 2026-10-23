Option Strict On
Option Explicit On

' W07 7-2 · 결과·중간 데이터 클래스 (AxisMatching.vb에서 이동)

' 결과 한 줄 · 길이는 전부 mm, 좌표는 조립품 좌표
' 그리드 바인딩(DataGridView)이 공개 속성만 읽으므로 자동 구현 속성으로 둔다
Public Class AxisMatchResult

    Public Const strErrorKindOk As String = "OK"
    Public Const strErrorKindPositionMismatch As String = "위치 불일치"

    Public Property strErrorKind As String = ""
    Public Property dblAxisDistance As Double
    Public Property dblDiameterA As Double
    Public Property dblDiameterB As Double
    Public Property strOccurrenceNameA As String = ""
    Public Property strOccurrenceNameB As String = ""
    Public Property dblTargetX As Double
    Public Property dblTargetY As Double
    Public Property dblTargetZ As Double

End Class

' 원통면 하나의 축 · 숫자만 보관 (mm · 조립품 좌표)
Friend Class CylinderAxis
    Public intLeafIndex As Integer
    Public strOccurrenceName As String
    Public dblPointX As Double
    Public dblPointY As Double
    Public dblPointZ As Double
    Public dblDirectionX As Double
    Public dblDirectionY As Double
    Public dblDirectionZ As Double
    Public dblRadius As Double
    ' W06 · 피처 판정 결과 · 문자열만 보관
    Public strFeatureState As String = ""
    Public strHoleFeatureName As String = ""
    Public strFeaturePath As String = ""
    ' W06 · 색 표시용 Face 프록시 · 이 모듈 밖으로는 호출자가 넘긴 목록으로만 내보내고 AxisMatchResult에는 넣지 않는다
    Public invFace As Inventor.Face
End Class

' 리프 하나의 수집 결과 · 진단 출력용
Friend Class LeafSummary
    Public strOccurrenceName As String
    Public strDefinitionName As String
    Public intCylinderCount As Integer
    Public intFailedFaceCount As Integer
    Public strLeafError As String = ""
End Class

' 진단 출력용 카운터
Friend Class MatchCounters
    Public intCylinderCount As Integer
    Public intFailedFaceCount As Integer
    Public intComparisonCount As Integer
    Public intNotParallelCount As Integer
    Public intOutOfRangeCount As Integer
    Public intOkCount As Integer
    Public intPositionMismatchCount As Integer
    ' W06 · 피처 필터
    Public intHoleCount As Integer
    Public intNotHoleCount As Integer
    Public intFeatureReadFailedCount As Integer
    Public intOuterRingCount As Integer
End Class
