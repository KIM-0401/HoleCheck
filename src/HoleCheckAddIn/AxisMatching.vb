Option Strict On
Option Explicit On

' 원통면 축 기반 1차 매칭 · A·B 병렬 작업의 접점 파일
' 지금은 껍데기만 둔다 · RunAxisMatching은 빈 목록을 반환
' 결과에는 숫자와 문자열만 담는다 (COM 객체 금지 · decisions.md "W05 · 결과 구조에는 숫자와 문자열만 담음")

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

Public Module AxisMatching

    ' dblCandidateRange · 후보 범위(mm) · 축간 거리가 이보다 크면 짝 아님
    ' dblAlignmentTolerance · 정렬 공차(mm) · 축간 거리가 이하이면 OK, 초과하면서 후보 범위 이내이면 위치 불일치
    Public Function RunAxisMatching(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal dblCandidateRange As Double, ByVal dblAlignmentTolerance As Double) As System.Collections.Generic.List(Of AxisMatchResult)
        Dim colResults As New System.Collections.Generic.List(Of AxisMatchResult)()
        Return colResults
    End Function

End Module
