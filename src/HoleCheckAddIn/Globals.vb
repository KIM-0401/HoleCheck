Option Strict On
Option Explicit On

' W07 7-2 · 여러 파일이 함께 쓰는 상수 (AxisMatching.vb · HoleCheckForm.vb에서 이동 · 중복은 하나로)
Friend Module Globals

    ' API 길이 단위 cm → 내부 mm (W04 결정)
    ' 내부 mm → API cm (W04 결정의 역변환)
    Friend Const dblCentimeterToMillimeter As Double = 10.0

    ' 폼 공차 입력 시작값 · W07 매칭 미연결 기간에는 입력 검증에만 쓰인다 (W08 Cap/Tol 로 이어짐)
    Friend Const dblDefaultCandidateRange As Double = 2.0
    Friend Const dblDefaultAlignmentTolerance As Double = 0.1

    ' W07 · 같은 리프 안 동심 묶음 기준 (mm)
    Friend Const dblConcentricClusterThreshold As Double = 0.01

End Module
