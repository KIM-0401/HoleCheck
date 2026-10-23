Option Strict On
Option Explicit On

' W07 7-2 · 여러 파일이 함께 쓰는 상수 (AxisMatching.vb · HoleCheckForm.vb에서 이동 · 중복은 하나로)
Friend Module Globals

    ' API 길이 단위 cm → 내부 mm (W04 결정)
    ' 내부 mm → API cm (W04 결정의 역변환)
    Friend Const dblCentimeterToMillimeter As Double = 10.0

    ' 두 축 방향이 이 각도 이내면 평행으로 본다 (반대 방향 포함)
    Friend Const dblParallelToleranceDegrees As Double = 0.1

    ' 진단 출력 · 같은 occurrence 안 구멍 간격에서 이 거리(mm) 이하는 동축으로 보고 제외
    Friend Const dblCoaxialDistanceThreshold As Double = 0.001

    Friend Const dblDefaultCandidateRange As Double = 2.0
    Friend Const dblDefaultAlignmentTolerance As Double = 0.1

    ' W07 · 같은 리프 안 동심 묶음 기준 (mm)
    Friend Const dblConcentricClusterThreshold As Double = 0.01

End Module
