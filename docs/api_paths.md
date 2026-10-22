# API 접근 경로표 (Inventor 2027 Interop 기준)

| # | 목적 | 확인한 이름 | 타입 | 비고 |
|---|---|---|---|---|
| 1 | 활성 조립품 | | | |
| 2 | 순회(재귀) | ComponentOccurrence.SubOccurrences | ComponentOccurrencesEnumerator | |
| 3 | 순회(평탄) | ComponentOccurrences.AllLeafOccurrences(Optional Object) | ComponentOccurrencesEnumerator | 괄호 형태 속성 |
| 4 | 부품 위치 | ComponentOccurrence.Transformation | Inventor.Matrix | 단위·누적 여부는 실측 |
| 5 | 경계 상자 | ComponentOccurrence.RangeBox | Inventor.Box | MinPoint/MaxPoint 확인 대기 |
| 6 | 면 | ComponentOccurrence.SurfaceBodies | Inventor.SurfaceBodies | SurfaceBody.Faces 확인 대기 |
| 7 | 면 종류 | Face.SurfaceType | SurfaceTypeEnum | kCylinderSurface 확인, kPlaneSurface 대기 |
| 8 | 원통 축 | Cylinder.BasePoint | Inventor.Point | 단위·좌표계는 실측 대기 |
| 8 | 원통 축 | Cylinder.AxisVector | Inventor.UnitVector | |
| 8 | 원통 축 | Cylinder.Radius | Double | |
| 9 | 면 테두리 | Face.EdgeLoops / Face.Edges | EdgeLoops / Edges | EdgeLoop 멤버 확인 대기 |
| 추가 | 면 | Face.CreatedByFeature / Face.ThreadInfos | PartFeature / ObjectCollection | 6주차 이후 재료 |
| 10 | 엣지 종류 | | | |
| 11 | 원 정보 | | | |
| 12 | 부품 정의 | ComponentOccurrence.Definition | ComponentDefinition | PartComponentDefinition 변환 필요 |
| 13 | 구멍 피처 | PartFeatures.HoleFeatures | HoleFeatures | |
| 14 | 탭 | HoleFeature.Tapped | Boolean | |
| 14 | 탭 | HoleFeature.TapInfo | Object (HoleTapInfo 또는 TaperedThreadInfo) | 변환 필요 |
| 14 | 탭 | HoleTapInfo.ThreadDesignation | String | 나사 호칭 |
| 15 | 클리어런스 | HoleFeature.IsClearanceHole | Boolean | |
| 15 | 클리어런스 | HoleFeature.ClearanceInfo | HoleClearanceInfo | 타입 명확 |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerSize | String | |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerType / FastenerStandard | String | |
| 15 | 클리어런스 | HoleClearanceInfo.FastenerFitType | FastenerFitType | |
| 추가 | 구멍 | HoleFeature.HoleDiameter / HoleType / HoleCenterPoints | Parameter / HoleTypeEnum / ObjectCollection | 8주차 재료 |
| 16 | 패턴·미러 | PartFeatures.RectangularPatternFeatures | RectangularPatternFeatures | |
| 16 | 패턴·미러 | PartFeatures.CircularPatternFeatures | CircularPatternFeatures | |
| 16 | 패턴·미러 | PartFeatures.MirrorFeatures | MirrorFeatures | |
| 추가 | 패턴 | PartFeatures.SketchDrivenPatternFeatures | SketchDrivenPatternFeatures | 계획서에 없음 |
| 추가 | 스케치 기반 패턴 부모 피처 | SketchDrivenPatternFeature.Definition.ParentFeatures | ObjectCollection | 피처 직접 속성 아님 · 실행 미검증 |
| 추가 | 직사각형 패턴 부모 피처 | RectangularPatternFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 검증 · 홀간 거리 확인-1 의 피처 경로에서 확인 |
| 추가 | 원형 패턴 부모 피처 | CircularPatternFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 미검증 · 빌드만 통과 · 원형 패턴 모델 없음 |
| 추가 | 미러 부모 피처 | MirrorFeature.ParentFeatures | ObjectCollection | 피처 직접 속성 · 실행 검증 · 홀간 거리 확인-1 의 피처 경로에서 확인 |
| 추가 | 압출컷 | PartFeatures.ExtrudeFeatures | ExtrudeFeatures | 피처 없는 구멍 확인용 |
| 17 | 활성 뷰 | Application.ActiveView | Inventor.View | W05 · Interop 리플렉션으로 확인 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 카메라 | View.Camera | Inventor.Camera | 읽기 전용 속성 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 화면 맞춤 | Camera.Fit() | Sub | View.Fit(Optional DoUpdate As Boolean)도 있음 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 목표점 지정 | Camera.Target | Inventor.Point (get/set) | cm · 조립품 좌표 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 적용 | Camera.Apply() / Camera.ApplyWithoutTransition() | Sub | Apply는 전환 애니메이션 있음(추정) (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 시점 | Camera.Eye / Camera.UpVector | Point / UnitVector (get/set) | W05 범위 밖 · 이름만 확인 (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 17 | 점 생성 | TransientGeometry.CreatePoint(Optional X, Y, Z As Double) | Inventor.Point | Application.TransientGeometry (W05 · 존재만 확인, 동작은 W10에서 확정) |
| 추가 | 색 표시 세트 | Document.HighlightSets (AssemblyDocument.HighlightSets) | Inventor.HighlightSets (get only) | W06 · Interop 리플렉션으로 확인 |
| 추가 | 색 표시 세트 | HighlightSets.Add() | Inventor.HighlightSet | 인자 없음 (W06) |
| 추가 | 색 표시 세트 | HighlightSet.Color | Inventor.Color (get/set) | W06 |
| 추가 | 색 표시 세트 | HighlightSet.AddItem(Entity As Object) | Sub | 조립품에서는 Face 프록시를 넣음 (W06) |
| 추가 | 색 표시 세트 | HighlightSet.Clear() / HighlightSet.Delete() | Sub | 이름만 확인 · 누적 정리는 W06 범위 밖 |
| 추가 | 색 생성 | Application.TransientObjects | Inventor.TransientObjects (get only) | W06 |
| 추가 | 색 생성 | TransientObjects.CreateColor(Red As Byte, Green As Byte, Blue As Byte, Optional Opacity As Double = 1) | Inventor.Color | W06 |
