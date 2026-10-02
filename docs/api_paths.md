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
| 추가 | 압출컷 | PartFeatures.ExtrudeFeatures | ExtrudeFeatures | 피처 없는 구멍 확인용 |