Option Strict On
Option Explicit On

' W07 7-2 · 수집 · 피처 추적 · 홀 면 선택 (AxisMatching.vb에서 이동)
Friend Module CollectFromPart

    ' W06 · CreatedByFeature를 깊이 0으로 보고 ParentFeatures를 따라 올라가는 최대 깊이
    Friend Const intMaxFeatureDepth As Integer = 6

    ' W06 · 원통면의 피처 판정 상태
    Friend Const strFeatureStateHole As String = "홀"
    Private Const strFeatureStateNotHole As String = "홀 아님"
    Private Const strFeatureStateReadFailed As String = "피처 읽기 실패"
    Friend Const strFeatureStateOuterRing As String = "바깥 링 제외"

    ' 1단계 · 수집 · AllLeafOccurrences 순회 · 리프 Transformation과 프록시 Geometry에 변환을 곱하지 않는다 (W04 결정)
    Friend Function CollectCylinderAxes(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal colLeafSummaries As System.Collections.Generic.List(Of LeafSummary), ByVal objCounters As MatchCounters) As System.Collections.Generic.List(Of CylinderAxis)
        Dim colAxes As New System.Collections.Generic.List(Of CylinderAxis)()
        Dim intLeafIndex As Integer = 0

        For Each invLeaf As Inventor.ComponentOccurrence In invAssemblyDocument.ComponentDefinition.Occurrences.AllLeafOccurrences()
            Dim objSummary As New LeafSummary()
            objSummary.strOccurrenceName = invLeaf.Name
            objSummary.strDefinitionName = GetDefinitionName(invLeaf)
            colLeafSummaries.Add(objSummary)

            Try
                For Each invBody As Inventor.SurfaceBody In invLeaf.SurfaceBodies
                    For Each invFace As Inventor.Face In invBody.Faces
                        Try
                            If invFace.SurfaceType <> Inventor.SurfaceTypeEnum.kCylinderSurface Then
                                Continue For
                            End If

                            Dim invCylinder As Inventor.Cylinder = CType(invFace.Geometry, Inventor.Cylinder)
                            Dim objAxis As New CylinderAxis()
                            objAxis.intLeafIndex = intLeafIndex
                            objAxis.strOccurrenceName = objSummary.strOccurrenceName
                            objAxis.dblPointX = invCylinder.BasePoint.X * dblCentimeterToMillimeter
                            objAxis.dblPointY = invCylinder.BasePoint.Y * dblCentimeterToMillimeter
                            objAxis.dblPointZ = invCylinder.BasePoint.Z * dblCentimeterToMillimeter
                            objAxis.dblDirectionX = invCylinder.AxisVector.X
                            objAxis.dblDirectionY = invCylinder.AxisVector.Y
                            objAxis.dblDirectionZ = invCylinder.AxisVector.Z
                            objAxis.dblRadius = invCylinder.Radius * dblCentimeterToMillimeter

                            objAxis.invFace = invFace
                            ClassifyFaceFeature(invFace, objAxis)

                            colAxes.Add(objAxis)
                            objSummary.intCylinderCount += 1
                        Catch
                            objSummary.intFailedFaceCount += 1
                        End Try
                    Next
                Next
            Catch objException As System.Exception
                objSummary.strLeafError = objException.Message
            End Try

            objCounters.intCylinderCount += objSummary.intCylinderCount
            objCounters.intFailedFaceCount += objSummary.intFailedFaceCount
            intLeafIndex += 1
        Next

        Return colAxes
    End Function

    ' W06 · 면을 만든 피처에서 HoleFeature 조상을 찾아 objAxis에 상태·이름·경로를 적는다
    ' 읽기 실패는 면 수집 실패와 따로 센다 (원통면 자체는 수집됨)
    Private Sub ClassifyFaceFeature(ByVal invFace As Inventor.Face, ByVal objAxis As CylinderAxis)
        Try
            Dim colPath As New System.Collections.Generic.List(Of String)()
            Dim colVisited As New System.Collections.Generic.HashSet(Of String)()
            Dim invHoleFeature As Inventor.HoleFeature = FindHoleAncestor(invFace.CreatedByFeature, 0, colVisited, colPath)

            objAxis.strFeaturePath = String.Join(" > ", colPath)
            If invHoleFeature Is Nothing Then
                objAxis.strFeatureState = strFeatureStateNotHole
            Else
                objAxis.strFeatureState = strFeatureStateHole
                objAxis.strHoleFeatureName = invHoleFeature.Name
            End If
        Catch objException As System.Exception
            objAxis.strFeatureState = strFeatureStateReadFailed
            objAxis.strFeaturePath = objException.Message
        End Try
    End Sub

    ' W06 · 재귀 · 깊이 상한 intMaxFeatureDepth · 방문한 피처 이름을 기록해 순환을 막는다
    ' 피처 이름은 한 부품 문서 안에서 고유하므로 방문 키로 쓴다
    Private Function FindHoleAncestor(ByVal invFeature As Inventor.PartFeature, ByVal intDepth As Integer, ByVal colVisited As System.Collections.Generic.HashSet(Of String), ByVal colPath As System.Collections.Generic.List(Of String)) As Inventor.HoleFeature
        If invFeature Is Nothing Then
            colPath.Add("(피처 없음)")
            Return Nothing
        End If

        Dim strFeatureName As String = invFeature.Name
        If Not colVisited.Add(strFeatureName) Then
            colPath.Add("(순환 " & strFeatureName & ")")
            Return Nothing
        End If
        colPath.Add(invFeature.Type.ToString() & ":" & strFeatureName)

        If TypeOf invFeature Is Inventor.HoleFeature Then
            Return CType(invFeature, Inventor.HoleFeature)
        End If

        If intDepth >= intMaxFeatureDepth Then
            colPath.Add("(깊이 상한 " & intMaxFeatureDepth.ToString() & ")")
            Return Nothing
        End If

        Dim invParentFeatures As Inventor.ObjectCollection = GetParentFeatures(invFeature)
        If invParentFeatures Is Nothing Then
            Return Nothing
        End If

        For Each objParent As Object In invParentFeatures
            If Not TypeOf objParent Is Inventor.PartFeature Then
                Continue For
            End If
            Dim invHoleFeature As Inventor.HoleFeature = FindHoleAncestor(CType(objParent, Inventor.PartFeature), intDepth + 1, colVisited, colPath)
            If invHoleFeature IsNot Nothing Then
                Return invHoleFeature
            End If
        Next

        Return Nothing
    End Function

    ' W06 · ParentFeatures는 PartFeature 공통 멤버가 아니라 패턴·미러 피처에만 있다 · 그 외 피처는 Nothing
    '       SketchDrivenPatternFeature는 피처에 ParentFeatures가 없고 Definition(SketchDrivenPatternDefinition)에 있다
    Private Function GetParentFeatures(ByVal invFeature As Inventor.PartFeature) As Inventor.ObjectCollection
        If TypeOf invFeature Is Inventor.RectangularPatternFeature Then
            Return CType(invFeature, Inventor.RectangularPatternFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.CircularPatternFeature Then
            Return CType(invFeature, Inventor.CircularPatternFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.MirrorFeature Then
            Return CType(invFeature, Inventor.MirrorFeature).ParentFeatures
        End If
        If TypeOf invFeature Is Inventor.SketchDrivenPatternFeature Then
            Return CType(invFeature, Inventor.SketchDrivenPatternFeature).Definition.ParentFeatures
        End If
        Return Nothing
    End Function

    ' W06 · 홀로 확정된 면만 남기고, 같은 occurrence · 같은 HoleFeature · 축 일치(평행 · 축간 거리 dblCoaxialDistanceThreshold 이하) 묶음에서
    '       최소 반지름 하나만 남긴다 · 나머지는 카운터보어 바깥 링으로 보고 짝 계산에서 뺀다
    Friend Function SelectHoleAxes(ByVal colAxes As System.Collections.Generic.List(Of CylinderAxis), ByVal objCounters As MatchCounters) As System.Collections.Generic.List(Of CylinderAxis)
        Dim colSelected As New System.Collections.Generic.List(Of CylinderAxis)()
        Dim dblParallelCosine As Double = System.Math.Cos(dblParallelToleranceDegrees * System.Math.PI / 180.0)

        For Each objAxis As CylinderAxis In colAxes
            If objAxis.strFeatureState <> strFeatureStateHole Then
                Continue For
            End If

            Dim intGroupIndex As Integer = -1
            For intIndex As Integer = 0 To colSelected.Count - 1
                Dim objKept As CylinderAxis = colSelected(intIndex)
                If objKept.intLeafIndex <> objAxis.intLeafIndex OrElse objKept.strHoleFeatureName <> objAxis.strHoleFeatureName Then
                    Continue For
                End If

                Dim dblDot As Double = objKept.dblDirectionX * objAxis.dblDirectionX +
                                       objKept.dblDirectionY * objAxis.dblDirectionY +
                                       objKept.dblDirectionZ * objAxis.dblDirectionZ
                If System.Math.Abs(dblDot) < dblParallelCosine Then
                    Continue For
                End If

                If GetPointToLineDistance(objAxis, objKept) <= dblCoaxialDistanceThreshold Then
                    intGroupIndex = intIndex
                    Exit For
                End If
            Next

            If intGroupIndex < 0 Then
                colSelected.Add(objAxis)
            ElseIf objAxis.dblRadius < colSelected(intGroupIndex).dblRadius Then
                colSelected(intGroupIndex).strFeatureState = strFeatureStateOuterRing
                colSelected(intGroupIndex) = objAxis
            Else
                objAxis.strFeatureState = strFeatureStateOuterRing
            End If
        Next

        For Each objAxis As CylinderAxis In colAxes
            Select Case objAxis.strFeatureState
                Case strFeatureStateHole
                    objCounters.intHoleCount += 1
                Case strFeatureStateOuterRing
                    objCounters.intOuterRingCount += 1
                Case strFeatureStateNotHole
                    objCounters.intNotHoleCount += 1
                Case strFeatureStateReadFailed
                    objCounters.intFeatureReadFailedCount += 1
            End Select
        Next

        Return colSelected
    End Function

    Private Function GetDefinitionName(ByVal invLeaf As Inventor.ComponentOccurrence) As String
        Try
            Return CType(invLeaf.Definition.Document, Inventor.Document).DisplayName
        Catch objException As System.Exception
            Return "읽기 실패: " & objException.Message
        End Try
    End Function

    ' W07 7-5 · EdgeLoop 수집 · 축 경로(CollectCylinderAxes) 옆에 둔다 · 순회 · 리프 요약 · 카운터 방식은 CollectCylinderAxes 와 같다
    ' 평면 face → 안쪽 EdgeLoop → 모든 edge 가 kCircleCurve 인 루프만 채택 → 같은 리프 안 동심 묶음
    ' 프록시 Geometry 는 이미 조립품 좌표 · Transformation 은 곱하지 않고 cm→mm 만 한다 (W04 결정 · W07 decisions)
    Friend Function CollectHoleClusters(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal colLeafSummaries As System.Collections.Generic.List(Of LoopLeafSummary), ByVal objCounters As LoopCounters) As System.Collections.Generic.List(Of HoleCluster)
        Dim colClusters As New System.Collections.Generic.List(Of HoleCluster)()
        Dim intLeafIndex As Integer = 0

        For Each invLeaf As Inventor.ComponentOccurrence In invAssemblyDocument.ComponentDefinition.Occurrences.AllLeafOccurrences()
            Dim objSummary As New LoopLeafSummary()
            objSummary.strOccurrenceName = invLeaf.Name
            objSummary.strDefinitionName = GetDefinitionName(invLeaf)
            colLeafSummaries.Add(objSummary)

            ' 동심 묶음은 같은 리프 안에서만 찾는다
            Dim colLeafClusters As New System.Collections.Generic.List(Of HoleCluster)()

            Try
                For Each invBody As Inventor.SurfaceBody In invLeaf.SurfaceBodies
                    For Each invFace As Inventor.Face In invBody.Faces
                        Try
                            If invFace.SurfaceType <> Inventor.SurfaceTypeEnum.kPlaneSurface Then
                                Continue For
                            End If
                            objCounters.intPlaneFaceCount += 1

                            For Each invLoop As Inventor.EdgeLoop In invFace.EdgeLoops
                                Try
                                    If invLoop.IsOuterEdgeLoop Then
                                        Continue For
                                    End If
                                    objCounters.intInnerLoopCount += 1

                                    Dim objCircleInfo As LoopCircleInfo = ReadCircleLoop(invLoop, intLeafIndex, objSummary.strOccurrenceName, objCounters)
                                    If objCircleInfo Is Nothing Then
                                        Continue For
                                    End If

                                    objSummary.intAcceptedCircleCount += 1
                                    AddToCluster(colLeafClusters, objCircleInfo)
                                Catch
                                    objCounters.intReadFailedCount += 1
                                End Try
                            Next
                        Catch
                            objCounters.intReadFailedCount += 1
                        End Try
                    Next
                Next
            Catch objException As System.Exception
                objSummary.strLeafError = objException.Message
            End Try

            objSummary.intClusterCount = colLeafClusters.Count
            For Each objCluster As HoleCluster In colLeafClusters
                If objCluster.colDiameters.Count >= 2 Then
                    objSummary.intMultiDiameterClusterCount += 1
                End If
            Next
            colClusters.AddRange(colLeafClusters)

            objCounters.intLeafCount += 1
            objCounters.intClusterCount += objSummary.intClusterCount
            objCounters.intMultiDiameterClusterCount += objSummary.intMultiDiameterClusterCount
            intLeafIndex += 1
        Next

        Return colClusters
    End Function

    ' W07 7-5 · 안쪽 루프 하나 판정 · 모든 edge 가 kCircleCurve 면 첫 edge 의 원으로 LoopCircleInfo 를 만든다
    ' 탈락은 사유별로 센다 · 원호가 하나라도 있으면 "원호 포함", 그 밖의 곡선이면 "기타 곡선 포함" · 예외는 호출자가 "읽기 실패"로 센다
    Private Function ReadCircleLoop(ByVal invLoop As Inventor.EdgeLoop, ByVal intLeafIndex As Integer, ByVal strOccurrenceName As String, ByVal objCounters As LoopCounters) As LoopCircleInfo
        Dim bolHasArc As Boolean = False
        Dim bolHasOtherCurve As Boolean = False
        For Each invEdge As Inventor.Edge In invLoop.Edges
            If invEdge.GeometryType = Inventor.CurveTypeEnum.kCircleCurve Then
                Continue For
            End If
            If invEdge.GeometryType = Inventor.CurveTypeEnum.kCircularArcCurve Then
                bolHasArc = True
            Else
                bolHasOtherCurve = True
            End If
        Next

        If bolHasArc Then
            objCounters.intRejectedArcCount += 1
            Return Nothing
        End If
        If bolHasOtherCurve Then
            objCounters.intRejectedOtherCurveCount += 1
            Return Nothing
        End If

        Dim intEdgeCount As Integer = invLoop.Edges.Count
        Dim invCircle As Inventor.Circle = CType(invLoop.Edges.Item(1).Geometry, Inventor.Circle)

        Dim objCircleInfo As New LoopCircleInfo()
        objCircleInfo.intLeafIndex = intLeafIndex
        objCircleInfo.strOccurrenceName = strOccurrenceName
        objCircleInfo.dblCenterX = ConvertCentimeterToMillimeter(invCircle.Center.X)
        objCircleInfo.dblCenterY = ConvertCentimeterToMillimeter(invCircle.Center.Y)
        objCircleInfo.dblCenterZ = ConvertCentimeterToMillimeter(invCircle.Center.Z)
        objCircleInfo.dblDiameter = ConvertCentimeterToMillimeter(invCircle.Radius * 2.0)
        objCircleInfo.intLoopEdgeCount = intEdgeCount

        objCounters.intAcceptedLoopCount += 1
        If intEdgeCount >= 2 Then
            objCounters.intMultiEdgeAcceptedLoopCount += 1
        End If

        Return objCircleInfo
    End Function

    ' W07 7-5 · 중심 3D 거리 ≤ dblConcentricClusterThreshold 인 첫 묶음에 지름을 더하고, 없으면 이 원 중심으로 새 묶음을 만든다
    Private Sub AddToCluster(ByVal colLeafClusters As System.Collections.Generic.List(Of HoleCluster), ByVal objCircleInfo As LoopCircleInfo)
        For Each objCluster As HoleCluster In colLeafClusters
            Dim dblDistance As Double = GetDistance3D(objCluster.dblCenterX, objCluster.dblCenterY, objCluster.dblCenterZ,
                                                      objCircleInfo.dblCenterX, objCircleInfo.dblCenterY, objCircleInfo.dblCenterZ)
            If dblDistance <= dblConcentricClusterThreshold Then
                objCluster.AddDiameter(objCircleInfo.dblDiameter)
                Return
            End If
        Next

        Dim objNewCluster As New HoleCluster()
        objNewCluster.intLeafIndex = objCircleInfo.intLeafIndex
        objNewCluster.strOccurrenceName = objCircleInfo.strOccurrenceName
        objNewCluster.dblCenterX = objCircleInfo.dblCenterX
        objNewCluster.dblCenterY = objCircleInfo.dblCenterY
        objNewCluster.dblCenterZ = objCircleInfo.dblCenterZ
        objNewCluster.AddDiameter(objCircleInfo.dblDiameter)
        colLeafClusters.Add(objNewCluster)
    End Sub

End Module
