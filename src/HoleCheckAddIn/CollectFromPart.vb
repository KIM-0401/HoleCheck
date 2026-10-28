Option Strict On
Option Explicit On

' W07 · EdgeLoop 수집 (7-2 에서 AxisMatching.vb 에서 이동 · 7-6 에서 축 수집 · 피처 판정 · 홀 면 선택 제거)
Friend Module CollectFromPart

    ' W06 피처 추적 · W07 수집에서는 호출하지 않음 · W08 HoleFeature 연계에서 재사용
    ' W06 · CreatedByFeature를 깊이 0으로 보고 ParentFeatures를 따라 올라가는 최대 깊이
    Friend Const intMaxFeatureDepth As Integer = 6

    ' W08 8-3 · 연계 1차에서 CreatedByFeature 를 얻은 경로 · 진단 카운터용
    Private Const intFeatureSourceNone As Integer = 0
    Private Const intFeatureSourceProxy As Integer = 1
    Private Const intFeatureSourceNative As Integer = 2

    ' W06 피처 추적 · W07 수집에서는 호출하지 않음 · W08 HoleFeature 연계에서 재사용
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

    ' W06 피처 추적 · W07 수집에서는 호출하지 않음 · W08 HoleFeature 연계에서 재사용
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

    Private Function GetDefinitionName(ByVal invLeaf As Inventor.ComponentOccurrence) As String
        Try
            Return CType(invLeaf.Definition.Document, Inventor.Document).DisplayName
        Catch objException As System.Exception
            Return "읽기 실패: " & objException.Message
        End Try
    End Function

    ' W07 7-5 · EdgeLoop 수집 · 순회 · 리프 요약 · 카운터 방식은 7-5 당시 CollectCylinderAxes 와 같다 (7-6 에서 축 수집 제거)
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

            ' W07 7-6 · 리프 단위 탈락 수는 전체 카운터의 리프 시작 전 값과의 차이로 기록한다 (수집 로직은 그대로)
            Dim intRejectedArcBefore As Integer = objCounters.intRejectedArcCount
            Dim intRejectedOtherCurveBefore As Integer = objCounters.intRejectedOtherCurveCount
            Dim intReadFailedBefore As Integer = objCounters.intReadFailedCount

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

                                    Dim intFeatureSource As Integer = intFeatureSourceNone
                                    Dim strFeatureTypeName As String = ""
                                    Dim objCircleInfo As LoopCircleInfo = ReadCircleLoop(invLoop, intLeafIndex, objSummary.strOccurrenceName, objCounters, intFeatureSource, strFeatureTypeName)
                                    If objCircleInfo Is Nothing Then
                                        Continue For
                                    End If

                                    objSummary.intAcceptedCircleCount += 1
                                    ' W08 8-3 · 연계 카운터는 묶음 단위 · 새 묶음을 만든 첫 루프의 결과만 센다
                                    If AddToCluster(colLeafClusters, objCircleInfo) Then
                                        CountClusterLink(objCounters, intFeatureSource, strFeatureTypeName)
                                    End If
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

            objSummary.intRejectedArcCount = objCounters.intRejectedArcCount - intRejectedArcBefore
            objSummary.intRejectedOtherCurveCount = objCounters.intRejectedOtherCurveCount - intRejectedOtherCurveBefore
            objSummary.intReadFailedCount = objCounters.intReadFailedCount - intReadFailedBefore

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
    ' W08 8-3 · 채택한 원은 LinkHoleFeatureBySideWall 로 연계 1차를 시도한다 · intFeatureSource · strFeatureTypeName 은 진단 카운터용
    Private Function ReadCircleLoop(ByVal invLoop As Inventor.EdgeLoop, ByVal intLeafIndex As Integer, ByVal strOccurrenceName As String, ByVal objCounters As LoopCounters, ByRef intFeatureSource As Integer, ByRef strFeatureTypeName As String) As LoopCircleInfo
        intFeatureSource = intFeatureSourceNone
        strFeatureTypeName = ""
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

        LinkHoleFeatureBySideWall(invLoop.Edges.Item(1), objCircleInfo, intFeatureSource, strFeatureTypeName)

        Return objCircleInfo
    End Function

    ' W08 8-3 · 연계 1차 · 루프 첫 edge 의 인접 면 중 평면이 아닌 첫 면(옆벽) → CreatedByFeature → FindHoleFeatureBySeed
    ' 옆벽은 SurfaceType 으로만 고른다 (프록시는 같은 면도 다른 래퍼일 수 있어 Is 비교 안 함)
    ' CreatedByFeature 는 프록시에서 직접 읽고(W06 방식) Nothing 이면 FaceProxy.NativeObject 에서 한 번 더 읽는다
    ' 결과는 objCircleInfo 의 연계 필드에만 남긴다 · COM 객체는 이 함수 밖으로 내보내지 않는다 · 예외는 여기서 끝낸다
    Private Sub LinkHoleFeatureBySideWall(ByVal invEdge As Inventor.Edge, ByVal objCircleInfo As LoopCircleInfo, ByRef intFeatureSource As Integer, ByRef strFeatureTypeName As String)
        intFeatureSource = intFeatureSourceNone
        strFeatureTypeName = "(Nothing)"
        objCircleInfo.intLinkStage = 0

        Dim invSideFace As Inventor.Face = Nothing
        Try
            For Each invFace As Inventor.Face In invEdge.Faces
                If invFace.SurfaceType <> Inventor.SurfaceTypeEnum.kPlaneSurface Then
                    invSideFace = invFace
                    Exit For
                End If
            Next
        Catch objException As System.Exception
            objCircleInfo.strLinkNote = "옆벽 읽기 실패: " & objException.Message
            Return
        End Try
        If invSideFace Is Nothing Then
            objCircleInfo.strLinkNote = "옆벽 없음 (인접 면이 모두 평면)"
            Return
        End If

        Dim invFeature As Inventor.PartFeature = ReadCreatedByFeature(invSideFace)
        If invFeature IsNot Nothing Then
            intFeatureSource = intFeatureSourceProxy
        ElseIf TypeOf invSideFace Is Inventor.FaceProxy Then
            Dim invNativeFace As Inventor.Face = Nothing
            Try
                invNativeFace = CType(invSideFace, Inventor.FaceProxy).NativeObject
            Catch
                invNativeFace = Nothing
            End Try
            If invNativeFace IsNot Nothing Then
                invFeature = ReadCreatedByFeature(invNativeFace)
                If invFeature IsNot Nothing Then
                    intFeatureSource = intFeatureSourceNative
                End If
            End If
        End If

        If invFeature Is Nothing Then
            If TypeOf invSideFace Is Inventor.FaceProxy Then
                objCircleInfo.strLinkNote = "CreatedByFeature 없음 (프록시 · NativeObject 모두 Nothing)"
            Else
                objCircleInfo.strLinkNote = "CreatedByFeature 없음 (프록시 아닌 면 · Nothing)"
            End If
            Return
        End If

        Try
            strFeatureTypeName = invFeature.Type.ToString()
        Catch
            strFeatureTypeName = "(Type 읽기 실패)"
        End Try

        Dim strReason As String = ""
        Dim invHoleFeature As Inventor.HoleFeature = FindHoleFeatureBySeed(invFeature, strReason)
        If invHoleFeature Is Nothing Then
            objCircleInfo.strLinkNote = strReason
            Return
        End If

        Try
            objCircleInfo.strHoleFeatureName = invHoleFeature.Name
        Catch objException As System.Exception
            objCircleInfo.strLinkNote = "HoleFeature 이름 읽기 실패: " & objException.Message
            Return
        End Try
        objCircleInfo.intLinkStage = 1
    End Sub

    ' W08 8-3 · CreatedByFeature 읽기 · 예외도 Nothing 으로 본다
    Private Function ReadCreatedByFeature(ByVal invFace As Inventor.Face) As Inventor.PartFeature
        Try
            Return invFace.CreatedByFeature
        Catch
            Return Nothing
        End Try
    End Function

    ' W08 8-3 · 시드 역추적 · FindHoleAncestor 축약판 · 부모 피처 첫 항목만 따라 올라감 · 깊이 상한 intMaxFeatureDepth
    ' 깊이는 FindHoleAncestor 와 같이 CreatedByFeature 를 0 으로 센다 · 못 찾으면 Nothing · strReason 에 사유와 마지막 피처 (Type:이름)
    Private Function FindHoleFeatureBySeed(ByVal invFeature As Inventor.PartFeature, ByRef strReason As String) As Inventor.HoleFeature
        strReason = ""
        Dim invCurrentFeature As Inventor.PartFeature = invFeature

        For intDepth As Integer = 0 To intMaxFeatureDepth
            Dim strCurrentFeature As String = ""
            Try
                If TypeOf invCurrentFeature Is Inventor.HoleFeature Then
                    Return CType(invCurrentFeature, Inventor.HoleFeature)
                End If
                strCurrentFeature = invCurrentFeature.Type.ToString() & ":" & invCurrentFeature.Name
            Catch objException As System.Exception
                strReason = "피처 읽기 실패 (깊이 " & intDepth.ToString() & "): " & objException.Message
                Return Nothing
            End Try

            If intDepth >= intMaxFeatureDepth Then
                strReason = "깊이 상한 " & intMaxFeatureDepth.ToString() & " (" & strCurrentFeature & ")"
                Return Nothing
            End If

            Dim objFirstParent As Object = Nothing
            Try
                Dim invParentFeatures As Inventor.ObjectCollection = GetParentFeatures(invCurrentFeature)
                If invParentFeatures Is Nothing Then
                    strReason = "구멍 피처 아님 · 부모 없음 (" & strCurrentFeature & ")"
                    Return Nothing
                End If
                ' 첫 항목만 · 인덱서 시작 번호가 미확인이라 For Each 로 첫 항목을 얻는다 (api_paths.md ObjectCollection 행)
                For Each objParent As Object In invParentFeatures
                    objFirstParent = objParent
                    Exit For
                Next
            Catch objException As System.Exception
                strReason = "부모 피처 읽기 실패 (" & strCurrentFeature & "): " & objException.Message
                Return Nothing
            End Try

            If objFirstParent Is Nothing Then
                strReason = "부모 피처 목록 비어 있음 (" & strCurrentFeature & ")"
                Return Nothing
            End If
            If Not TypeOf objFirstParent Is Inventor.PartFeature Then
                strReason = "첫 부모가 PartFeature 아님 (" & strCurrentFeature & ")"
                Return Nothing
            End If
            invCurrentFeature = CType(objFirstParent, Inventor.PartFeature)
        Next

        strReason = "깊이 상한 " & intMaxFeatureDepth.ToString()
        Return Nothing
    End Function

    ' W08 8-3 · 새 묶음 하나의 연계 1차 결과를 카운터에 더한다 · 묶음 단위 (묶음 첫 루프 기준)
    ' intLinkFailCount = CreatedByFeature 가 프록시 · NativeObject 모두 Nothing 인 묶음 · HoleFeature 를 못 찾은 묶음 수는 진단에서 intLinkStage 로 센다
    Private Sub CountClusterLink(ByVal objCounters As LoopCounters, ByVal intFeatureSource As Integer, ByVal strFeatureTypeName As String)
        Select Case intFeatureSource
            Case intFeatureSourceProxy
                objCounters.intLinkProxySuccessCount += 1
            Case intFeatureSourceNative
                objCounters.intLinkNativeSuccessCount += 1
            Case Else
                objCounters.intLinkFailCount += 1
        End Select

        Dim intTypeCount As Integer = 0
        objCounters.colCreatedByFeatureTypeCounts.TryGetValue(strFeatureTypeName, intTypeCount)
        objCounters.colCreatedByFeatureTypeCounts(strFeatureTypeName) = intTypeCount + 1
    End Sub

    ' W07 7-5 · 중심 3D 거리 ≤ dblConcentricClusterThreshold 인 첫 묶음에 지름을 더하고, 없으면 이 원 중심으로 새 묶음을 만든다
    ' W08 8-3 · 새 묶음을 만들면 True · 첫 루프의 연계 필드 5개를 복사 · 기존 묶음에 지름만 더할 때는 연계 필드를 바꾸지 않는다
    Private Function AddToCluster(ByVal colLeafClusters As System.Collections.Generic.List(Of HoleCluster), ByVal objCircleInfo As LoopCircleInfo) As Boolean
        For Each objCluster As HoleCluster In colLeafClusters
            Dim dblDistance As Double = GetDistance3D(objCluster.dblCenterX, objCluster.dblCenterY, objCluster.dblCenterZ,
                                                      objCircleInfo.dblCenterX, objCircleInfo.dblCenterY, objCircleInfo.dblCenterZ)
            If dblDistance <= dblConcentricClusterThreshold Then
                objCluster.AddDiameter(objCircleInfo.dblDiameter)
                Return False
            End If
        Next

        Dim objNewCluster As New HoleCluster()
        objNewCluster.intLeafIndex = objCircleInfo.intLeafIndex
        objNewCluster.strOccurrenceName = objCircleInfo.strOccurrenceName
        objNewCluster.dblCenterX = objCircleInfo.dblCenterX
        objNewCluster.dblCenterY = objCircleInfo.dblCenterY
        objNewCluster.dblCenterZ = objCircleInfo.dblCenterZ
        objNewCluster.AddDiameter(objCircleInfo.dblDiameter)
        objNewCluster.strHoleFeatureName = objCircleInfo.strHoleFeatureName
        objNewCluster.bolTapped = objCircleInfo.bolTapped
        objNewCluster.strThreadDesignation = objCircleInfo.strThreadDesignation
        objNewCluster.intLinkStage = objCircleInfo.intLinkStage
        objNewCluster.strLinkNote = objCircleInfo.strLinkNote
        colLeafClusters.Add(objNewCluster)
        Return True
    End Function

End Module
