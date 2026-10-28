Option Strict On
Option Explicit On

' W07 · 실행 진입 · 진단 출력 (7-2 에서 AxisMatching.vb 에서 이동 · 7-6 에서 축 매칭 · 축 진단 제거)
' 결과에는 숫자와 문자열만 담는다 (COM 객체 금지 · decisions.md "W05 · 결과 구조에는 숫자와 문자열만 담음")
' 매칭 · 판정은 아직 연결하지 않는다 (W08 에서 MatchPair 로 연결)
Public Module MatchingEngine

    ' W07 · 리프 쌍 하나의 최근접 묶음 · Loop 진단 출력용 · 판정 없음
    Private Class LeafPairNearest
        Public intLeafIndexA As Integer
        Public intLeafIndexB As Integer
        Public intClusterNumberA As Integer
        Public intClusterNumberB As Integer
        Public objClusterA As HoleCluster
        Public objClusterB As HoleCluster
        Public dblDistance As Double
    End Class

    ' W07 7-5 · EdgeLoop 수집 실행 · 7-6 에서 축 경로 제거 · 쌍별 거리는 진단 txt 에만 쓴다
    ' strDiagnosticPath · 저장한 Loop 진단 txt 경로
    Public Function RunLoopCollection(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByRef strDiagnosticPath As String) As System.Collections.Generic.List(Of HoleCluster)
        Dim objCounters As New LoopCounters()
        Dim colLeafSummaries As New System.Collections.Generic.List(Of LoopLeafSummary)()

        Dim colClusters As System.Collections.Generic.List(Of HoleCluster) = CollectHoleClusters(invAssemblyDocument, colLeafSummaries, objCounters)

        strDiagnosticPath = WriteLoopDiagnostics(invAssemblyDocument, colLeafSummaries, colClusters, objCounters)

        Return colClusters
    End Function

    ' W07 7-5 · Loop 진단 txt · Feature 진단과 같은 폴더(바탕화면) · 같은 인코딩(BOM 포함 UTF-8) · 같은 이름 규칙(접두만 Loop)
    ' 저장 알림 메시지박스는 띄우지 않는다 · 폼 상태 라벨에 표시
    Private Function WriteLoopDiagnostics(ByVal invAssemblyDocument As Inventor.AssemblyDocument, ByVal colLeafSummaries As System.Collections.Generic.List(Of LoopLeafSummary), ByVal colClusters As System.Collections.Generic.List(Of HoleCluster), ByVal objCounters As LoopCounters) As String
        Dim objText As New System.Text.StringBuilder()
        objText.AppendLine("단위 mm (API cm ×10) · 좌표는 조립품 좌표 · Transformation 은 곱하지 않음")
        objText.AppendLine("문서: " & invAssemblyDocument.FullFileName)
        objText.AppendLine("시각: " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        objText.AppendLine("수집: 평면 face · 안쪽 EdgeLoop · 모든 edge 가 kCircleCurve · 동심 묶음 기준 " & FormatDouble(dblConcentricClusterThreshold) & " mm (같은 리프 안 · 묶음 중심은 첫 원 중심)")
        objText.AppendLine()

        objText.AppendLine("==== 카운터 ====")
        objText.AppendLine("리프 수: " & objCounters.intLeafCount.ToString())
        objText.AppendLine("평면 face 수: " & objCounters.intPlaneFaceCount.ToString())
        objText.AppendLine("안쪽 루프 수: " & objCounters.intInnerLoopCount.ToString())
        objText.AppendLine("채택 루프 수: " & objCounters.intAcceptedLoopCount.ToString())
        objText.AppendLine("탈락 · 원호 포함: " & objCounters.intRejectedArcCount.ToString())
        objText.AppendLine("탈락 · 기타 곡선 포함: " & objCounters.intRejectedOtherCurveCount.ToString())
        objText.AppendLine("탈락 · 읽기 실패: " & objCounters.intReadFailedCount.ToString())
        objText.AppendLine("edge 2개 이상 채택 루프 수: " & objCounters.intMultiEdgeAcceptedLoopCount.ToString())
        objText.AppendLine("묶음 수: " & objCounters.intClusterCount.ToString())
        objText.AppendLine("지름 2개 이상 묶음 수: " & objCounters.intMultiDiameterClusterCount.ToString())
        objText.AppendLine()

        objText.AppendLine("==== 리프별 ====")
        ' W07 7-6 · 탈락 3열 추가 · 앞 6열은 7-5 와 같은 순서 · 리프 오류는 마지막 열 유지
        objText.AppendLine(JoinTab("번호", "occurrence", "부품 정의", "채택 원", "묶음", "지름 2개 이상 묶음", "탈락 원호 포함", "탈락 기타 곡선", "탈락 읽기 실패", "리프 오류"))
        For intIndex As Integer = 0 To colLeafSummaries.Count - 1
            Dim objSummary As LoopLeafSummary = colLeafSummaries(intIndex)
            objText.AppendLine(JoinTab(intIndex.ToString(),
                                       objSummary.strOccurrenceName,
                                       objSummary.strDefinitionName,
                                       objSummary.intAcceptedCircleCount.ToString(),
                                       objSummary.intClusterCount.ToString(),
                                       objSummary.intMultiDiameterClusterCount.ToString(),
                                       objSummary.intRejectedArcCount.ToString(),
                                       objSummary.intRejectedOtherCurveCount.ToString(),
                                       objSummary.intReadFailedCount.ToString(),
                                       objSummary.strLeafError))
        Next
        objText.AppendLine()

        objText.AppendLine("==== 묶음 목록 (" & colClusters.Count.ToString() & "개) ====")
        objText.AppendLine(JoinTab("리프 번호", "occurrence", "묶음 번호", "중심X", "중심Y", "중심Z", "지름 (오름차순)"))
        Dim intPreviousLeafIndex As Integer = -1
        Dim intClusterNumber As Integer = 0
        For Each objCluster As HoleCluster In colClusters
            ' 묶음 번호는 리프 안에서 0부터 · colClusters 는 리프 순서대로 이어 붙어 있다
            If objCluster.intLeafIndex <> intPreviousLeafIndex Then
                intPreviousLeafIndex = objCluster.intLeafIndex
                intClusterNumber = 0
            End If

            Dim colDiameterTexts As New System.Collections.Generic.List(Of String)()
            For Each dblDiameter As Double In objCluster.colDiameters
                colDiameterTexts.Add(FormatDouble(dblDiameter))
            Next

            objText.AppendLine(JoinTab(objCluster.intLeafIndex.ToString(),
                                       objCluster.strOccurrenceName,
                                       intClusterNumber.ToString(),
                                       FormatDouble(objCluster.dblCenterX),
                                       FormatDouble(objCluster.dblCenterY),
                                       FormatDouble(objCluster.dblCenterZ),
                                       String.Join(" / ", colDiameterTexts)))
            intClusterNumber += 1
        Next
        objText.AppendLine()

        WriteLeafPairNearest(objText, colLeafSummaries, colClusters)

        WriteLinkSection(objText, colClusters, objCounters)

        Dim strDocumentName As String = System.IO.Path.GetFileNameWithoutExtension(invAssemblyDocument.DisplayName)
        For Each chrInvalid As Char In System.IO.Path.GetInvalidFileNameChars()
            strDocumentName = strDocumentName.Replace(chrInvalid, "_"c)
        Next
        Dim strFileName As String = "HoleCheck_Loop_" & strDocumentName & "_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        Dim strPath As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory), strFileName)

        ' BOM 포함 UTF-8 · 메모장·PowerShell에서 한글이 깨지지 않게
        System.IO.File.WriteAllText(strPath, objText.ToString(), New System.Text.UTF8Encoding(True))

        Return strPath
    End Function

    ' W07 7-6 · 서로 다른 리프 쌍마다 묶음 중심 사이 3D 거리의 최솟값 · 거리 오름차순 전체 · 필터 없음 · 판정 없음
    ' 묶음이 없는 리프가 낀 쌍은 거리가 없으므로 줄을 쓰지 않고 개수만 적는다
    Private Sub WriteLeafPairNearest(ByVal objText As System.Text.StringBuilder, ByVal colLeafSummaries As System.Collections.Generic.List(Of LoopLeafSummary), ByVal colClusters As System.Collections.Generic.List(Of HoleCluster))
        ' 리프 번호별 묶음 목록 · 목록 안 위치 = 리프 안 묶음 번호 (묶음 목록 섹션과 같은 번호)
        Dim colLeafClusters As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of HoleCluster))()
        For intIndex As Integer = 0 To colLeafSummaries.Count - 1
            colLeafClusters.Add(New System.Collections.Generic.List(Of HoleCluster)())
        Next
        For Each objCluster As HoleCluster In colClusters
            colLeafClusters(objCluster.intLeafIndex).Add(objCluster)
        Next

        Dim colPairs As New System.Collections.Generic.List(Of LeafPairNearest)()
        Dim intSkippedPairCount As Integer = 0
        For intLeafIndexA As Integer = 0 To colLeafClusters.Count - 2
            For intLeafIndexB As Integer = intLeafIndexA + 1 To colLeafClusters.Count - 1
                Dim colClustersA As System.Collections.Generic.List(Of HoleCluster) = colLeafClusters(intLeafIndexA)
                Dim colClustersB As System.Collections.Generic.List(Of HoleCluster) = colLeafClusters(intLeafIndexB)
                If colClustersA.Count = 0 OrElse colClustersB.Count = 0 Then
                    intSkippedPairCount += 1
                    Continue For
                End If

                Dim objPair As New LeafPairNearest()
                objPair.intLeafIndexA = intLeafIndexA
                objPair.intLeafIndexB = intLeafIndexB
                objPair.dblDistance = System.Double.MaxValue
                For intNumberA As Integer = 0 To colClustersA.Count - 1
                    For intNumberB As Integer = 0 To colClustersB.Count - 1
                        Dim objClusterA As HoleCluster = colClustersA(intNumberA)
                        Dim objClusterB As HoleCluster = colClustersB(intNumberB)
                        Dim dblDistance As Double = GetDistance3D(objClusterA.dblCenterX, objClusterA.dblCenterY, objClusterA.dblCenterZ,
                                                                  objClusterB.dblCenterX, objClusterB.dblCenterY, objClusterB.dblCenterZ)
                        If dblDistance < objPair.dblDistance Then
                            objPair.dblDistance = dblDistance
                            objPair.intClusterNumberA = intNumberA
                            objPair.intClusterNumberB = intNumberB
                            objPair.objClusterA = objClusterA
                            objPair.objClusterB = objClusterB
                        End If
                    Next
                Next
                colPairs.Add(objPair)
            Next
        Next

        colPairs.Sort(AddressOf CompareLeafPairNearest)

        objText.AppendLine("==== 리프 쌍별 최근접 묶음 거리 (" & colPairs.Count.ToString() & "쌍 · 거리 오름차순 · 필터 없음 · 판정 없음) ====")
        objText.AppendLine("묶음 없는 리프가 낀 쌍 (줄 없음): " & intSkippedPairCount.ToString())
        objText.AppendLine(JoinTab("리프 A", "occurrence A", "묶음 A", "지름 A (오름차순)", "리프 B", "occurrence B", "묶음 B", "지름 B (오름차순)", "중심 3D 거리"))
        For Each objPair As LeafPairNearest In colPairs
            objText.AppendLine(JoinTab(objPair.intLeafIndexA.ToString(),
                                       objPair.objClusterA.strOccurrenceName,
                                       objPair.intClusterNumberA.ToString(),
                                       FormatDiameterList(objPair.objClusterA),
                                       objPair.intLeafIndexB.ToString(),
                                       objPair.objClusterB.strOccurrenceName,
                                       objPair.intClusterNumberB.ToString(),
                                       FormatDiameterList(objPair.objClusterB),
                                       FormatDouble(objPair.dblDistance)))
        Next
        objText.AppendLine()
    End Sub

    ' W08 8-3 · 연계 1차 결과 · 묶음 단위 · 기존 섹션 뒤에 붙인다
    Private Sub WriteLinkSection(ByVal objText As System.Text.StringBuilder, ByVal colClusters As System.Collections.Generic.List(Of HoleCluster), ByVal objCounters As LoopCounters)
        Dim intLinkedCount As Integer = 0
        For Each objCluster As HoleCluster In colClusters
            If objCluster.intLinkStage = 1 Then
                intLinkedCount += 1
            End If
        Next

        objText.AppendLine("==== 연계 (W08) ====")
        objText.AppendLine("기준: 묶음 첫 루프의 첫 edge → 평면 아닌 인접 면(옆벽) → CreatedByFeature (프록시 직접 · Nothing 이면 NativeObject) → 첫 부모만 따라 올라감 (깊이 상한 " & intMaxFeatureDepth.ToString() & ")")
        objText.AppendLine("1차 성공 묶음 수: " & intLinkedCount.ToString())
        objText.AppendLine("1차 실패 묶음 수: " & (colClusters.Count - intLinkedCount).ToString())
        objText.AppendLine("CreatedByFeature · 프록시 직접 성공: " & objCounters.intLinkProxySuccessCount.ToString())
        objText.AppendLine("CreatedByFeature · NativeObject 경유 성공: " & objCounters.intLinkNativeSuccessCount.ToString())
        objText.AppendLine("CreatedByFeature · 둘 다 Nothing (옆벽 없음 포함): " & objCounters.intLinkFailCount.ToString())
        objText.AppendLine()

        objText.AppendLine("---- CreatedByFeature 결과 타입별 묶음 수 ----")
        Dim colTypeNames As New System.Collections.Generic.List(Of String)(objCounters.colCreatedByFeatureTypeCounts.Keys)
        colTypeNames.Sort(System.StringComparer.Ordinal)
        For Each strTypeName As String In colTypeNames
            objText.AppendLine(JoinTab(strTypeName, objCounters.colCreatedByFeatureTypeCounts(strTypeName).ToString()))
        Next
        objText.AppendLine()

        objText.AppendLine("---- 1차 실패 목록 (" & (colClusters.Count - intLinkedCount).ToString() & "개) ----")
        objText.AppendLine(JoinTab("리프 번호", "occurrence", "묶음 번호", "사유"))
        ' 묶음 번호는 묶음 목록 섹션과 같은 규칙 (리프 안에서 0부터)
        Dim intPreviousLeafIndex As Integer = -1
        Dim intClusterNumber As Integer = 0
        For Each objCluster As HoleCluster In colClusters
            If objCluster.intLeafIndex <> intPreviousLeafIndex Then
                intPreviousLeafIndex = objCluster.intLeafIndex
                intClusterNumber = 0
            End If
            If objCluster.intLinkStage <> 1 Then
                objText.AppendLine(JoinTab(objCluster.intLeafIndex.ToString(),
                                           objCluster.strOccurrenceName,
                                           intClusterNumber.ToString(),
                                           objCluster.strLinkNote))
            End If
            intClusterNumber += 1
        Next
        objText.AppendLine()
    End Sub

    ' 거리 오름차순 · 같으면 리프 A · 리프 B 번호 순
    Private Function CompareLeafPairNearest(ByVal objPairA As LeafPairNearest, ByVal objPairB As LeafPairNearest) As Integer
        Dim intResult As Integer = objPairA.dblDistance.CompareTo(objPairB.dblDistance)
        If intResult <> 0 Then
            Return intResult
        End If
        intResult = objPairA.intLeafIndexA.CompareTo(objPairB.intLeafIndexA)
        If intResult <> 0 Then
            Return intResult
        End If
        Return objPairA.intLeafIndexB.CompareTo(objPairB.intLeafIndexB)
    End Function

    ' 지름 리스트 · 묶음 목록 섹션과 같은 표기 ("a / b")
    Private Function FormatDiameterList(ByVal objCluster As HoleCluster) As String
        Dim colDiameterTexts As New System.Collections.Generic.List(Of String)()
        For Each dblDiameter As Double In objCluster.colDiameters
            colDiameterTexts.Add(FormatDouble(dblDiameter))
        Next
        Return String.Join(" / ", colDiameterTexts)
    End Function

End Module
