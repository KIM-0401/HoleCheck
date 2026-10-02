Option Strict On
Option Explicit On

' 디자이너 파일 없이 코드로 만든 폼 · W05 원통면 축 기반 1차 매칭 결과 표시
' 입력 2개(후보 범위 · 정렬 공차) → RunAxisMatching → 4열 그리드 · 더블클릭 시 Camera.Fit 후 목표점으로 이동
Public Class HoleCheckForm
    Inherits System.Windows.Forms.Form

    ' 내부 mm → API cm (W04 결정의 역변환)
    Private Const dblCentimeterToMillimeter As Double = 10.0

    Private Const dblDefaultCandidateRange As Double = 2.0
    Private Const dblDefaultAlignmentTolerance As Double = 0.1

    Private invApplication As Inventor.Application

    ' 그리드 줄 번호(0부터) = 이 목록의 인덱스 · 정렬을 막아 순서를 유지한다
    Private colResults As New System.Collections.Generic.List(Of AxisMatchResult)()

    Private objCandidateRangeInput As System.Windows.Forms.NumericUpDown
    Private objAlignmentToleranceInput As System.Windows.Forms.NumericUpDown
    Private objRunButton As System.Windows.Forms.Button
    Private objStatusLabel As System.Windows.Forms.Label
    Private objResultGrid As System.Windows.Forms.DataGridView
    Private objSelectionLabel As System.Windows.Forms.Label

    Public Sub New(ByVal invInventorApplication As Inventor.Application)
        MyBase.New()

        invApplication = invInventorApplication

        Me.Text = "HoleCheck"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Size = New System.Drawing.Size(640, 480)

        BuildControls()

        AddHandler objRunButton.Click, AddressOf OnRunClick
        AddHandler objResultGrid.SelectionChanged, AddressOf OnGridSelectionChanged
        AddHandler objResultGrid.CellDoubleClick, AddressOf OnGridCellDoubleClick
    End Sub

    Private Sub BuildControls()
        ' 입력 줄
        Dim objInputPanel As New System.Windows.Forms.FlowLayoutPanel()
        objInputPanel.Dock = System.Windows.Forms.DockStyle.Top
        objInputPanel.Height = 36
        objInputPanel.Padding = New System.Windows.Forms.Padding(4)

        objCandidateRangeInput = CreateNumberInput(dblDefaultCandidateRange)
        objAlignmentToleranceInput = CreateNumberInput(dblDefaultAlignmentTolerance)

        objRunButton = New System.Windows.Forms.Button()
        objRunButton.Text = "실행"
        objRunButton.AutoSize = True

        objInputPanel.Controls.Add(CreateInputLabel("후보 범위(mm)"))
        objInputPanel.Controls.Add(objCandidateRangeInput)
        objInputPanel.Controls.Add(CreateInputLabel("정렬 공차(mm)"))
        objInputPanel.Controls.Add(objAlignmentToleranceInput)
        objInputPanel.Controls.Add(objRunButton)

        ' 상태 한 줄
        objStatusLabel = New System.Windows.Forms.Label()
        objStatusLabel.Dock = System.Windows.Forms.DockStyle.Top
        objStatusLabel.Height = 24
        objStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        objStatusLabel.Text = "실행 전"

        ' 그리드 4열
        objResultGrid = New System.Windows.Forms.DataGridView()
        objResultGrid.Dock = System.Windows.Forms.DockStyle.Fill
        objResultGrid.ReadOnly = True
        objResultGrid.AllowUserToAddRows = False
        objResultGrid.AllowUserToDeleteRows = False
        objResultGrid.RowHeadersVisible = False
        objResultGrid.MultiSelect = False
        objResultGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        objResultGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        objResultGrid.Columns.Add("colNumber", "번호")
        objResultGrid.Columns.Add("colErrorKind", "오류 종류")
        objResultGrid.Columns.Add("colValue1", "값1 · 축간 거리(mm)")
        objResultGrid.Columns.Add("colValue2", "값2 · A지름 / B지름(mm)")
        For Each objColumn As System.Windows.Forms.DataGridViewColumn In objResultGrid.Columns
            objColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Next

        ' 선택 줄의 occurrence 쌍
        objSelectionLabel = New System.Windows.Forms.Label()
        objSelectionLabel.Dock = System.Windows.Forms.DockStyle.Bottom
        objSelectionLabel.Height = 24
        objSelectionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        objSelectionLabel.Text = ""

        ' Dock은 나중에 추가한 컨트롤부터 자리를 잡는다 · Fill을 먼저 넣는다
        Me.Controls.Add(objResultGrid)
        Me.Controls.Add(objSelectionLabel)
        Me.Controls.Add(objStatusLabel)
        Me.Controls.Add(objInputPanel)
    End Sub

    Private Function CreateNumberInput(ByVal dblInitialValue As Double) As System.Windows.Forms.NumericUpDown
        Dim objInput As New System.Windows.Forms.NumericUpDown()
        objInput.DecimalPlaces = 3
        objInput.Increment = 0.1D
        objInput.Minimum = 0D
        objInput.Maximum = 1000D
        objInput.Value = CDec(dblInitialValue)
        objInput.Width = 80
        Return objInput
    End Function

    Private Function CreateInputLabel(ByVal strText As String) As System.Windows.Forms.Label
        Dim objLabel As New System.Windows.Forms.Label()
        objLabel.Text = strText
        objLabel.AutoSize = True
        objLabel.Padding = New System.Windows.Forms.Padding(0, 6, 0, 0)
        Return objLabel
    End Function

    Private Sub OnRunClick(ByVal objSender As Object, ByVal objArgs As System.EventArgs)
        Try
            Dim dblCandidateRange As Double = CDbl(objCandidateRangeInput.Value)
            Dim dblAlignmentTolerance As Double = CDbl(objAlignmentToleranceInput.Value)

            If dblCandidateRange <= 0 OrElse dblAlignmentTolerance <= 0 Then
                System.Windows.Forms.MessageBox.Show("후보 범위와 정렬 공차는 0보다 커야 합니다.", "HoleCheck")
                Return
            End If
            If dblAlignmentTolerance >= dblCandidateRange Then
                System.Windows.Forms.MessageBox.Show("정렬 공차는 후보 범위보다 작아야 합니다.", "HoleCheck")
                Return
            End If

            Dim invAssemblyDocument As Inventor.AssemblyDocument = TryCast(invApplication.ActiveDocument, Inventor.AssemblyDocument)
            If invAssemblyDocument Is Nothing Then
                System.Windows.Forms.MessageBox.Show("조립품(.iam) 문서를 활성화하고 실행하세요.", "HoleCheck")
                Return
            End If

            colResults = AxisMatching.RunAxisMatching(invAssemblyDocument, dblCandidateRange, dblAlignmentTolerance)
            FillGrid()
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · 실행 오류")
        End Try
    End Sub

    ' OK 줄도 전부 표시
    Private Sub FillGrid()
        objResultGrid.Rows.Clear()
        objSelectionLabel.Text = ""

        Dim intOkCount As Integer = 0
        Dim intPositionMismatchCount As Integer = 0

        For intIndex As Integer = 0 To colResults.Count - 1
            Dim objResult As AxisMatchResult = colResults(intIndex)
            objResultGrid.Rows.Add((intIndex + 1).ToString(),
                                objResult.strErrorKind,
                                FormatMillimeter(objResult.dblAxisDistance),
                                FormatMillimeter(objResult.dblDiameterA) & " / " & FormatMillimeter(objResult.dblDiameterB))

            If objResult.strErrorKind = AxisMatchResult.strErrorKindOk Then
                intOkCount += 1
            ElseIf objResult.strErrorKind = AxisMatchResult.strErrorKindPositionMismatch Then
                intPositionMismatchCount += 1
            End If
        Next

        objStatusLabel.Text = "전체 " & colResults.Count.ToString() & "줄 · OK " & intOkCount.ToString() & " · 위치 불일치 " & intPositionMismatchCount.ToString()
    End Sub

    Private Sub OnGridSelectionChanged(ByVal objSender As Object, ByVal objArgs As System.EventArgs)
        Dim objResult As AxisMatchResult = GetResultAt(If(objResultGrid.CurrentRow Is Nothing, -1, objResultGrid.CurrentRow.Index))
        If objResult Is Nothing Then
            objSelectionLabel.Text = ""
            Return
        End If
        objSelectionLabel.Text = objResult.strOccurrenceNameA & " ↔ " & objResult.strOccurrenceNameB
    End Sub

    ' 더블클릭 · 계획서 5주차대로 Camera.Fit 후 목표점 지정 · Eye는 계산하지 않는다
    Private Sub OnGridCellDoubleClick(ByVal objSender As Object, ByVal objArgs As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            Dim objResult As AxisMatchResult = GetResultAt(objArgs.RowIndex)
            If objResult Is Nothing Then
                Return
            End If

            Dim invView As Inventor.View = invApplication.ActiveView
            If invView Is Nothing Then
                System.Windows.Forms.MessageBox.Show("활성 뷰가 없습니다.", "HoleCheck")
                Return
            End If

            ' 목표점 mm → cm
            Dim invTarget As Inventor.Point = invApplication.TransientGeometry.CreatePoint(objResult.dblTargetX / dblCentimeterToMillimeter,
                                                                                            objResult.dblTargetY / dblCentimeterToMillimeter,
                                                                                            objResult.dblTargetZ / dblCentimeterToMillimeter)

            Dim invCamera As Inventor.Camera = invView.Camera
            invCamera.Fit()
            invCamera.Target = invTarget
            invCamera.Apply()
        Catch objException As System.Exception
            System.Windows.Forms.MessageBox.Show(objException.Message & System.Environment.NewLine & System.Environment.NewLine & objException.StackTrace, "HoleCheck · 이동 오류")
        End Try
    End Sub

    Private Function GetResultAt(ByVal intRowIndex As Integer) As AxisMatchResult
        If intRowIndex < 0 OrElse intRowIndex >= colResults.Count Then
            Return Nothing
        End If
        Return colResults(intRowIndex)
    End Function

    Private Function FormatMillimeter(ByVal dblValue As Double) As String
        Return dblValue.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture)
    End Function

End Class
