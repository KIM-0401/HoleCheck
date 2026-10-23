Option Strict On
Option Explicit On

' W07 7-2 · 더블클릭 카메라 이동 (HoleCheckForm.OnGridCellDoubleClick에서 이동)
Friend Module Zoom

    ' 더블클릭 · 계획서 5주차대로 Camera.Fit 후 목표점 지정 · Eye는 계산하지 않는다
    ' 목표점은 mm · 예외는 호출자(폼)가 처리한다
    Friend Sub MoveCameraToTarget(ByVal invApplication As Inventor.Application, ByVal dblTargetX As Double, ByVal dblTargetY As Double, ByVal dblTargetZ As Double)
        Dim invView As Inventor.View = invApplication.ActiveView
        If invView Is Nothing Then
            System.Windows.Forms.MessageBox.Show("활성 뷰가 없습니다.", "HoleCheck")
            Return
        End If

        ' 목표점 mm → cm
        Dim invTarget As Inventor.Point = invApplication.TransientGeometry.CreatePoint(dblTargetX / dblCentimeterToMillimeter,
                                                                                        dblTargetY / dblCentimeterToMillimeter,
                                                                                        dblTargetZ / dblCentimeterToMillimeter)

        Dim invCamera As Inventor.Camera = invView.Camera
        invCamera.Fit()
        invCamera.Target = invTarget
        invCamera.Apply()
    End Sub

End Module
