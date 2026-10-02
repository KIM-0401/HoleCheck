Option Strict On
Option Explicit On

' 디자이너 파일 없이 코드로 만든 빈 폼 · W04는 Show()까지만
Public Class HoleCheckForm
    Inherits System.Windows.Forms.Form

    Private invApplication As Inventor.Application

    Public Sub New(ByVal invInventorApplication As Inventor.Application)
        MyBase.New()

        invApplication = invInventorApplication

        Me.Text = "HoleCheck"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Size = New System.Drawing.Size(480, 360)
    End Sub

End Class
