Public Class splash
    Public cur As Form
    Private Sub splash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 2
        If ProgressBar1.Value <= 30 Then
            Label1.Text = "Initializing....."

        ElseIf ProgressBar1.Value <= 50 Then

            Label1.Text = "Loading components....."

        ElseIf ProgressBar1.Value <= 70 Then

            Label1.Text = "Integrating Database...."

        ElseIf ProgressBar1.Value <= 100 Then

            Label1.Text = "Please wait..."

        End If

        If ProgressBar1.Value = 100 Then

            Timer1.Dispose()

            Me.Visible = False

            cur = New splash()

            Ujian.Show()


        End If
    End Sub

  
End Class