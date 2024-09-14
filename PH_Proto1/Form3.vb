Public Class Form3
    Public woID, BRid As String
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Form2.Hide()
            AxAcroPDF1.src = My.Application.Info.DirectoryPath & "\label\" & woID & BRid & "_Label.pdf"
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Form3_Close(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Form2.Show()
            If LoginForm1.lbSave <> 0 Then
                Kill(My.Application.Info.DirectoryPath & "\label\" & woID & BRid & "_Label.pdf")
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
End Class