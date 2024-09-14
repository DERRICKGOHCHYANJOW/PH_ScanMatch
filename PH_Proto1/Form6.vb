Public Class Form6
    Public stSeq As String
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim s1 As Integer = 1
        Dim s2 As Integer = 9
        Dim e1 As Integer = 0
        Dim ts As String = ""
        ListBox1.Items.Clear()
        Try
            If Len(stSeq) > 9 Then
                ts = Replace(stSeq, ",", "")
                e1 = Len(ts)
                Do While s1 <= e1
                    ListBox1.Items.Add(Mid(ts, s1, s2))
                    s1 = s1 + s2
                Loop
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
End Class