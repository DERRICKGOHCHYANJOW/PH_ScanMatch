Imports System
Imports System.IO
Imports System.Threading
Public Class Form4
    Public curInd As Integer = 0
    Public d3 As DataTable
    Public dt As DataTable
    Private Sub Form4_Close(sender As Object, e As EventArgs) Handles MyBase.Closed
        Try
            Application.Exit()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myStream As Stream = Nothing
        Dim inpath, file1 As String
        Dim cardL(100) As String

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = OpenFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    inpath = Microsoft.VisualBasic.Left(OpenFileDialog1.FileName, Len(OpenFileDialog1.FileName) - (Len(OpenFileDialog1.SafeFileName) + 1))
                    file1 = OpenFileDialog1.SafeFileName
                    OpenFileDialog1.InitialDirectory = inpath
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open. 
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        Else
            MsgBox("Load Aborted!", MsgBoxStyle.Information)
            Exit Sub
        End If

        Try
            dt = GetCSVData(inpath, file1)
            dt = clearDT(dt)
            cardL = grabCardT(dt)
            d3 = grabFinal(dt)


            TextBox1.Text = Mid(dt.Rows(0).Item(0), InStr(dt.Rows(0).Item(0), ":") + 2)
            TextBox2.Text = dt.Rows.Count - 1 - 1 - 1
            TextBox3.Text = dt.Rows(dt.Rows.Count - 1).Item(dt.Columns.Count - 3) - dt.Rows(dt.Rows.Count - 2).Item(dt.Columns.Count - 3)
            TextBox7.Text = dt.Rows(dt.Rows.Count - 2).Item(dt.Columns.Count - 3)
            TextBox4.Text = d3.Rows(curInd).Item(2)
            TextBox5.Text = d3.Rows(curInd).Item(1)
            TextBox6.Text = d3.Rows(curInd).Item(0)
            TextBox8.Text = d3.Rows(curInd).Item(6)
            TextBox9.Text = d3.Rows(curInd).Item(5)
            TextBox10.Text = curInd + 1
            Button4.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Function GetCSVData(ByVal strFolderPath As String, ByVal strFileName As String) As DataTable
        Using LoadCon As New FileIO.TextFieldParser(strFolderPath & "/" & strFileName)
            LoadCon.TextFieldType = FileIO.FieldType.Delimited
            LoadCon.Delimiters = New String() {","}
            Dim conData As String()
            Dim ds As New DataTable()
            Dim firstD As String = ""


            Try
                conData = LoadCon.ReadFields()
                If conData.Length = 1 Then
                    firstD = conData(0)
                    conData = LoadCon.ReadFields()
                End If
                For cc = 0 To conData.Length - 1
                    ds.Columns.Add(New DataColumn("", GetType(String)))
                Next cc
            Catch ex As Exception
                MsgBox("execute error: " & ex.Message & Environment.NewLine & "Failed to Read File: ")
                Return Nothing
            End Try
            ds.Rows.Add()
            ds.Rows(0).Item(0) = firstD
            Try
                ds.Rows.Add(conData)
                While Not LoadCon.EndOfData
                    conData = LoadCon.ReadFields()
                    ds.Rows.Add(conData)
                End While
            Catch ex As Exception
                Return Nothing
            End Try
            Return ds
        End Using
    End Function
    Private Function clearDT(ByVal ds As DataTable)
        Try
            For i = ds.Rows.Count - 1 To 2 Step -1
                If ds.Rows(i).Item(50) = 0 Then
                    ds.Rows(i).Delete()
                End If
            Next
        Catch ex As Exception
            Return Nothing
        End Try
        Try
            For i = ds.Columns.Count - 3 To 3 Step -1
                If ds.Rows(ds.Rows.Count - 1).Item(i) = 0 Then
                    ds.Columns.RemoveAt(i)
                End If
            Next
        Catch ex As Exception
            Return Nothing
        End Try
        ds.AcceptChanges()
        Return ds
    End Function
    Private Function grabCardT(ByVal ds As DataTable)
        Dim i As Integer = 3
        Dim cc As Integer = 0
        Dim cardlist(100) As String

        Try
            Do While ds.Rows(1).Item(i) <> "TOTAL CARDS"
                cardlist(cc) = ds.Rows(1).Item(i)
                i = i + 1
                cc = cc + 1
            Loop
        Catch ex As Exception
            Return Nothing
        End Try
        ReDim Preserve cardlist(cc - 1)
        Return cardlist
    End Function
    Private Function grabFinal(ByVal ds As DataTable)
        Dim d2 As New DataTable
        Dim cc As Integer = 0
        Dim myC As Char = "0"c
        Dim tmpS As String = ""
        Dim tmpE As String = ""

        Try
            For i = 0 To 7
                d2.Columns.Add(New DataColumn("", GetType(String)))
            Next
        Catch ex As Exception
            Return Nothing
        End Try
        Try
            For i = 2 To ds.Rows.Count - 1
                d2.Rows.Add()
                d2.Rows(cc).Item(0) = ds.Rows(i).Item(0)
                d2.Rows(cc).Item(1) = (ds.Rows(i).Item(1)).padLeft(5, myC)
                d2.Rows(cc).Item(2) = ds.Rows(i).Item(2)
                For k = 3 To ds.Columns.Count - 4
                    If ds.Rows(i).Item(k) <> "" Then
                        If tmpS = "" Then
                            tmpS = ds.Rows(1).Item(k) & (ds.Rows(i).Item(1)).padLeft(5, myC) & (ds.Rows(i).Item(k)).padLeft(6, myC)
                            Dim hash = New System.Security.Cryptography.SHA256Managed().ComputeHash(System.Text.Encoding.UTF8.GetBytes(ds.Rows(1).Item(k) & (ds.Rows(i).Item(1)).padLeft(5, myC) & (ds.Rows(i).Item(k)).padLeft(6, myC)))
                            tmpE = Convert.ToBase64String(hash)
                        Else
                            tmpS = tmpS & "," & ds.Rows(1).Item(k) & (ds.Rows(i).Item(1)).padLeft(5, myC) & (ds.Rows(i).Item(k)).padLeft(6, myC)
                            Dim hash = New System.Security.Cryptography.SHA256Managed().ComputeHash(System.Text.Encoding.UTF8.GetBytes(ds.Rows(1).Item(k) & (ds.Rows(i).Item(1)).padLeft(5, myC) & (ds.Rows(i).Item(k)).padLeft(6, myC)))
                            tmpE = tmpE & "," & Convert.ToBase64String(hash)
                        End If
                    End If
                Next
                d2.Rows(cc).Item(3) = tmpS
                d2.Rows(cc).Item(4) = tmpE
                d2.Rows(cc).Item(5) = ds.Rows(i).Item(ds.Columns.Count - 3)
                d2.Rows(cc).Item(6) = ds.Rows(i).Item(ds.Columns.Count - 2)
                cc = cc + 1
                tmpS = ""
                tmpE = ""
            Next
        Catch ex As Exception
            Return Nothing
        End Try
        Return d2
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If IsNothing(d3) <> True Then
                If curInd <= d3.Rows.Count - 2 Then
                    curInd = curInd + 1
                    TextBox10.Text = curInd + 1
                    TextBox4.Text = d3.Rows(curInd).Item(2)
                    TextBox5.Text = d3.Rows(curInd).Item(1)
                    TextBox6.Text = d3.Rows(curInd).Item(0)
                    TextBox8.Text = d3.Rows(curInd).Item(6)
                    TextBox9.Text = d3.Rows(curInd).Item(5)
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If IsNothing(d3) <> True Then
                If curInd >= 1 Then
                    curInd = curInd - 1
                    TextBox10.Text = curInd + 1
                    TextBox4.Text = d3.Rows(curInd).Item(2)
                    TextBox5.Text = d3.Rows(curInd).Item(1)
                    TextBox6.Text = d3.Rows(curInd).Item(0)
                    TextBox8.Text = d3.Rows(curInd).Item(6)
                    TextBox9.Text = d3.Rows(curInd).Item(5)
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Me.Hide()
            Form5.ds = d3
            Form5.reconDS = dt
            Form5.Show()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            TextBox1.ReadOnly = True
            TextBox2.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox7.ReadOnly = True
            TextBox10.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox5.ReadOnly = True
            TextBox6.ReadOnly = True
            TextBox8.ReadOnly = True
            TextBox9.ReadOnly = True
            Button4.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

End Class