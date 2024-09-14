Imports System
Imports System.Data.OleDb
Imports System.IO
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Public Class Form1
    Public curIndex As Integer
    Public ds As DataTable

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If IsNothing(ds) <> True Then
                If curIndex > 1 Then
                    curIndex = curIndex - 1
                    TextBox9.Text = curIndex
                    TextBox3.Text = ds.Rows(curIndex - 1).Item(2)
                    TextBox4.Text = ds.Rows(curIndex - 1).Item(5)
                    TextBox5.Text = ds.Rows(curIndex - 1).Item(6)
                    TextBox6.Text = ds.Rows(curIndex - 1).Item(7)
                    TextBox7.Text = ds.Rows(curIndex - 1).Item(3)
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
            Form2.d2 = ds
            Form2.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If IsNothing(ds) <> True Then
                If curIndex <= ds.Rows.Count - 1 Then
                    curIndex = curIndex + 1
                    TextBox9.Text = curIndex
                    TextBox3.Text = ds.Rows(curIndex - 1).Item(2)
                    TextBox4.Text = ds.Rows(curIndex - 1).Item(5)
                    TextBox5.Text = ds.Rows(curIndex - 1).Item(6)
                    TextBox6.Text = ds.Rows(curIndex - 1).Item(7)
                    TextBox7.Text = ds.Rows(curIndex - 1).Item(3)
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cc1 As String = ""

        Dim oradb As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.165.149.52)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=PCMS.MANILA.GEMALTO.COM)));User Id=" & LoginForm1.myUser & ";Password=" & LoginForm1.myOPass & ";"
        Dim conn As New OracleConnection(oradb)
        Dim myquery As String
        Dim cmd As New OracleCommand
        Try

            If TextBox1.Text <> "" Then
                If Len(TextBox1.Text) >= 5 Then
                    If LoginForm1.Detest = 0 Then
                        Try
                            Try
                                conn.Open()
                            Catch ex As OracleException
                                'MsgBox("execute dbAccess: " & ex.Message & Environment.NewLine & "Aufruf: ")
                                MsgBox("Error querying database. Unable to connect. " & ex.Message)
                                Return
                            End Try
                        Catch ex As Exception
                            MsgBox("Error querying database. Unable to connect. " & ex.Message)
                        End Try
                        Dim dt As New DataSet
                        myquery = "select trim(to_char(c.workorderid,'XXXXXX')) as workorderid,c.indexnumber as ind,c.uniqueidentifier as myId,multipleaccountkey as SEQNO " &
                             ",get_token(c.exportedkeyvalue4, 51,'~') as BRACHOFACC,get_token(c.exportedkeyvalue4, 51,'~') as BRANCHNAME, get_token(c.exportedkeyvalue4, 50,'~') as BRANCHCODE " &
                             ",get_token(c.exportedkeyvalue4, 52,'~') as CLUSTERCODE, get_token(c.exportedkeyvalue4, 53,'~') as CLUSTERNAME, get_token(c.exportedkeyvalue4, 4,'~') as PRODUCTTYPE " &
                             "from (select uniqueidentifier, indexnumber, workorderid,multipleaccountkey,exportedkeyvalue4 from card  union all Select uniqueidentifier,indexnumber, workorderid,multipleaccountkey,exportedkeyvalue4 " &
                             "from card_arc) c where trim(to_char(c.workorderid,'XXXXXX')) ='" & UCase(Trim(TextBox1.Text)) & "' order by c.indexnumber"
                        Try
                            Using myAdapter As New OracleDataAdapter(myquery, conn)
                                myAdapter.Fill(dt)
                            End Using
                        Catch ex As Exception
                            MsgBox("Error querying database. Failed Grabbing query. " & ex.Message)
                            Return
                        End Try
                        conn.Dispose()
                        ds = dt.Tables(0)
                    End If
                    Try
                        If LoginForm1.Detest = 1 Then
                            Dim filen As String = "testSample.xls"
                            Dim pathn As String = My.Application.Info.DirectoryPath
                            ds = GetXlsData8(pathn, filen)
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
                    End Try
                    ds.Columns.Add(New DataColumn("", GetType(String)))
                    curIndex = ds.Rows(0).Item(1)
                    TextBox9.Text = curIndex
                    cc1 = GetBrCount(ds)
                    TextBox2.Text = cc1
                    TextBox8.Text = ds.Rows.Count
                    TextBox3.Text = ds.Rows(curIndex - 1).Item(2)
                    TextBox4.Text = ds.Rows(curIndex - 1).Item(5)
                    TextBox5.Text = ds.Rows(curIndex - 1).Item(6)
                    TextBox6.Text = ds.Rows(curIndex - 1).Item(7)
                    TextBox7.Text = ds.Rows(curIndex - 1).Item(3)
                    Button4.Visible = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If LoginForm1.dbcheck = 0 Then
                '-------------------------------------------------------------------------------------------
                'test connection to the oracle database
                '-------------------------------------------------------------------------------------------
                Dim oradb As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.165.149.52)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=PCMS.MANILA.GEMALTO.COM)));User Id=" & LoginForm1.myUser & ";Password=" & LoginForm1.myOPass & ";"
                Dim conn As New OracleConnection(oradb)

                Try
                    Try
                        conn.Open()              'try wo table to see if access to database is working
                    Catch ex As OracleException
                        'MsgBox("execute dbAccess: " & ex.Message & Environment.NewLine & "Aufruf: ")
                        MsgBox("Error querying database. Unable to connect. " & ex.Message)
                        Return
                    End Try
                    conn.Dispose()
                Catch ex As OracleException
                    'MsgBox("connection: " & ex.Message & " > " & ex.Source)
                    MsgBox("Error querying database. Unable to connect. " & ex.Message)
                End Try

            End If
            If LoginForm1.Detest = 1 Then
                TextBox1.Text = "123456"
                Button4.Visible = False
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                TextBox4.ReadOnly = True
                TextBox5.ReadOnly = True
                TextBox6.ReadOnly = True
                TextBox7.ReadOnly = True
                TextBox8.ReadOnly = True
                TextBox9.ReadOnly = True
            Else
                TextBox2.ReadOnly = True
                TextBox3.ReadOnly = True
                TextBox4.ReadOnly = True
                TextBox5.ReadOnly = True
                TextBox6.ReadOnly = True
                TextBox7.ReadOnly = True
                TextBox8.ReadOnly = True
                TextBox9.ReadOnly = True
                Button4.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

    End Sub
    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.Closed
        Try
            Application.Exit()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Function GetXlsData8(ByVal strFolderPath As String, ByVal strFileName As String) As DataTable
        Dim strConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFolderPath & "\" & strFileName & ";Extended Properties='Excel 8.0;IMEX=1;'"
        Dim conn As New OleDb.OleDbConnection(strConnString)
        Dim mychk As Boolean = False
        Dim myRR As Integer = 0
        Try
            conn.Open()
            Dim dt As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing) 'This will give u list of all the worksheets in the excel file 
            For i = 0 To dt.Rows.Count - 1
                If InStr(dt.Rows(i).Item(2), "Page 1") <> 0 Then
                    mychk = True
                    myRR = i
                    Exit For
                Else
                    If InStr(dt.Rows(i).Item(2), "Sheet1$") <> 0 Then
                        mychk = True
                        myRR = i
                        Exit For
                    End If
                End If
            Next
            If mychk = True Then
                Dim cmd As New OleDb.OleDbCommand()

                cmd.CommandText = "SELECT * FROM [" & Replace(dt.Rows(myRR).Item(2), "'", "") & "]"
                cmd.Connection = conn

                Dim da As New OleDb.OleDbDataAdapter()

                da.SelectCommand = cmd

                Dim ds As New DataSet()
                da.Fill(ds)
                da.Dispose()

                Return ds.Tables(0)
            Else

                Return Nothing
                Exit Function
            End If
        Catch
            Return Nothing
        Finally
            conn.Close()
        End Try
    End Function
    Private Function GetBrCount(ByVal d1 As DataTable) As Integer
        Dim count As Integer = 1
        Dim iniBr As String = ""
        Dim chkbr As String = ""

        Try

            iniBr = d1.Rows(0).Item(5)
            For i = 1 To d1.Rows.Count - 1
                chkbr = d1.Rows(i).Item(5)
                If chkbr <> iniBr Then
                    iniBr = chkbr
                    count = count + 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        Return count
    End Function

End Class

