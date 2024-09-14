Imports System.IO

Public Class Form5
    Public ds As DataTable
    Public reconDS As DataTable
    Public BrName As String = ""
    Public BrCode As String = ""
    Public ClCode As String = ""
    Public Tcnt As String = ""
    Public Tpak As String = ""
    Public CardT As String = ""
    Public PakC As String = ""
    Public ScanList As String = ""
    Public Spak As Integer = 0
    Public STc As Integer = 0
    Public result As Boolean = False
    Public result2 As Boolean = False
    Public currList As String = ""
    Public currList2 As String = ""
    Public objWriter As StreamWriter
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            TextBox1.Text = ""
            TextBox1.Focus()

            Label10.Visible = True
            Label12.Visible = True
            Label14.Visible = True
            Label9.Visible = False
            Label11.Visible = False
            Label13.Visible = False

            TextBox2.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox5.ReadOnly = True
            TextBox6.ReadOnly = True
            TextBox7.ReadOnly = True
            TextBox8.ReadOnly = True
            TextBox9.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

    End Sub
    Private Sub Form5_Close(sender As Object, e As EventArgs) Handles MyBase.Closed
        Try
            Application.Exit()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Try
            If e.KeyData = Keys.Enter Then
                If TextBox1.TextLength > 1 Then
                    If result2 = False Then
                        grabRec(ds, TextBox1.Text, BrName, BrCode, ClCode, Tcnt, Tpak, result)
                    End If
                    ProcessInput(TextBox1.Text)
                    TextBox1.Clear()
                    e.SuppressKeyPress = True
                Else
                    TextBox1.Clear()
                    TextBox1.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub ProcessInput(ByVal myText As String)
        Try
            If result = True Then
                result2 = True
                If InStr(currList, myText) <> 0 Then
                    scanStr(myText, currList, currList2, CardT, PakC)
                    If InStr(ScanList, myText) = 0 Then
                        TextBox2.Text = BrName
                        TextBox3.Text = BrCode
                        TextBox4.Text = ClCode
                        TextBox5.Text = Tcnt
                        TextBox6.Text = Tpak
                        TextBox7.Text = CardT
                        STc = STc + 1
                        TextBox8.Text = CInt(Tcnt) - STc
                        Spak = Spak + CInt(PakC)
                        TextBox9.Text = CInt(Tpak) - Spak
                        ScanList = ScanList & myText & ","
                        If CInt(TextBox8.Text) = 0 Then
                            If LoginForm1.mpbox = 1 Then
                                MsgBox(CardT & " group is completed! Prepare to scan Next Group", MsgBoxStyle.Information, "Done")
                            End If
                            greenlight()
                            UpdateS(BrCode)
                            If LoginForm1.p2logs = 1 Then
                                genLog("1", BrName, BrCode, Tpak)
                            End If
                            TextBox2.Clear()
                            TextBox3.Clear()
                            TextBox4.Clear()
                            TextBox5.Clear()
                            TextBox6.Clear()
                            TextBox7.Clear()
                            TextBox8.Clear()
                            TextBox9.Clear()
                            Spak = 0
                            STc = 0
                            result2 = False
                        End If
                    Else
                        redlight()
                        MsgBox(myText & " has been scanned before!! Check Duplication!!", MsgBoxStyle.Critical, "Pak scanned BEFORE")
                        amberlight()
                        If LoginForm1.p2logs = 1 Then
                            genLog("2", BrName, BrCode, Tpak)
                        End If
                    End If
                Else
                    redlight()
                    MsgBox(myText & " not found in SAME GROUP transmittal!!", MsgBoxStyle.Critical, "NOT in SAME GROUP")
                    amberlight()
                    If LoginForm1.p2logs = 1 Then
                        genLog("3", ",", ",", ",")
                    End If
                End If
            Else
                redlight()
                MsgBox(myText & " not found in transmittal!!", MsgBoxStyle.Critical, "Wrong Pak")
                amberlight()
                If LoginForm1.p2logs = 1 Then
                    genLog("3", "-", "-", "-")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub grabRec(ByVal dt As DataTable, ByVal chkRec As String, ByRef BrN As String, ByRef BrC As String, ByRef ClC As String, ByRef Tc As String, ByRef Tp As String, ByRef res As Boolean)
        Try
            For i = 0 To dt.Rows.Count - 2
                If InStr(dt.Rows(i).Item(4), chkRec) <> 0 Then
                    BrN = dt.Rows(i).Item(2)
                    BrC = dt.Rows(i).Item(1)
                    ClC = dt.Rows(i).Item(0)
                    Tc = dt.Rows(i).Item(6)
                    Tp = dt.Rows(i).Item(5)
                    currList = dt.Rows(i).Item(4)
                    currList2 = dt.Rows(i).Item(3)
                    res = True
                    Exit For
                Else
                    res = False
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub redlight()
        Try
            Label9.Visible = True
            Label10.Visible = False
            Label11.Visible = False
            Label14.Visible = False
            Label13.Visible = True
            Label12.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub amberlight()
        Try
            Label9.Visible = False
            Label10.Visible = True
            Label11.Visible = False
            Label14.Visible = True
            Label13.Visible = False
            Label12.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub greenlight()
        Try
            Label9.Visible = False
            Label10.Visible = False
            Label11.Visible = True
            Label14.Visible = True
            Label13.Visible = True
            Label12.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Private Sub scanStr(ByVal s0 As String, ByVal s1 As String, ByVal s2 As String, ByRef cardTy As String, ByRef myP As String)
        Dim cindex As Integer = 1
        Dim tmpS As String = ""
        Dim tmpR As String = ""
        s1 = s1 & ","
        s2 = s2 & ","
        Try
            Do While tmpS <> s0
                tmpS = Mid(s1, 1, InStr(s1, ",") - 1)
                If tmpS <> s0 Then
                    s1 = Replace(s1, tmpS & ",", "")
                    cindex = cindex + 1
                End If
            Loop
            tmpR = Mid(s2, 1, InStr(s2, ",") - 1)
            If cindex <> 1 Then
                For k = 1 To cindex - 1
                    s2 = Replace(s2, tmpR & ",", "")
                    tmpR = Mid(s2, 1, InStr(s2, ",") - 1)
                Next
            End If
            cardTy = Mid(tmpR, 1, InStr(tmpR, "0") - 1)

            myP = Mid(tmpR, InStrRev(tmpR, "0", Len(tmpR) - 2) + 1)
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub UpdateS(ByVal brC As String)
        Dim myC As Char = "0"c
        Try
            For i = 2 To reconDS.Rows.Count - 2
                If reconDS.Rows(i).Item(1).padLeft(5, myC) = brC Then
                    reconDS.Rows(i).Item(reconDS.Columns.Count - 1) = "COMPLETED"
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub genLog(ByVal chkV As String, ByVal brN As String, ByVal brC As String, ByVal Tp As String)
        Dim myUser As String = ""
        Dim myCom As String = ""
        Dim myDate, myTime As String
        Dim zipfile As String


        myUser = Environment.UserName
        myCom = System.Windows.Forms.SystemInformation.ComputerName

        myDate = Now.ToString("dd/MM/yyyy")
        myTime = Now.ToString("HH:mm:ss")
        Try
            Dim recLog As String = "Process2_ScanReport" & "_" & Now.ToString("dd") & Now.ToString("MM") & Now.ToString("yyyy") & ".csv"

            zipfile = "ProcessTWO_report" + "_" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy") + ".zip"
            Dim exePath As String = LoginForm1.ziplocation
            Dim args1 As String = " e " + """" + My.Application.Info.DirectoryPath + "\2Logs\" + zipfile + """" + " -pGEM" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy") + " -o" + """" + My.Application.Info.DirectoryPath + "\2Logs\" + """"
            Dim args2 As String = " a " + """" + My.Application.Info.DirectoryPath + "\2Logs\" + zipfile + """" + " " + """" + My.Application.Info.DirectoryPath & "\2Logs\" & recLog + """" + " -pGEM" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy")

            If File.Exists((My.Application.Info.DirectoryPath + "\2Logs\" + zipfile)) Then
                System.Diagnostics.Process.Start(exePath, args1)
                Threading.Thread.Sleep(2000)
                File.Delete((My.Application.Info.DirectoryPath + "\2Logs\" + zipfile))
            End If

            If Not File.Exists((My.Application.Info.DirectoryPath + "\2Logs\" + recLog)) Then
                objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\2Logs\" & recLog, True)
                objWriter.WriteLine("Computer," & "User," & "DATE," & "TIME," & "BRANCH," & "BRCODE," & "COUNT," & "STATUS")
                objWriter.Close()
            End If
            Select Case chkV
                Case "1"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\2Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & "," & brN & "," & brC & "," & Tp & ",COMPLETE")
                    objWriter.Close()
                Case "2"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\2Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & "," & brN & "," & brC & "," & Tp & ",DOUBLE")
                    objWriter.Close()
                Case "3"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\2Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & "," & brN & "," & brC & "," & Tp & ",WRONG")
                    objWriter.Close()
            End Select

            If File.Exists((My.Application.Info.DirectoryPath + "\2Logs\" + recLog)) Then
                System.Diagnostics.Process.Start(exePath, args2)
                Threading.Thread.Sleep(2000)
                File.Delete((My.Application.Info.DirectoryPath + "\2Logs\" + recLog))
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim recLog As String = "P2_Card_Receiving_Report" & "_" & Now.ToString("yyyy") & Now.ToString("MM") & Now.ToString("dd") & ".csv"
        Try
            For i = 0 To reconDS.Rows.Count - 1
                For j = 0 To reconDS.Columns.Count - 1
                    If j = reconDS.Columns.Count - 1 Then
                        If IsDBNull(reconDS.Rows(i).Item(j)) Then
                            objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\Report\" & recLog, True)
                            objWriter.WriteLine(",")
                            objWriter.Close()
                        Else
                            objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\Report\" & recLog, True)
                            objWriter.WriteLine(reconDS.Rows(i).Item(j))
                            objWriter.Close()
                        End If
                    Else
                        If IsDBNull(reconDS.Rows(i).Item(j)) Then
                            objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\Report\" & recLog, True)
                            objWriter.Write(",")
                            objWriter.Close()
                        Else
                            objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\Report\" & recLog, True)
                            objWriter.Write(reconDS.Rows(i).Item(j) & ",")
                            objWriter.Close()
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        MsgBox("Report Generated", MsgBoxStyle.Information, "Done")
        TextBox1.Focus()
    End Sub
End Class