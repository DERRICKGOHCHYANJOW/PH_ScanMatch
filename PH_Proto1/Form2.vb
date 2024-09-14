Imports System.IO

Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.Barcode
Imports iTextSharp.text.pdf.BarcodeQRCode

Public Class Form2
    Public d2 As DataTable
    Public brname As String = ""
    Public brnum As String = ""
    Public clname As String = ""
    Public sqNum As String = ""
    Public brCount As Integer = 0
    Public rmBr As Integer = 0
    Public recIn As Integer = 0
    Public stLot As String = ""
    Public uiLot As String = ""
    Public result As Boolean = False
    Public result2 As Boolean = False
    Public scnUID As String = ""
    Public cardType As String = ""
    Public seqNo As String = ""
    Public objWriter As StreamWriter
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            cardType = Trim(d2.Rows(0).Item(9))
            TextBox1.Text = Trim(d2.Rows(0).Item(0))
            TextBox2.Text = ""
            TextBox2.Focus()
            TextBox1.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox5.ReadOnly = True
            TextBox6.ReadOnly = True
            TextBox7.ReadOnly = True
            TextBox8.ReadOnly = True
            Label10.Visible = True
            Label12.Visible = True
            Label14.Visible = True
            Label9.Visible = False
            Label11.Visible = False
            Label13.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub Form2_Close(sender As Object, e As EventArgs) Handles MyBase.Closed
        Try
            Application.Exit()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Private Sub ProcessInput(ByVal myText As String)
        Try
            If Len(myText) = 12 Then
                If Mid(myText, 1, 2) = "PH" Then
                    If result2 = False Then
                        If InStr(scnUID, myText) = 0 Then
                            result = GrabRec(d2, myText, brname, brnum, clname, sqNum, brCount, recIn, stLot, uiLot)
                            If result = False Then
                                MsgBox(myText & " not found in this workorder!!", MsgBoxStyle.Critical, "Wrong Pak")
                                If LoginForm1.p1logs = 1 Then
                                    genLog("3", TextBox1.Text, brnum, brCount, scnUID)
                                End If
                            Else
                                result2 = True
                                scnUID = scnUID & myText & ","
                                TextBox4.Text = brname
                                TextBox5.Text = brnum
                                TextBox6.Text = clname
                                TextBox7.Text = sqNum
                                TextBox3.Text = brCount
                                rmBr = brCount - 1
                                TextBox8.Text = rmBr
                                If rmBr = 0 Then
                                    If LoginForm1.mpbox = 1 Then
                                        MsgBox(brnum & " branch is completed! Prepare to scan Next Branch", MsgBoxStyle.Information, "Done")
                                    End If
                                    greenlight()
                                    genLabel(brname, brnum, clname, brCount, TextBox1.Text)
                                    If LoginForm1.p1logs = 1 Then
                                        genLog("1", TextBox1.Text, brnum, brCount, scnUID)
                                    End If
                                    result2 = False
                                    stLot = ""
                                    uiLot = ""
                                    TextBox4.Clear()
                                    TextBox5.Clear()
                                    TextBox6.Clear()
                                    TextBox7.Clear()
                                    TextBox3.Clear()
                                    TextBox8.Clear()
                                Else
                                    seqNo = grabSEQ(d2, myText)
                                    TextBox7.Text = seqNo
                                    upDateSeq(seqNo, uiLot)
                                    amberlight()
                                End If
                            End If
                        Else
                            redlight()
                            MsgBox(myText & " has been scanned before!! Check Duplication!!", MsgBoxStyle.Critical, "Pak scanned BEFORE")
                            If LoginForm1.p1logs = 1 Then
                                genLog("2", TextBox1.Text, brnum, brCount, scnUID)
                            End If
                            amberlight()
                        End If
                    Else
                        If result = False Then
                            redlight()
                            MsgBox(myText & " not found in this workorder!!", MsgBoxStyle.Critical, "Wrong Pak")
                            If LoginForm1.p1logs = 1 Then
                                genLog("3", TextBox1.Text, brnum, brCount, scnUID)
                            End If
                            amberlight()
                        Else
                            If InStr(stLot, myText) <> 0 Then
                                If InStr(scnUID, myText) = 0 Then
                                    scnUID = scnUID & myText & ","
                                    rmBr = rmBr - 1
                                    TextBox8.Text = rmBr
                                    If rmBr = 0 Then
                                        If LoginForm1.mpbox = 1 Then
                                            MsgBox(brnum & " branch is completed! Prepare to scan Next Branch", MsgBoxStyle.Information, "Done")
                                        End If
                                        greenlight()
                                        genLabel(brname, brnum, clname, brCount, TextBox1.Text)
                                        If LoginForm1.p1logs = 1 Then
                                            genLog("1", TextBox1.Text, brnum, brCount, scnUID)
                                        End If
                                        result2 = False
                                        stLot = ""
                                        uiLot = ""
                                        TextBox4.Clear()
                                        TextBox5.Clear()
                                        TextBox6.Clear()
                                        TextBox7.Clear()
                                        TextBox3.Clear()
                                        TextBox8.Clear()
                                    Else
                                        seqNo = grabSEQ(d2, myText)
                                        TextBox7.Text = seqNo
                                        upDateSeq(seqNo, uiLot)
                                        amberlight()
                                    End If
                                Else
                                    redlight()
                                    MsgBox(myText & " has been scanned before!! Check Duplication!!", MsgBoxStyle.Critical, "Pak scanned BEFORE")
                                    If LoginForm1.p1logs = 1 Then
                                        genLog("2", TextBox1.Text, brnum, brCount, scnUID)
                                    End If
                                    amberlight()
                                End If
                            Else
                                redlight()
                                MsgBox(myText & " does not belong to this Branch!!", MsgBoxStyle.Critical, "Wrong Pak")
                                If LoginForm1.p1logs = 1 Then
                                    genLog("3", TextBox1.Text, brnum, brCount, scnUID)
                                End If
                                amberlight()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        Try
            If e.KeyData = Keys.Enter Then
                ProcessInput(TextBox2.Text)
                TextBox2.Clear()
                e.SuppressKeyPress = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Function GrabRec(ByVal dt As DataTable, id1 As String, ByRef bname As String, ByRef bnum As String, ByRef clname As String, ByRef sqN As String, ByRef bcount As Integer, ByRef bindex As Integer, ByRef bLot As String, ByRef uLot As String) As Boolean
        Dim fRec As Boolean = False
        Try
            For i = 0 To dt.Rows.Count - 1
                If id1 = dt.Rows(i).Item(2) Then
                    bname = dt.Rows(i).Item(5)
                    bnum = dt.Rows(i).Item(6)
                    clname = dt.Rows(i).Item(7)
                    sqN = dt.Rows(i).Item(3)
                    bcount = findBranchC(dt, bnum)
                    bindex = dt.Rows(i).Item(1)
                    bLot = findRecUID(dt, bnum, bcount)
                    uLot = findRecSID(dt, bnum, bcount)
                    fRec = True
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        If fRec = False Then
            Return False
        Else
            Return True
        End If
    End Function
    Function findBranchC(ByVal dt As DataTable, ByVal bnum As String) As Integer
        Dim k As Integer = 0
        Try
            For i = 0 To dt.Rows.Count - 1
                If bnum = dt.Rows(i).Item(6) Then
                    k = k + 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        Return k
    End Function
    Function findRecUID(ByVal dt As DataTable, ByVal bnum As String, ByVal cc As Integer) As String
        Dim ss As String = ""
        Try
            For i = 0 To dt.Rows.Count - 1
                If bnum = dt.Rows(i).Item(6) Then
                    For j = 0 To cc - 1
                        ss = ss & dt.Rows(i + j).Item(2) & ","
                    Next
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        Return ss
    End Function
    Function findRecSID(ByVal dt As DataTable, ByVal bnum As String, ByVal cc As Integer) As String
        Dim ss As String = ""
        For i = 0 To dt.Rows.Count - 1
            If bnum = dt.Rows(i).Item(6) Then
                For j = 0 To cc - 1
                    ss = ss & dt.Rows(i + j).Item(3) & ","
                Next
                Exit For
            End If
        Next
        Return ss
    End Function
    Function grabSEQ(ByVal dt As DataTable, ByVal sid As String) As String
        Dim ss As String = ""
        Try
            For i = 0 To dt.Rows.Count - 1
                If sid = dt.Rows(i).Item(2) Then
                    ss = dt.Rows(i).Item(3)
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        Return ss
    End Function
    Sub redlight()
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
    Sub amberlight()
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
    Sub greenlight()
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

    Sub genLabel(ByVal branchN As String, ByVal BrCode As String, ByVal clcode As String, ByVal toCnt As String, ByVal meWO As String)
        Dim myTemp As String
        Dim pgSize As New iTextSharp.text.Rectangle((LoginForm1.lbx / 25.4) * 72, (LoginForm1.lby / 25.4) * 72)
        Dim bcSize As New iTextSharp.text.Rectangle(10, 10)
        Dim doc As New iTextSharp.text.Document(pgSize, 0, 0, 0, 0)
        Dim writer As PdfWriter = PdfWriter.GetInstance(doc, New FileStream(My.Application.Info.DirectoryPath & "\label\" & meWO & BrCode & "_Label.pdf", FileMode.Create))
        Dim Lucida As Font = FontFactory.GetFont("LUCIDA CONSOLE", 8, 0)
        Dim lprint As String = ""

        doc.Open()
        Try
            Dim cb1 As PdfContentByte = writer.DirectContent
            Dim ct1 As New ColumnText(cb1)
            myTemp = "BRANCH  : " & branchN
            Dim c1 As New Chunk(myTemp, Lucida)
            Dim myText1 As New Phrase(c1)
            ct1.SetSimpleColumn(myText1, ((LoginForm1.lb1x / 25.4) * 72), ((LoginForm1.lb1y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb1y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            myTemp = "CODE       : " & BrCode
            Dim c2 As New Chunk(myTemp, Lucida)
            Dim myText2 As New Phrase(c2)
            ct1.SetSimpleColumn(myText2, ((LoginForm1.lb2x / 25.4) * 72), ((LoginForm1.lb2y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb2y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            myTemp = "CLUSTER : " & clcode
            Dim c3 As New Chunk(myTemp, Lucida)
            Dim myText3 As New Phrase(c3)
            ct1.SetSimpleColumn(myText3, ((LoginForm1.lb3x / 25.4) * 72), ((LoginForm1.lb3y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb3y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            myTemp = "COUNT     : " & toCnt
            Dim c4 As New Chunk(myTemp, Lucida)
            Dim myText4 As New Phrase(c4)
            ct1.SetSimpleColumn(myText4, ((LoginForm1.lb4x / 25.4) * 72), ((LoginForm1.lb4y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb4y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            myTemp = "WORKORDER : " & meWO
            Dim c5 As New Chunk(myTemp, Lucida)
            Dim myText5 As New Phrase(c5)
            ct1.SetSimpleColumn(myText5, ((LoginForm1.lb5x / 25.4) * 72), ((LoginForm1.lb5y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb5y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            myTemp = "TYPE        : " & cardType
            Dim c6 As New Chunk(myTemp, Lucida)
            Dim myText6 As New Phrase(c6)
            ct1.SetSimpleColumn(myText6, ((LoginForm1.lb6x / 25.4) * 72), ((LoginForm1.lb6y / 25.4) * 72), ((LoginForm1.lbx / 25.4) * 72), ((LoginForm1.lb6y / 25.4) * 72) + 15, 15, Element.ALIGN_LEFT)
            ct1.Go()

            Dim cb11 As PdfContentByte = writer.DirectContent

            Dim padCnt As String
            Dim myC As Char = "0"c
            padCnt = toCnt.PadLeft(6, myC)

            '------------------------------------Generate QRcode-------------------------------------------
            '===============================================================================================
            Dim tmpS As String = ""
            Dim hash = New System.Security.Cryptography.SHA256Managed().ComputeHash(System.Text.Encoding.UTF8.GetBytes(cardType & BrCode & padCnt))
            tmpS = Convert.ToBase64String(hash)

            Dim qrcode As BarcodeQRCode = New BarcodeQRCode(tmpS, 1, 1, Nothing)
            Dim qrcodeImage As Image = qrcode.GetImage()
            qrcodeImage.SetAbsolutePosition(LoginForm1.lbBarX, LoginForm1.lbBarY)
            qrcodeImage.ScalePercent(150)
            doc.Add(qrcodeImage)

        Catch ex As Exception
            doc.Close()
            Return
            Exit Sub
        End Try
        doc.Close()

        If LoginForm1.lbShow <> 0 Then
            Form3.woID = meWO
            Form3.BRid = BrCode
            Form3.Show()
        End If
        If LoginForm1.lbPrint <> 0 Then
            Try
                Dim MyProcess As New Process
                MyProcess.StartInfo.CreateNoWindow = False
                MyProcess.StartInfo.Verb = "print"
                MyProcess.StartInfo.FileName = My.Application.Info.DirectoryPath & "/Label/" & meWO & BrCode & "_Label.pdf"
                MyProcess.Start()
                MyProcess.WaitForExit(7000)
                Try
                    MyProcess.CloseMainWindow()
                    MyProcess.Close()
                Catch ex As Exception

                End Try
            Catch ex As Exception
                MessageBox.Show(ex.Message, "error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Sub genLog(ByVal chkV As String, ByVal workID As String, ByVal brCo As String, ByVal brCnt As String, ByVal myuID As String)
        Dim myUser As String = ""
        Dim myCom As String = ""
        Dim myDate, myTime As String
        Dim zipfile As String


        myUser = Environment.UserName
        myCom = System.Windows.Forms.SystemInformation.ComputerName

        myDate = Now.ToString("dd/MM/yyyy")
        myTime = Now.ToString("HH:mm:ss")

        Dim recLog As String = "Process1_ScanReport" & "_" & Now.ToString("dd") & Now.ToString("MM") & Now.ToString("yyyy") & ".csv"

        Try
            zipfile = "ProcessOne_report" + "_" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy") + ".zip"
            Dim exePath As String = LoginForm1.ziplocation
            Dim args1 As String = " e " + """" + My.Application.Info.DirectoryPath + "\1Logs\" + zipfile + """" + " -pGEM" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy") + " -o" + """" + My.Application.Info.DirectoryPath + "\1Logs\" + """"
            Dim args2 As String = " a " + """" + My.Application.Info.DirectoryPath + "\1Logs\" + zipfile + """" + " " + """" + My.Application.Info.DirectoryPath & "\1Logs\" & recLog + """" + " -pGEM" + Now.ToString("dd") + Now.ToString("MM") + Now.ToString("yyyy")

            If File.Exists((My.Application.Info.DirectoryPath + "\1Logs\" + zipfile)) Then
                System.Diagnostics.Process.Start(exePath, args1)
                Threading.Thread.Sleep(2000)
                File.Delete((My.Application.Info.DirectoryPath + "\1Logs\" + zipfile))
            End If

            If Not File.Exists((My.Application.Info.DirectoryPath + "\1Logs\" + recLog)) Then
                objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\1Logs\" & recLog, True)
                objWriter.WriteLine("Computer," & "User," & "DATE," & "TIME," & "WORKORDER," & "BRANCH," & "COUNT," & "UID," & "STATUS")
                objWriter.Close()
            End If
            Select Case chkV
                Case "1"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\1Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & "," & workID & "," & brCo & "," & brCnt & "," & ",COMPLETE")
                    objWriter.Close()
                Case "2"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\1Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & ",WorkOrder:" & workID & ",Branch:" & brCo & ",Paks:" & brCnt & ",UID:" & myuID & ",DOUBLE")
                    objWriter.Close()
                Case "3"
                    objWriter = My.Computer.FileSystem.OpenTextFileWriter(My.Application.Info.DirectoryPath & "\1Logs\" & recLog, True)
                    objWriter.WriteLine(myCom & "," & myUser & "," & myDate & "," & myTime & ",WorkOrder:" & workID & ",Branch:" & brCo & ",Paks:" & brCnt & ",UID:" & myuID & ",WRONG")
                    objWriter.Close()
            End Select
            If File.Exists((My.Application.Info.DirectoryPath + "\1Logs\" + recLog)) Then
                System.Diagnostics.Process.Start(exePath, args2)
                Threading.Thread.Sleep(2000)
                File.Delete((My.Application.Info.DirectoryPath + "\1Logs\" + recLog))
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
    Private Sub upDateSeq(ByVal mSq As String, ByRef uLot As String)
        Try
            If InStr(uLot, mSq) <> 0 Then
                uLot = Replace(uLot, mSq & ",", "")
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Form6.stSeq = uiLot
            Form6.Show()
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
    End Sub
End Class