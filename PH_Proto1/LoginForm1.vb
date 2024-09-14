Imports System.IO
Public Class LoginForm1
    Public myUser, myPass, myOPass, dbcheck, Detest, mpbox, p1logs, rmode, ziplocation, p2logs As String
    Public lbPrint, lbShow, lbSave, lbx, lby, lb1x, lb2x, lb3x, lb4x, lb5x, lb6x, lb1y, lb2y, lb3y, lb4y, lb5y, lb6y, lbBarX, lbBarY As String
    Public myUser1, myUser2, myUser3, myUser4, myUser5, myPass1, myPass2, myPass3, myPass4, myPass5 As String
    Dim u1, u2, u3, u4, u5, p1, p2, p3, p4, p5 As String
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Dim pchk As String = ""
        Try
            If Detest = 0 Then
                If UsernameTextBox.TextLength <> 0 Then
                    If UsernameTextBox.Text = u1 Or UsernameTextBox.Text = u2 Or UsernameTextBox.Text = u3 Or UsernameTextBox.Text = u4 Or UsernameTextBox.Text = u5 Then
                        If PasswordTextBox.TextLength <> 0 Then

                            Select Case UsernameTextBox.Text
                                Case u1
                                    If PasswordTextBox.Text = p1 Then
                                        If rmode = 0 Then
                                            Form1.Show()
                                        Else
                                            Form4.Show()
                                        End If
                                        Me.Hide()
                                    Else
                                        MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                                        PasswordTextBox.Clear()
                                        UsernameTextBox.Clear()
                                        UsernameTextBox.Focus()
                                    End If
                                Case u2
                                    If PasswordTextBox.Text = p2 Then
                                        If rmode = 0 Then
                                            Form1.Show()
                                        Else
                                            Form4.Show()
                                        End If
                                        Me.Hide()
                                    Else
                                        MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                                        PasswordTextBox.Clear()
                                        UsernameTextBox.Clear()
                                        UsernameTextBox.Focus()
                                    End If
                                Case u3
                                    If PasswordTextBox.Text = p3 Then
                                        If rmode = 0 Then
                                            Form1.Show()
                                        Else
                                            Form4.Show()
                                        End If
                                        Me.Hide()
                                    Else
                                        MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                                        PasswordTextBox.Clear()
                                        UsernameTextBox.Clear()
                                        UsernameTextBox.Focus()
                                    End If
                                Case u4
                                    If PasswordTextBox.Text = p4 Then
                                        If rmode = 0 Then
                                            Form1.Show()
                                        Else
                                            Form4.Show()
                                        End If
                                        Me.Hide()
                                    Else
                                        MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                                        PasswordTextBox.Clear()
                                        UsernameTextBox.Clear()
                                        UsernameTextBox.Focus()
                                    End If
                                Case u5
                                    If PasswordTextBox.Text = p5 Then
                                        If rmode = 0 Then
                                            Form1.Show()
                                        Else
                                            Form4.Show()
                                        End If
                                        Me.Hide()
                                    Else
                                        MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                                        PasswordTextBox.Clear()
                                        UsernameTextBox.Clear()
                                        UsernameTextBox.Focus()
                                    End If
                            End Select
                        Else
                            MsgBox("Invalid Password!!", MsgBoxStyle.Critical, "Invalid Password")
                            PasswordTextBox.Clear()
                            UsernameTextBox.Clear()
                            UsernameTextBox.Focus()
                        End If

                    Else
                        MsgBox("Invalid Username!!", MsgBoxStyle.Critical, "Invalid Username")
                        UsernameTextBox.Clear()
                        UsernameTextBox.Focus()
                    End If
                Else
                    MsgBox("Invalid Username!!", MsgBoxStyle.Critical, "Invalid Username")
                    UsernameTextBox.Focus()
                End If
            Else
                If rmode = 0 Then
                    Form1.Show()
                Else
                    Form4.Show()
                End If
                Me.Hide()
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        PasswordTextBox.Clear()
        UsernameTextBox.Clear()
        UsernameTextBox.Focus()
    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim file1, inpath As String
        Dim arrLines() As String
        Dim dd As Integer

        Dim password As String = "Gemalto123$" 'default password for cipher program
        Dim wrapper As New ClassLibrary1.Simple3Des(password)

        inpath = My.Application.Info.DirectoryPath
        file1 = inpath & "\" & "config.ini"

        Try
            If System.IO.File.Exists(file1) = True Then
                arrLines = File.ReadAllLines(file1)
                dd = arrLines.Length
            Else
                Console.WriteLine("config file Does Not Exist. Unable to run!!")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try

        Try
            For i = 0 To dd - 1
                If LSet(arrLines(i), 11) = "LabelPrint=" Then
                    lbPrint = Mid(arrLines(i), 12)
                ElseIf LSet(arrLines(i), 11) = "zipLocator=" Then
                    ziplocation = Mid(arrLines(i), 12)
                ElseIf LSet(arrLines(i), 11) = "LabelSizeX=" Then
                    lbx = Mid(arrLines(i), 12)
                ElseIf LSet(arrLines(i), 11) = "LabelSizeY=" Then
                    lby = Mid(arrLines(i), 12)
                ElseIf LSet(arrLines(i), 10) = "LabelShow=" Then
                    lbShow = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "LabelSave=" Then
                    lbSave = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln1X=" Then
                    lb1x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln2X=" Then
                    lb2x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln3X=" Then
                    lb3x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln4X=" Then
                    lb4x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln5X=" Then
                    lb5x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln6X=" Then
                    lb6x = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln1Y=" Then
                    lb1y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln2Y=" Then
                    lb2y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln3Y=" Then
                    lb3y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln4Y=" Then
                    lb4y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln5Y=" Then
                    lb5y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "Labelln6Y=" Then
                    lb6y = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "LabelBarX=" Then
                    lbBarX = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 10) = "LabelBarY=" Then
                    lbBarY = Mid(arrLines(i), 11)
                ElseIf LSet(arrLines(i), 8) = "runMode=" Then
                    rmode = Mid(arrLines(i), 9)
                ElseIf LSet(arrLines(i), 8) = "dbdebug=" Then
                    dbcheck = Mid(arrLines(i), 9)
                ElseIf LSet(arrLines(i), 7) = "MsgBox=" Then
                    mpbox = Mid(arrLines(i), 8)
                ElseIf LSet(arrLines(i), 7) = "p1logs=" Then
                    p1logs = Mid(arrLines(i), 8)
                ElseIf LSet(arrLines(i), 7) = "p2logs=" Then
                    p2logs = Mid(arrLines(i), 8)

                ElseIf LSet(arrLines(i), 6) = "Debug=" Then
                    Detest = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "User=" Then
                    myUser = Mid(arrLines(i), 6)
                ElseIf LSet(arrLines(i), 4) = "pwd=" Then
                    myPass = Mid(arrLines(i), 5)
                ElseIf LSet(arrLines(i), 6) = "User1=" Then
                    myUser1 = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "pwd1=" Then
                    myPass1 = Mid(arrLines(i), 6)
                ElseIf LSet(arrLines(i), 6) = "User2=" Then
                    myUser2 = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "pwd2=" Then
                    myPass2 = Mid(arrLines(i), 6)
                ElseIf LSet(arrLines(i), 6) = "User3=" Then
                    myUser3 = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "pwd3=" Then
                    myPass3 = Mid(arrLines(i), 6)
                ElseIf LSet(arrLines(i), 6) = "User4=" Then
                    myUser4 = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "pwd4=" Then
                    myPass4 = Mid(arrLines(i), 6)
                ElseIf LSet(arrLines(i), 6) = "User5=" Then
                    myUser5 = Mid(arrLines(i), 7)
                ElseIf LSet(arrLines(i), 5) = "pwd5=" Then
                    myPass5 = Mid(arrLines(i), 6)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & " Error!!", MsgBoxStyle.Critical, "Exception")
        End Try
        Try
            myOPass = wrapper.DecryptData(myPass)
            u1 = wrapper.DecryptData(myUser1)
            u2 = wrapper.DecryptData(myUser2)
            u3 = wrapper.DecryptData(myUser3)
            u4 = wrapper.DecryptData(myUser4)
            u5 = wrapper.DecryptData(myUser5)
            p1 = wrapper.DecryptData(myPass1)
            p2 = wrapper.DecryptData(myPass2)
            p3 = wrapper.DecryptData(myPass3)
            p4 = wrapper.DecryptData(myPass4)
            p5 = wrapper.DecryptData(myPass5)
        Catch ex As System.Security.Cryptography.CryptographicException
            MsgBox("The database password could not be decrypted with the password.")
        End Try

    End Sub
End Class
