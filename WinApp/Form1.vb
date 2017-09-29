Imports System.Threading
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports Excel = Microsoft.Office.Interop.Excel

Public Structure Vote
    Public verifyCode As String
    Public number As String
    Public time As String
    Public note As String

End Structure
Public Class Form1
    Delegate Sub MyDelegate(ByVal text As String)
    Delegate Sub ExcelDeligate()
    Public processThread As Thread
    Public ending As Boolean = False
    Public keys As List(Of String) = New List(Of String)
    Public keysOriginal As List(Of String) = New List(Of String)
    Public numbers As List(Of String) = New List(Of String)
    Public startText As String = "INITVAL"
    Public players() As List(Of Vote) = {New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote)}
    Public lastVote As Vote
    Public SyncObj = New Object
    Public receive As Boolean = False
    Public DEBUG As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not DEBUG Then
            On Error Resume Next
        End If
        If Not SerialPort1.IsOpen Then
            SerialPort1.Open()
            Dim t As Thread
            t = New Thread(AddressOf processSerial)
            t.Start()
            Button1.Text = "Close"
        Else
            If Not ComboBox1.SelectedItem Is Nothing Then
                If ComboBox1.SelectedItem.ToString.ToUpper.Trim.StartsWith("COM") Then
                    SerialPort1.PortName = ComboBox1.SelectedItem.ToString.ToUpper.Trim
                Else

                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            ending = True
            Button1.Text = "Open"
        End If

    End Sub
    Function isInKys(ByVal key As String, ByVal delete As Boolean) As Boolean
        If Not DEBUG Then
            On Error Resume Next
        End If
        If keys.Contains(key) Then
            If delete Then
                keys.Remove(key)
            End If
            Return True
        Else
            Return False
        End If
    End Function
    Sub fileParser()
        If Not DEBUG Then
            On Error Resume Next
        End If
        Using ioReader As New TextFieldParser("keys.csv")
            ioReader.TextFieldType = FileIO.FieldType.Delimited
            ioReader.SetDelimiters(",")
            While Not ioReader.EndOfData
                Dim X = ioReader.ReadFields()

                If Not X Is Nothing Then
                    For Each i As String In X
                        keys.Add(i.Trim())
                        keysOriginal.Add(i.Trim())
                    Next
                End If

            End While
        End Using
    End Sub
    Sub ExcelWrite()
        If Not DEBUG Then
            On Error Resume Next
        End If
        Dim excel_app As New Excel.ApplicationClass()

        ' Make Excel visible (optional).
        excel_app.Visible = True
        Dim strPath As String = System.IO.Path.GetDirectoryName( _
    System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        ' Open the workbook.
        Dim workbook As Excel.Workbook = excel_app.Workbooks.Open(strPath & "\stat.xlsx")
        Dim sA As Excel.Worksheet = workbook.Sheets("A")
        Dim sB As Excel.Worksheet = workbook.Sheets("B")
        Dim sC As Excel.Worksheet = workbook.Sheets("C")
        Dim sD As Excel.Worksheet = workbook.Sheets("D")
        Dim sE As Excel.Worksheet = workbook.Sheets("E")
        Dim sF As Excel.Worksheet = workbook.Sheets("F")
        Dim sheets As Excel.Worksheet() = {sA, sB, sC, sD, sE, sF}
        Dim lastRidx As Integer() = {-1, -1, -1, -1, -1, -1}
        While True
            Dim tempP() As List(Of Vote) = players.Clone()
            For i As Integer = 0 To tempP.Length - 1
                Dim votes = tempP(i)
                Dim sheet = sheets(i)
                Dim idx = lastRidx(i)
                If votes.Count > idx + 1 Then
                    lastRidx(i) = lastRidx(i) + 1
                Else
                    Continue For
                End If
                idx = lastRidx(i)
                Dim v = votes(idx)

                Dim last As Integer
                For j As Integer = 4 To sheet.Columns.Count
                    Dim c As Excel.Range
                    c = sheet.Cells(j, "A")
                    If Trim(c.Value) = "" Then
                        last = j
                        Exit For
                    End If
                Next
                sheet.Cells(last, "A") = v.number
                sheet.Cells(last, "B") = v.verifyCode
                sheet.Cells(last, "C") = v.time
                If Not v.note Is Nothing Then
                    sheet.Cells(last, "D") = v.note
                End If
            Next
            Thread.Sleep(1000)

        End While
    End Sub
    Sub UI(ByVal text As String)
        If Not DEBUG Then
            On Error Resume Next
        End If
        TextBox1.Text = text

        If CheckBox2.Checked Then
            TextBox3.AppendText(vbNewLine & text)
        Else
            TextBox3.Text = TextBox3.Text & vbNewLine & text
        End If
        Using sw As StreamWriter = File.AppendText("log.txt")
            sw.WriteLine(text)
        End Using
    End Sub
    Private Sub SerialPort1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        ' Log(SerialPort1.ReadLine)

    End Sub
    Sub processSerial()
        If Not DEBUG Then
            On Error Resume Next
        End If
        While True
            Dim d = New MyDelegate(AddressOf UI)
            Dim data = ""
            If Not ending Then
                data = SerialPort1.ReadTo(vbNewLine)
                If Not data = "" Then
                    TextBox1.BeginInvoke(d, data)
                    processLine(data)

                End If
            Else
                SerialPort1.Close()
                Thread.Sleep(1000)
                ending = False
                Thread.CurrentThread.Abort()
            End If
        End While
    End Sub
    Sub startExcel()
        If Not DEBUG Then
            On Error Resume Next
        End If
        Dim t As Thread
        t = New Thread(AddressOf ExcelWrite)
        t.Start()
    End Sub
    Sub addVote(player As Integer, vote As Vote)
        If Not DEBUG Then
            On Error Resume Next
        End If
        players(player).Add(vote)
    End Sub
    Sub processLine(text As String)
        If Not DEBUG Then
            On Error Resume Next
        End If
        Dim t = text.Trim().ToUpper()
        If t.StartsWith("+CMGL") Then
            Dim x = t.Split(""",""")
            If x.Length > 5 Then
                Dim v = New Vote()
                v.number = x(3)
                v.time = x(5)
                lastVote = v
            End If
        ElseIf t.StartsWith(startText.Trim().ToUpper & " ") Then
            Dim arr = t.Split(" ")

            If arr.Length > 2 Then
                Dim player = arr(1).ToUpper()
                Dim code = arr(2)
                lastVote.verifyCode = code
                If isInKys(code, True) Then
                    Select Case player
                        Case "A"
                            My.Settings.A = My.Settings.A + 1
                            addVote(0, lastVote)
                        Case "B"
                            My.Settings.B = My.Settings.B + 1
                            addVote(1, lastVote)
                        Case "C"
                            My.Settings.C = My.Settings.C + 1
                            addVote(2, lastVote)
                        Case "D"
                            My.Settings.D = My.Settings.D + 1
                            addVote(3, lastVote)
                        Case "E"
                            My.Settings.E = My.Settings.E + 1
                            addVote(4, lastVote)
                        Case Else
                            My.Settings.F = My.Settings.F + 1
                            lastVote.note = t & " :Unknown vote"
                            addVote(5, lastVote)
                    End Select
                    My.Settings.Save()
                Else
                    My.Settings.F = My.Settings.F + 1
                    lastVote.note = t & " :Verification failed"
                    addVote(5, lastVote)
                End If
            Else
                My.Settings.F = My.Settings.F + 1
                lastVote.note = t & " :Wrong pattern"
                addVote(5, lastVote)

            End If


        End If


    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        If CheckBox1.Checked Then
            TextBox3.Height = 200
            TextBox3.Top = TextBox1.Top - 200
        Else
            TextBox3.Height = 0
            TextBox3.Top = TextBox1.Top
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If Not DEBUG Then
            On Error Resume Next
        End If
        If e.KeyCode = 13 Then
            PortWrite(TextBox2.Text)
            e.Handled = True
        End If
    End Sub
    Sub PortWrite(cmd)
        If Not DEBUG Then
            On Error Resume Next
        End If
        If SerialPort1.IsOpen Then
            SerialPort1.WriteLine(cmd & vbCr)
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        If CheckBox3.Checked Then
            PortWrite("AT^CURC=1")
        Else
            PortWrite("AT^CURC=0")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        If CheckBox4.Checked Then
            PortWrite("AT+CMGF=1")
        Else
            PortWrite("AT+CMGF=0")
        End If
    End Sub
    Sub deleteAllReadMsg()
        PortWrite("AT+CMGL=""ALL"";+CMGD=0,3")

    End Sub
    Sub viewAllMsgs()
        PortWrite("AT+CMGL=""ALL""")
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not DEBUG Then
            On Error Resume Next
        End If
        If Button2.Text.ToLower.Contains("start") Then
            receive = True
            PortWrite("AT+CMGF=1")
            viewAllMsgs()
            Button2.Text = "Stop recieving"
            Label8.Text = "Reading 15 msgs per " & TrackBar1.Value & "ms."
            TextBox4.Enabled = False
        Else
            receive = False
            Button2.Text = "Start recieving"
            Label8.Text = "Stopped."
            TextBox4.Enabled = True
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not DEBUG Then
            On Error Resume Next
        End If
        fileParser()
        For Each sp As String In My.Computer.Ports.SerialPortNames
            ComboBox1.Items.Add(sp)
        Next
        ComboBox1.SelectedItem = My.Settings.port
        Timer1.Interval = TrackBar1.Value
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        startText = TextBox4.Text.Trim()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not DEBUG Then
            On Error Resume Next
        End If
        startExcel()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        Label1.Text = "Votes for A, " & My.Settings.A
        Label2.Text = "Votes for B, " & My.Settings.B
        Label3.Text = "Votes for C, " & My.Settings.C
        Label4.Text = "Votes for D, " & My.Settings.D
        Label5.Text = "Votes for E, " & My.Settings.E
        Label6.Text = "Inval votes, " & My.Settings.F
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not DEBUG Then
            On Error Resume Next
        End If
        Dim response = InputBox("Enter password", "Clearing votes", "1234")
        If Val(response) = Minute(Now) Then
            For Each v As Vote In players.Clone
                If Not keys.Contains(v.verifyCode) Then
                    keys.Add(v.verifyCode)
                End If
            Next
            players = {New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote), New List(Of Vote)}
            My.Settings.A = 0
            My.Settings.B = 0
            My.Settings.C = 0
            My.Settings.D = 0
            My.Settings.E = 0
            My.Settings.F = 0
            MsgBox("Successfully cleared!", MsgBoxStyle.Information)
        Else
            MsgBox("Invalid password!", MsgBoxStyle.Exclamation)
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not DEBUG Then
            On Error Resume Next
        End If
        My.Settings.port = ComboBox1.SelectedItem.ToString.ToUpper
        My.Settings.Save()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Not DEBUG Then
            On Error Resume Next
        End If
        If receive Then
            deleteAllReadMsg()
        End If

    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        If Not DEBUG Then
            On Error Resume Next
        End If
        Timer1.Interval = TrackBar1.Value
        Label8.Text = "Reading 15 msgs per " & TrackBar1.Value & "ms."
    End Sub

    Private Sub RegrantToken() Handles Button5.Click
        If Not DEBUG Then
            On Error Resume Next
        End If

        Dim t = InputBox("Enter the token you want to regrant", , "")
        If t.Length > 4 Then
            If keysOriginal.Contains(t) Then
                If keys.Contains(t) Then
                    MsgBox(t & " is not used yet!", MsgBoxStyle.Exclamation)
                Else
                    keys.Add(t)
                    MsgBox(t & " is successfully added to keys.", MsgBoxStyle.Information)
                End If
            Else
                MsgBox(t & " is not valid. Please check in your token list", MsgBoxStyle.Exclamation)
            End If


        Else
            If Not t = "" Then
                MsgBox("Invalid length of key", MsgBoxStyle.Exclamation)
            End If
        End If

    End Sub
End Class
