Imports Microsoft.Office.Interop
Imports System.IO
Imports System.IO.File

Public Class Form1

    Dim topval1 As Integer = 15
    Dim topval2 As Integer = 15

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim dynamicControl As New UserControl1
        With dynamicControl
            .Name = "Team1"
            .Text = "Test1"
            .Location = New Point(0, topval1)
        End With

        topval1 = topval1 + 90
        GroupBox1.Height = GroupBox1.Height + 90

        If GroupBox1.Height > GroupBox2.Height Then
            If Me.Height < Me.Height + GroupBox1.Height Then
                Me.Height = Me.Height + 90
            End If
        Else
            'do nothing
        End If

        GroupBox1.Controls.Add(dynamicControl)
        Dim op As Control
        For Each op In dynamicControl.Controls
            If op.GetType.ToString = "System.Windows.Forms.Button" Then
                AddHandler op.Click, AddressOf dynamicControl1_click
            End If
        Next

        'AddHandler DirectCast(dynamicControl, Control).MouseClick, AddressOf Global_MouseDown
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dynamicControl As New UserControl1
        With dynamicControl
            .Name = "Team2"
            .Text = "Test2"
            .Location = New Point(0, topval2)

        End With

        topval2 = topval2 + 90
        GroupBox2.Height = GroupBox2.Height + 90

        If GroupBox2.Height > GroupBox1.Height Then
            If Me.Height < Me.Height + GroupBox2.Height Then
                Me.Height = Me.Height + 90
            End If
        Else
            'do nothing
        End If
       
        GroupBox2.Controls.Add(dynamicControl)
        Dim op As Control
        For Each op In dynamicControl.Controls
            If op.GetType.ToString = "System.Windows.Forms.Button" Then
                AddHandler op.Click, AddressOf dynamicControl2_click
            End If
        Next

    End Sub

    Private Sub dynamicControl1_click(ByVal sender As Object, ByVal e As System.EventArgs)

        Label10.Text = 0
        Label27.Text = 0
        Label28.Text = 0
        Label29.Text = 0
        Label30.Text = 0
        Label31.Text = 0
        Label32.Text = 0

        Dim numerator1 As Integer = 0
        Dim denominator1 As Integer = 0
       
        For Each ctr As UserControl1 In GroupBox1.Controls
            Label10.Text = Label10.Text + Convert.ToInt32(ctr.points)
            Label27.Text = Label27.Text + ctr.Rebounds
            Label28.Text = Label28.Text + ctr.Assists
            Label29.Text = Label29.Text + ctr.Steals
            Label30.Text = Label30.Text + ctr.Blocks
            Label31.Text = Label31.Text + ctr.Turnovers
            Label32.Text = Label32.Text + ctr.Fouls

            numerator1 = numerator1 + (ctr.FGM + ctr.TM)
            denominator1 = denominator1 + (ctr.FGA + ctr.TA)
            Label33.Text = FormatNumber((numerator1 / denominator1 * 100), 2) & "%"

        Next

    End Sub

    Private Sub dynamicControl2_click(ByVal sender As Object, ByVal e As System.EventArgs)

        Label20.Text = 0
        Label21.Text = 0
        Label22.Text = 0
        Label23.Text = 0
        Label24.Text = 0
        Label25.Text = 0
        Label26.Text = 0

        Dim numerator2 As Integer = 0
        Dim denominator2 As Integer = 0

        For Each ctr As UserControl1 In GroupBox2.Controls
            Label20.Text = Label20.Text + Convert.ToInt32(ctr.points)
            Label21.Text = Label21.Text + ctr.Rebounds
            Label22.Text = Label22.Text + ctr.Assists
            Label23.Text = Label23.Text + ctr.Steals
            Label24.Text = Label24.Text + ctr.Blocks
            Label25.Text = Label25.Text + ctr.Turnovers
            Label26.Text = Label26.Text + ctr.Fouls

            numerator2 = numerator2 + (ctr.FGM + ctr.TM)
            denominator2 = denominator2 + (ctr.FGA + ctr.TA)
            Label34.Text = FormatNumber((numerator2 / denominator2 * 100), 2) & "%"

        Next

    End Sub

    Private Sub SaveToExcelToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToExcelToolStripMenuItem3.Click

        Dim App As New Excel.Application
        Dim WB As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        Dim xlrange As Excel.Range
        Dim misValue As Object = System.Reflection.Missing.Value

        App = New Excel.ApplicationClass
        WB = App.Workbooks.Add(misValue)

        xlSheet = WB.Sheets("Sheet1")

        'Try
        '    WB = App.Workbooks.Open("C:\BballStats.xls")
        'Catch e1 As Exception
        '    Exit Sub
        'End Try

        xlSheet.Cells(1, 1) = "Player"
        xlSheet.Cells(1, 2) = "PTS"
        xlSheet.Cells(1, 3) = "FGM"
        xlSheet.Cells(1, 4) = "FGA"
        xlSheet.Cells(1, 5) = "3PtM"
        xlSheet.Cells(1, 6) = "3PtA"
        xlSheet.Cells(1, 7) = "FTM"
        xlSheet.Cells(1, 8) = "FTA"
        xlSheet.Cells(1, 9) = "REB"
        xlSheet.Cells(1, 10) = "AST"
        xlSheet.Cells(1, 11) = "STL"
        xlSheet.Cells(1, 12) = "BLK"
        xlSheet.Cells(1, 13) = "TO"
        xlSheet.Cells(1, 14) = "PF"

        Dim IntRowNumber As Integer
        IntRowNumber = 2

        For Each ctr As UserControl1 In GroupBox1.Controls
            xlrange = xlSheet.Range("A" & CStr(IntRowNumber))
            xlrange.Value = ctr.player
            xlrange = xlSheet.Range("B" & CStr(IntRowNumber))
            xlrange.Value = ctr.points
            xlrange = xlSheet.Range("C" & CStr(IntRowNumber))
            xlrange.Value = ctr.FGM
            xlrange = xlSheet.Range("D" & CStr(IntRowNumber))
            xlrange.Value = ctr.FGA
            xlrange = xlSheet.Range("E" & CStr(IntRowNumber))
            xlrange.Value = ctr.TM
            xlrange = xlSheet.Range("F" & CStr(IntRowNumber))
            xlrange.Value = ctr.TA
            xlrange = xlSheet.Range("G" & CStr(IntRowNumber))
            xlrange.Value = ctr.FTM
            xlrange = xlSheet.Range("H" & CStr(IntRowNumber))
            xlrange.Value = ctr.FTA
            xlrange = xlSheet.Range("I" & CStr(IntRowNumber))
            xlrange.Value = ctr.Rebounds
            xlrange = xlSheet.Range("J" & CStr(IntRowNumber))
            xlrange.Value = ctr.Assists
            xlrange = xlSheet.Range("K" & CStr(IntRowNumber))
            xlrange.Value = ctr.Steals
            xlrange = xlSheet.Range("L" & CStr(IntRowNumber))
            xlrange.Value = ctr.Blocks
            xlrange = xlSheet.Range("M" & CStr(IntRowNumber))
            xlrange.Value = ctr.Turnovers
            xlrange = xlSheet.Range("N" & CStr(IntRowNumber))
            xlrange.Value = ctr.Fouls

            IntRowNumber = IntRowNumber + 1

        Next

        IntRowNumber = IntRowNumber + 1

        For Each ctr As UserControl1 In GroupBox2.Controls
            xlrange = xlSheet.Range("A" & CStr(IntRowNumber))
            xlrange.Value = ctr.player
            xlrange = xlSheet.Range("B" & CStr(IntRowNumber))
            xlrange.Value = ctr.points
            xlrange = xlSheet.Range("C" & CStr(IntRowNumber))
            xlrange.Value = ctr.FGM
            xlrange = xlSheet.Range("D" & CStr(IntRowNumber))
            xlrange.Value = ctr.FGA
            xlrange = xlSheet.Range("E" & CStr(IntRowNumber))
            xlrange.Value = ctr.TM
            xlrange = xlSheet.Range("F" & CStr(IntRowNumber))
            xlrange.Value = ctr.TA
            xlrange = xlSheet.Range("G" & CStr(IntRowNumber))
            xlrange.Value = ctr.FTM
            xlrange = xlSheet.Range("H" & CStr(IntRowNumber))
            xlrange.Value = ctr.FTA
            xlrange = xlSheet.Range("I" & CStr(IntRowNumber))
            xlrange.Value = ctr.Rebounds
            xlrange = xlSheet.Range("J" & CStr(IntRowNumber))
            xlrange.Value = ctr.Assists
            xlrange = xlSheet.Range("K" & CStr(IntRowNumber))
            xlrange.Value = ctr.Steals
            xlrange = xlSheet.Range("L" & CStr(IntRowNumber))
            xlrange.Value = ctr.Blocks
            xlrange = xlSheet.Range("M" & CStr(IntRowNumber))
            xlrange.Value = ctr.Turnovers
            xlrange = xlSheet.Range("N" & CStr(IntRowNumber))
            xlrange.Value = ctr.Fouls

            IntRowNumber = IntRowNumber + 1

            'Dim str_temp As String
            'str_temp = ""

            'str_temp = str_temp & ctr.player & ctr.points & vbCrLf

            'If ctr.GetType.ToString = "System.Windows.Forms.TextBox" Then
            'End If

        Next

        Dim time As DateTime = DateTime.Now
        Dim format As String = "Mdyy_HHmm"

        Try

            Dim saveFileDialog As New SaveFileDialog
            saveFileDialog.Filter = "Excel File|*.xlsx"
            saveFileDialog.Title = "Save an Excel File"
            saveFileDialog.FileName = "Stats" & time.ToString(format)
            saveFileDialog.ShowDialog()
            If saveFileDialog.FileName <> "" Then
                WB.SaveAs(saveFileDialog.FileName)
            End If

        Catch ex As Exception

            WB.Close()
            App.Quit()

            releaseObject(WB)
            releaseObject(App)
            releaseObject(xlSheet)
            releaseObject(xlrange)

            MsgBox(ex)
        End Try

        'WB.SaveAs("C:\Stats_" & time.ToString(format) & ".Xls")

        WB.Close()
        App.Quit()

        releaseObject(WB)
        releaseObject(App)
        releaseObject(xlSheet)
        releaseObject(xlrange)

        'MsgBox("Excel file created , you can find the file c:\Stats_" & time.ToString(format))
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub



    Private Sub RestartToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestartToolStripMenuItem3.Click
        Dim result = MessageBox.Show("Restart?", "Are You sure?", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            Exit Sub
        ElseIf result = DialogResult.OK Then
            Application.Restart()
        End If
    End Sub

    Private Sub AboutBBallStatsToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutBBallStatsToolStripMenuItem2.Click
        Dim ButtonOpen
        ButtonOpen = About
        ButtonOpen.Show()
    End Sub

    Private Sub SaveToCSVToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToCSVToolStripMenuItem.Click


        Try

            Dim time As DateTime = DateTime.Now
            Dim format As String = "Mdyy_HHmm"

            Dim saveFileDialog As New SaveFileDialog
            saveFileDialog.Filter = "Comma Separated Values|*.csv"
            saveFileDialog.Title = "Save a CSV File"
            saveFileDialog.FileName = "Stats" & time.ToString(format)
            saveFileDialog.ShowDialog()

            Dim filename As String = saveFileDialog.FileName

            Dim objWriter As New System.IO.StreamWriter(filename)
            objWriter.Write("Player,")
            objWriter.Write("PTS,")
            objWriter.Write("FGM,")
            objWriter.Write("FGA,")
            objWriter.Write("3PtM,")
            objWriter.Write("3PtA,")
            objWriter.Write("FTM,")
            objWriter.Write("FTA,")
            objWriter.Write("REB,")
            objWriter.Write("AST,")
            objWriter.Write("STL,")
            objWriter.Write("BLK,")
            objWriter.Write("TO,")
            objWriter.Write("PF")
            objWriter.WriteLine()


            For Each ctr As UserControl1 In GroupBox1.Controls
                objWriter.Write(ctr.player.ToString & ",")
                objWriter.Write(ctr.points.ToString & ",")
                objWriter.Write(ctr.FGM.ToString & ",")
                objWriter.Write(ctr.FGA.ToString & ",")
                objWriter.Write(ctr.TM.ToString & ",")
                objWriter.Write(ctr.TA.ToString & ",")
                objWriter.Write(ctr.FTM.ToString & ",")
                objWriter.Write(ctr.FTA.ToString & ",")
                objWriter.Write(ctr.Rebounds.ToString & ",")
                objWriter.Write(ctr.Assists.ToString & ",")
                objWriter.Write(ctr.Steals.ToString & ",")
                objWriter.Write(ctr.Blocks.ToString & ",")
                objWriter.Write(ctr.Turnovers.ToString & ",")
                objWriter.Write(ctr.Fouls.ToString)
                objWriter.WriteLine()
            Next

            objWriter.WriteLine()

            For Each ctr As UserControl1 In GroupBox2.Controls
                objWriter.Write(ctr.player.ToString & ",")
                objWriter.Write(ctr.points.ToString & ",")
                objWriter.Write(ctr.FGM.ToString & ",")
                objWriter.Write(ctr.FGA.ToString & ",")
                objWriter.Write(ctr.TM.ToString & ",")
                objWriter.Write(ctr.TA.ToString & ",")
                objWriter.Write(ctr.FTM.ToString & ",")
                objWriter.Write(ctr.FTA.ToString & ",")
                objWriter.Write(ctr.Rebounds.ToString & ",")
                objWriter.Write(ctr.Assists.ToString & ",")
                objWriter.Write(ctr.Steals.ToString & ",")
                objWriter.Write(ctr.Blocks.ToString & ",")
                objWriter.Write(ctr.Turnovers.ToString & ",")
                objWriter.Write(ctr.Fouls.ToString)
                objWriter.WriteLine()
            Next
            objWriter.Close()

        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub
End Class
