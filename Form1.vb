Imports Microsoft.Office.Interop
Imports System.Linq
Imports System.Globalization

Public Class Form1
    Private Sub btnProcessFile_Click(sender As Object, e As EventArgs) Handles btnProcessFile.Click
        Me.Cursor = Cursors.WaitCursor
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            result = FolderBrowserDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                lblFileName.Text = OpenFileDialog1.FileName & " -> " & FolderBrowserDialog1.SelectedPath
                Dim dlgOptions As New Options
                result = dlgOptions.ShowDialog()
                If result = Windows.Forms.DialogResult.OK Then
                    If Not String.IsNullOrEmpty(dlgOptions.txtCurrentSemester.Text) Then
                        lblFileName.Text = lblFileName.Text & vbCrLf & "Semester: " & dlgOptions.txtCurrentSemester.Text
                    End If
                    btnCancel.Visible = True
                    BackgroundWorker1.RunWorkerAsync({OpenFileDialog1.FileName, FolderBrowserDialog1.SelectedPath, dlgOptions.txtCurrentSemester.Text, dlgOptions.txtFirstScoreCol.Text, dlgOptions.txtLastScoreCol.Text})
                End If
            End If
        End If
    End Sub

    Enum ColumnIDs
        FirstName = 1
        LastName = 2
        ClassStanding = 4
        Emphasis = 5
        Semester = 6
    End Enum

    Private Function ConvertExcelColumnToInteger(colLetter As String) As Integer
        Dim intVal As Integer = 0
        Dim exp As Integer = 0
        Dim rv As Integer = 0
        For Each ch As Char In colLetter.ToCharArray().Reverse()
            If exp = 0 Then
                intVal = Microsoft.VisualBasic.AscW(ch) - Microsoft.VisualBasic.AscW("A"c)
            Else
                intVal = Microsoft.VisualBasic.AscW(ch) - Microsoft.VisualBasic.AscW("A"c) + 1
            End If
            rv += intVal * (Math.Pow(26, exp))
            exp += 1
        Next
        Return rv
    End Function

    Private Function GetStringCellValue(worksheet As Excel.Worksheet, row As Integer, column As Integer) As String
        Dim rc As String = GetCellValue(Of String)(worksheet, row, column)
        If Not String.IsNullOrEmpty(rc) Then
            rc = rc.Replace(Microsoft.VisualBasic.ChrW(&HA0), " "c)
        End If
        Return rc
    End Function

    Private Function GetCellValue(Of T)(worksheet As Excel.Worksheet, row As Integer, column As Integer) As T
        Return CType(CType(worksheet.UsedRange.Cells(row, column), Excel.Range).Value, T)
    End Function

    Class AssessmentScore
        Public StudentID As String
        Public Semester As String
        Public SemesterSort As String
        Public ClassStanding As String
        Public Emphasis As String
        Public Name As String
        Public Score As Decimal
    End Class

    Private Function GetSortableSemesterValue(semester As String) As String
        Dim split As String()
        If semester.Contains(" ") Then
            split = semester.Split(" ")
        Else
            Throw New Exception("Unexpected semester identifier '" & semester & "'.")
        End If
        Return split(1) + split(0).ToUpper(CultureInfo.CurrentCulture).Replace("FALL", "2").Replace("SPRING", "1")
    End Function

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim args As String() = CType(e.Argument, String())
        Dim sourceFile As String = args(0)
        Dim outputFolder As String = args(1)
        Dim currentSemester As String = args(2)
        If Not String.IsNullOrEmpty(currentSemester) Then
            currentSemester = GetSortableSemesterValue(currentSemester)
        End If
        Dim firstScoreCol As Integer = ConvertExcelColumnToInteger(args(3))
        Dim lastScoreCol As Integer = ConvertExcelColumnToInteger(args(4))
        Dim excel As Excel.Application = Nothing
        Dim workbook As Excel.Workbook = Nothing
        Dim worksheet As Excel.Worksheet = Nothing
        Dim word As Word.Application = Nothing
        Try
            excel = New Excel.Application()
            workbook = excel.Workbooks.Open(sourceFile)
            worksheet = workbook.Worksheets(1)

            ' Sanity check
            Dim columnTitle As String = GetStringCellValue(worksheet, 1, ColumnIDs.FirstName)
            If Not "First Name".Equals(columnTitle, StringComparison.CurrentCultureIgnoreCase) Then
                Throw New Exception("The selected file does not fit the expected format.")
            End If

            Dim scores As New List(Of AssessmentScore)

            Log("Reading scores...", True)
            Dim row As Integer = 2
            Dim studentName As String = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
            Do While Not String.IsNullOrEmpty(studentName)
                If BackgroundWorker1.CancellationPending Then
                    e.Cancel = True
                    Return
                End If
                Log("Reading scores... " & row, False)
                For i As Integer = firstScoreCol To lastScoreCol
                    Dim scoreStr As String = GetStringCellValue(worksheet, row, i)
                    If Not String.IsNullOrEmpty(scoreStr) Then
                        Dim score As Decimal = Decimal.Parse(scoreStr)
                        scores.Add(New AssessmentScore() With {
                                        .StudentID = studentName,
                                        .Semester = GetStringCellValue(worksheet, row, ColumnIDs.Semester),
                                        .SemesterSort = GetSortableSemesterValue(GetStringCellValue(worksheet, row, ColumnIDs.Semester)),
                                        .ClassStanding = GetStringCellValue(worksheet, row, ColumnIDs.ClassStanding),
                                        .Emphasis = GetStringCellValue(worksheet, row, ColumnIDs.Emphasis),
                                        .Name = GetStringCellValue(worksheet, 1, i),
                                        .Score = score
                                    })
                    End If
                Next

                row += 1
                studentName = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
                ' Tolerate a few empty rows
                If String.IsNullOrEmpty(studentName) Then
                    row += 1
                    studentName = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
                    If String.IsNullOrEmpty(studentName) Then
                        row += 1
                        studentName = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
                        If String.IsNullOrEmpty(studentName) Then
                            row += 1
                            studentName = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
                            If String.IsNullOrEmpty(studentName) Then
                                row += 1
                                studentName = (GetStringCellValue(worksheet, row, ColumnIDs.FirstName) & " " & GetStringCellValue(worksheet, row, ColumnIDs.LastName)).Trim()
                            End If
                        End If
                    End If
                End If
            Loop

            If Not String.IsNullOrEmpty(currentSemester) Then
                Dim studentNames As ILookup(Of String, String) = (From score In scores
                                                          Where score.SemesterSort = currentSemester
                                                          Select score.StudentID).ToLookup(Function(studentID As String)
                                                                                               Return studentID
                                                                                           End Function)

                scores = (From score In scores
                          Where studentNames.Contains(score.StudentID)
                          Select score).ToList()
            End If

            If BackgroundWorker1.CancellationPending Then
                e.Cancel = True
                Return
            End If
            Log("Calculating averages...", True)
            ' Get the roll up data
            Dim averages = From score In scores
                           Group By sid = score.StudentID, sem = score.SemesterSort, standing = score.ClassStanding, name = score.Name
                           Into averageScore = Average(score.Score), list = Group
                           Order By sem
            If BackgroundWorker1.CancellationPending Then
                e.Cancel = True
                Return
            End If
            Dim studentAverages = From avg In averages
                                  Group By sid = avg.sid
                                  Into list = Group
                                  Order By sid.Split(" "c)(1)

            If BackgroundWorker1.CancellationPending Then
                e.Cancel = True
                Return
            End If
            Log("Writing reports...", True)
            word = New Word.Application
            For Each savg In studentAverages
                If BackgroundWorker1.CancellationPending Then
                    e.Cancel = True
                    Return
                End If
                Dim ex = savg.list.First().list.First()
                Log(vbTab & ex.StudentID, True)

                ' Write the report document
                Dim doc As Word.Document = word.Documents.Add()
                Dim para As Word.Paragraph = doc.Content.Paragraphs.Add()
                para.Range.Text = "Assessment Scores For " & ex.StudentID
                para.Range.Font.Bold = True
                para.Range.InsertParagraphAfter()

                Dim table As Word.Table = doc.Tables.Add(doc.Bookmarks.Item("\endofdoc").Range, savg.list.Count() + 1, 4)
                table.Cell(1, 1).Range.Text = "Semester"
                table.Cell(1, 1).Range.Font.Bold = True
                table.Cell(1, 2).Range.Text = "Assessment"
                table.Cell(1, 2).Range.Font.Bold = True
                table.Cell(1, 3).Range.Text = "Your Score"
                table.Cell(1, 3).Range.Font.Bold = True
                table.Cell(1, 4).Range.Text = "Average Score"
                table.Cell(1, 4).Range.Font.Bold = True

                row = 2
                For Each avg In savg.list
                    ' Get the comparative average score for students in the same group
                    Dim compQ = From score In scores
                                Where score.SemesterSort = avg.sem AndAlso score.ClassStanding = avg.standing AndAlso score.Emphasis = ex.Emphasis AndAlso score.Name = avg.name
                                Select score.Score
                    table.Cell(row, 1).Range.Text = avg.list.First().Semester
                    table.Cell(row, 1).Range.Font.Bold = False
                    table.Cell(row, 2).Range.Text = avg.name
                    table.Cell(row, 2).Range.Font.Bold = False
                    table.Cell(row, 3).Range.Text = avg.averageScore.ToString("0.##")
                    table.Cell(row, 3).Range.Font.Bold = False
                    If compQ.Count() > 0 Then
                        Dim comp As Decimal = compQ.Average()
                        table.Cell(row, 4).Range.Text = comp.ToString("0.##")
                    Else
                        table.Cell(row, 4).Range.Text = "N/A"
                    End If
                    table.Cell(row, 4).Range.Font.Bold = False
                    row += 1
                Next
                table = Nothing
                para = Nothing

                doc.SaveAs2(System.IO.Path.Combine(outputFolder, ex.StudentID & " Assessment Scores.docx"), AddToRecentFiles:=False)
                doc.Close()
                doc = Nothing
            Next
            Log("Done!", True)
            Dim MyProcess As New Process()
            MyProcess.StartInfo.FileName = "explorer.exe"
            MyProcess.StartInfo.Arguments = outputFolder
            MyProcess.Start()
            MyProcess.WaitForExit()
            MyProcess.Close()
            MyProcess.Dispose()
        Finally
            If Not workbook Is Nothing Then
                workbook.Close()
            End If
            word = Nothing
            worksheet = Nothing
            workbook = Nothing
            excel = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If Not e.Error Is Nothing Then
            MessageDialog.MessageLabel.Text = e.Error.Message
            MessageDialog.ShowDialog()
        End If
        Me.Cursor = Cursors.Default
        btnCancel.Visible = False
    End Sub

    Private Sub Log(message As String, append As Boolean)
        BeginInvoke(New MethodInvoker(Sub()
                                          If append Then
                                              txtLog.Text &= message & vbCrLf
                                          Else
                                              txtLog.Text = message & vbCrLf
                                          End If
                                          txtLog.SelectionStart = txtLog.TextLength
                                          txtLog.ScrollToCaret()
                                      End Sub))
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        BackgroundWorker1.CancelAsync()
    End Sub
End Class
