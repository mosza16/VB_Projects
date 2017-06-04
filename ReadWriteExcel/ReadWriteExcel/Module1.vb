Imports Excel = Microsoft.Office.Interop.Excel

Module Module1
    Dim worksheet As Excel.Worksheet
    Dim workbook As Excel.Workbook
    Dim APP As New Excel.Application
    Dim xlUp = Excel.XlDirection.xlUp
    Dim xlToLeft = Excel.XlDirection.xlToLeft
    Dim xlValues = Excel.XlFindLookIn.xlValues,
        xlWhole = Excel.XlLookAt.xlWhole,
        xlByRows = Excel.XlSearchOrder.xlByRows,
        xlNext = Excel.XlSearchDirection.xlNext,
        xlPart = Excel.XlLookAt.xlPart

    Sub Main()
        Try
            workbook = APP.Workbooks.Open("D:\visual studio 2017\Projects\ReadWriteExcel\ReadWriteExcel\test3.xlsx")
            worksheet = workbook.Worksheets("sheet1")
            Dim LastRow As Long, LastCol As Integer
            With worksheet
                LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                Console.WriteLine("last row is {0}", LastRow)
                Console.WriteLine("last Cloumn is {0}", LastCol)
                Console.WriteLine("---------------------------------------")
                Dim lCount As Long = 1
                Dim rFoundCell = .Range("E1")
                Dim searchText = "N"
                Dim preventRow As Long = -1
                Console.WriteLine("********************************************")
                Console.WriteLine(" Strat Searching !!! '{0}' ", searchText)
                Console.WriteLine("********************************************")
                Console.WriteLine("---------------------------------------")
                While True
                    rFoundCell = find(worksheet, searchText, "E1", 5)
                    If rFoundCell IsNot Nothing Then
                        Dim foundRow As Long = rFoundCell.Row
                        Dim foundColumn As Long = rFoundCell.Column
                        If preventRow < foundRow Then
                            Dim updateTime As Date = Now()
                            Dim userID
                            Dim Password
                            For i As Long = foundRow To 1 Step -1
                                If Not .Cells(i, 2).value Is Nothing Then
                                    userID = .Cells(i, 2).value
                                    Password = .Cells(i, 3).value
                                    Exit For
                                End If
                            Next
                            rFoundCell.Value = "S"
                            .Cells(foundRow, 6).value = FormatDateTime(updateTime)
                            Console.WriteLine(" found!!! {0}", lCount)
                            Console.WriteLine(" found in row: {0}", foundRow)
                            Console.WriteLine(" found in column: {0}", foundColumn)
                            Console.WriteLine("USER ID IS {0}  AND PASSWORD IS {1}", userID, Password)
                            Console.WriteLine("---------------------------------------")
                            preventRow = foundRow
                            lCount += 1
                        Else
                            Console.WriteLine(" finish!!!")
                            Exit While
                        End If

                    Else
                        Console.WriteLine(" not found")
                        Console.WriteLine(" finish!!!")
                        Exit While
                    End If
                End While
            End With
            workbook.Save()
            workbook.Close()
            APP.Quit()
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try

    End Sub
    Sub test2()
        Try
            workbook = APP.Workbooks.Open("D:\visual studio 2017\Projects\ReadWriteExcel\ReadWriteExcel\test2.xlsx")
            worksheet = workbook.Worksheets("sheet1")
            Dim LastRow As Long, LastCol As Integer
            With worksheet
                LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                Console.WriteLine("last row is {0}", LastRow)
                Console.WriteLine("last Cloumn is {0}", LastCol)
                Console.WriteLine("---------------------------------------")
                Dim lCount As Long = 1
                Dim rFoundCell = .Range("C1")
                Dim searchText = "N"
                Dim preventRow As Long = -1
                Console.WriteLine("********************************************")
                Console.WriteLine(" Strat Searching !!! '{0}' ", searchText)
                Console.WriteLine("********************************************")
                Console.WriteLine("---------------------------------------")
                While True
                    rFoundCell = find(worksheet, searchText, "C1", 3)
                    If Not rFoundCell Is Nothing Then
                        Dim foundRow As Long = rFoundCell.Row
                        Dim foundColumn As Long = rFoundCell.Column
                        If preventRow < foundRow Then
                            Dim updateTime As Date = Now()
                            Dim userID = .Cells(foundRow, 1).value
                            Dim Password = .Cells(foundRow, 2).value
                            rFoundCell.Value = "S"
                            .Cells(foundRow, 4).value = FormatDateTime(updateTime)
                            Console.WriteLine(" found!!! {0}", lCount)
                            Console.WriteLine(" found in row: {0}", foundRow)
                            Console.WriteLine(" found in column: {0}", foundColumn)
                            Console.WriteLine("USER ID IS {0}  AND PASSWORD IS {1}", userID, Password)
                            Console.WriteLine("---------------------------------------")
                            preventRow = foundRow
                            lCount += 1
                        Else
                            Console.WriteLine(" finish!!!")
                            Exit While
                        End If

                    Else
                        Console.WriteLine(" not found")
                        Console.WriteLine(" finish!!!")
                        Exit While
                    End If
                End While
            End With
            workbook.Save()
            workbook.Close()
            APP.Quit()
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Sub
    Public Function find(worksheet, searchText, rangeExcel, ColumnIndex) As Excel.Range
        With worksheet
            Dim rFoundCell = .Range(rangeExcel)
            rFoundCell = .Columns(ColumnIndex).Find(What:=searchText, After:=rFoundCell,
                        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows,
                        SearchDirection:=xlNext, MatchCase:=False)
            Return rFoundCell
        End With
    End Function
End Module
