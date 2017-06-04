Imports Excel = Microsoft.Office.Interop.Excel
Public Class MyExcel
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
    Public Function find(worksheet, searchText, rangeExcel, ColumnIndex) As Excel.Range
        With worksheet
            Dim rFoundCell = .Range(rangeExcel)
            rFoundCell = .Columns(ColumnIndex).Find(What:=searchText, After:=rFoundCell,
                        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows,
                        SearchDirection:=xlNext, MatchCase:=False)
            Return rFoundCell
        End With
    End Function
End Class
