Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Data

Module Module2
    Sub CreateTable()
        Dim fileTest As String = "\\slfs01\shared\prasinos\ppexternal\downloads\ScrapReportpI.xlsx"
        If File.Exists(fileTest) Then
            File.Delete(fileTest) ' oh, file is still open
        End If

        Dim oExcel As Object
        oExcel = CreateObject("Excel.Application")
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oBook = oExcel.Workbooks.Add
        oSheet = oExcel.Worksheets(1)

        oSheet.Name = "Report"
        oSheet.Range("A1").Value = "First Name"
        oSheet.Range("B1").Value = "Year"
        oSheet.Range("C1").Value = "Salary"

        oSheet.Range("A2").Value = "Frank"
        oSheet.Range("B2").Value = "2012"
        oSheet.Range("C2").Value = "30000"

        oSheet.Range("A3").Value = "Frank"
        oSheet.Range("B3").Value = "2011"
        oSheet.Range("C3").Value = "25000"

        oSheet.Range("A4").Value = "Ann"
        oSheet.Range("B4").Value = "2011"
        oSheet.Range("C4").Value = "55000"

        oSheet.Range("A5").Value = "Ann"
        oSheet.Range("B5").Value = "2012"
        oSheet.Range("C5").Value = "35000"

        oSheet.Range("A6").Value = "Ann"
        oSheet.Range("B6").Value = "2010"
        oSheet.Range("C6").Value = "35000"

        ' OK, at this point we have Excel file with 1 sheet with data
        ' Now let's create pivot table

        ' first get range of cells from sheet 1 that will be used by pivot
        Dim xlRange As Excel.Range = CType(oSheet, Excel.Worksheet).Range("A1:C6")

        ' create second sheet
        If oExcel.Application.Sheets.Count() < 2 Then
            oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(2)
        End If
        oSheet.Name = "Pivot Table"

        ' specify first cell for pivot table on the second sheet
        Dim xlRange2 As Excel.Range = CType(oSheet, Excel.Worksheet).Range("B3")

        ' Create pivot cache and table
        Dim ptCache As Excel.PivotCache = oBook.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, xlRange)
        Dim ptTable As Excel.PivotTable = oSheet.PivotTables.Add(PivotCache:=ptCache, TableDestination:=xlRange2, TableName:="Summary")

        ' create Pivot Field, note that pivot field name is the same as column name in sheet 1
        Dim ptField As Excel.PivotField = ptTable.PivotFields("Salary")
        With ptField
            .Orientation = Excel.XlPivotFieldOrientation.xlDataField
            .Function = Excel.XlConsolidationFunction.xlSum
            .Name = " Salary" ' by default name will be something like SumOfSalary, change it here to Salary, note space in front of it - 
            ' this field name cannot be the same as therefore that space
            ' also it cannot be empty

            '' add another field
            'ptField = ptTable.PivotFields("Year")
            'With ptField
            '    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
            '    .Function = Excel.XlConsolidationFunction.xlMax
            '    .Name = " Year" ' this is how you create another field, in my example I don't need it so let's comment it out
            'End With

            ' add column
            ptField = ptTable.PivotFields("First Name")
            With ptField
                .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                .Name = " "
            End With

        End With
        ' add grouping - again I don't need this in my example, this is just to show how to do it
        'oSheet.Range("C5").Group(1, 20, 40)

        oBook.SaveAs(fileTest)
        oBook.Close()
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing



    End Sub

    Private Function CreatePivotTable(ByVal OrigTable As DataTable, Optional ByVal pivotColumnOrdinal As Integer = 0, Optional ByVal pivotRowOrdinal As Integer = 1, _
           Optional ByVal pivotDataOrdinal As Integer = 3, Optional ByVal SortColumn As Boolean = True, Optional ByVal SortRow As Boolean = True) As DataTable

        Dim PivotTable As New DataTable
        Dim OrigArray() As DataRow
        Dim dr As DataRow
        Dim SortString As String
        Dim origRowInd As Integer
        Dim PivotRowInd As Integer
        Dim PivotcolInd As Integer

        Dim CurRowInd As Integer
        Dim CurColInd As Integer
        Dim teststr As String


        Try
            ' add pivot column name 
            PivotTable.Columns.Add(OrigTable.Columns(pivotRowOrdinal).ColumnName)
            ' add pivot column values in each row as column headers to new Table 

            If (SortColumn = True) Then
                SortString = OrigTable.Columns(pivotColumnOrdinal).ColumnName + " ASC"
            Else
                SortString = " "
            End If


            OrigArray = OrigTable.Select("", SortString, DataViewRowState.CurrentRows)

            For origRowInd = 0 To OrigArray.GetUpperBound(0)
                Try
                    PivotTable.Columns.Add(OrigArray(origRowInd).Item(pivotColumnOrdinal))

                Catch ex As Exception

                End Try
            Next

            For PivotcolInd = 0 To PivotTable.Columns.Count - 1
                teststr = PivotTable.Columns(PivotcolInd).ColumnName
            Next

            If (SortRow = True) Then
                SortString = OrigTable.Columns(pivotRowOrdinal).ColumnName + " ASC"
            Else
                SortString = " "
            End If

            OrigArray = OrigTable.Select("", SortString, DataViewRowState.CurrentRows)

            ' loop through rows 
            For origRowInd = 0 To OrigArray.GetUpperBound(0)
                teststr = OrigArray(origRowInd).Item(pivotRowOrdinal)
                For PivotRowInd = 0 To PivotTable.Rows.Count - 1
                    teststr = PivotTable.Rows(PivotRowInd).Item(0)
                    If (OrigArray(origRowInd).Item(pivotRowOrdinal) = PivotTable.Rows(PivotRowInd).Item(0)) Then
                        CurRowInd = PivotRowInd
                        GoTo RowFound
                    End If
                Next

                'add the DataRow to the new table 
                CurRowInd = PivotTable.Rows.Count
                dr = PivotTable.NewRow()
                dr.Item(0) = OrigArray(origRowInd).Item(pivotRowOrdinal)
                teststr = dr.Item(0)
                PivotTable.Rows.Add(dr)

                For PivotcolInd = 1 To PivotTable.Columns.Count - 1
                    PivotTable.Rows(CurRowInd).Item(PivotcolInd) = 0
                Next


RowFound:
                ' loop through columns 
                For PivotcolInd = 0 To PivotTable.Columns.Count - 1
                    teststr = OrigArray(origRowInd).Item(pivotColumnOrdinal)
                    If (OrigArray(origRowInd).Item(pivotColumnOrdinal) = PivotTable.Columns(PivotcolInd).ColumnName) Then
                        CurColInd = PivotcolInd
                        GoTo ColumnFound
                    End If
                Next

ColumnFound:
                PivotTable.Rows(CurRowInd).Item(CurColInd) = PivotTable.Rows(CurRowInd)(CurColInd) + OrigArray(origRowInd)(pivotDataOrdinal)
                teststr = PivotTable.Rows(CurRowInd).Item(0) + " - " + PivotTable.Columns(CurColInd).ColumnName + " - " + PivotTable.Rows(CurRowInd).Item(CurColInd)
            Next

            For CurRowInd = 0 To PivotTable.Rows.Count - 1
                For CurColInd = 0 To PivotTable.Columns.Count - 1
                    teststr = PivotTable.Rows(CurRowInd).Item(0) + " - " + PivotTable.Columns(CurColInd).ColumnName + " - " + PivotTable.Rows(CurRowInd).Item(CurColInd)
                Next
            Next

        Catch ex As Exception

        End Try

        Return PivotTable

    End Function
End Module
