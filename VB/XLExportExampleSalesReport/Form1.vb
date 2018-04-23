Imports DevExpress.Export.Xl
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

Namespace XLExportExampleSalesReport
    Partial Public Class Form1
        Inherits Form

        ' Obtain the data source.
        Private sales As List(Of SalesData) = SalesDataRepository.CreateSalesData()
        Private headerRowFormatting As XlCellFormatting
        Private dataRowFormatting As XlCellFormatting
        Private totalRowFormatting As XlCellFormatting
        Private grandTotalRowFormatting As XlCellFormatting

        Public Sub New()
            InitializeComponent()
            InitializeFormatting()
        End Sub

        Private Sub InitializeFormatting()
            ' Specify formatting settings for the header rows.
            headerRowFormatting = New XlCellFormatting()
            headerRowFormatting.Font = XlFont.BodyFont()
            headerRowFormatting.Font.Bold = True
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))
            headerRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom)

            ' Specify formatting settings for the data rows.
            dataRowFormatting = New XlCellFormatting()
            dataRowFormatting.Font = XlFont.BodyFont()
            dataRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0))

            ' Specify formatting settings for the total rows.
            totalRowFormatting = New XlCellFormatting()
            totalRowFormatting.Font = XlFont.BodyFont()
            totalRowFormatting.Font.Bold = True
            totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0))

            ' Specify formatting settings for the grand total row.
            grandTotalRowFormatting = New XlCellFormatting()
            grandTotalRowFormatting.Font = XlFont.BodyFont()
            grandTotalRowFormatting.Font.Bold = True
            grandTotalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.2))
        End Sub

        ' Export the document to XLSX format.
        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLSX.Click
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xlsx) Then
                ShowFile(fileName)
            End If
        End Sub

        ' Export the document to XLS format.
        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLS.Click
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xls) Then
                ShowFile(fileName)
            End If
        End Sub

        ' Export the document to CSV format.
        Private Sub btnExportToCSV_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToCSV.Click
            Dim fileName As String = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Csv) Then
                ShowFile(fileName)
            End If
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaulName As String) As String
            saveFileDialog1.Filter = filter
            saveFileDialog1.FileName = defaulName
            If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                Return Nothing
            End If
            Return saveFileDialog1.FileName
        End Function

        Private Sub ShowFile(ByVal fileName As String)
            If Not File.Exists(fileName) Then
                Return
            End If
            Dim dResult As DialogResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName), Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dResult = System.Windows.Forms.DialogResult.Yes Then
                Process.Start(fileName)
            End If
        End Sub

        Private Function ExportToFile(ByVal fileName As String, ByVal documentFormat As XlDocumentFormat) As Boolean
            Try
                Using stream As New FileStream(fileName, FileMode.Create)
                    ' Create an exporter instance.
                    Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
                    ' Create a new document and begin to write it to the specified stream.
                    Using document As IXlDocument = exporter.CreateDocument(stream)
                        ' Generate the document content.
                        GenerateDocument(document)
                    End Using
                End Using
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        Private Sub GenerateDocument(ByVal document As IXlDocument)
            ' Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture

            ' Add a new worksheet to the document.
            Using sheet As IXlSheet = document.CreateSheet()
                ' Specify the worksheet name.
                sheet.Name = "Sales Report"

                ' Specify print settings for the worksheet. 
                SetupPageParameters(sheet)

                ' Specify the summary row and summary column location for the grouped data.
                sheet.OutlineProperties.SummaryBelow = True
                sheet.OutlineProperties.SummaryRight = True

                ' Generate worksheet columns.
                GenerateColumns(sheet)

                ' Add the document title.
                GenerateTitle(sheet)

                ' Begin to group worksheet rows (create the outer group of rows).
                sheet.BeginGroup(False)

                ' Create the query expression to retrieve data from the sales list and group data by the State.
                ' Query variable is an IEnumerable<IGrouping<string, SalesData>>. 
                Dim statesQuery = From data In sales _
                                  Group data By data.State Into dataGroup = Group _
                                  Order By State _
                                  Select State

                ' Create data rows to display sales for each state.  
                For Each state As String In statesQuery
                    GenerateData(sheet, state)
                Next state

                ' Finalize the group creation.
                sheet.EndGroup()

                ' Create the grand total row.
                GenerateGrandTotalRow(sheet)

                ' Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange
            End Using
        End Sub

        Private Sub GenerateColumns(ByVal sheet As IXlSheet)
            ' Create the column "A" and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 18
            End Using

            ' Create the column "B" and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 166
            End Using

            Dim numberFormat As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"

            ' Begin to group worksheet columns starting from the column "C" to the column "F".
            sheet.BeginGroup(False)

            ' Create four successive columns ("C", "D", "E" and "F") and set the specific number format for their cells.
            For i As Integer = 0 To 3
                Using column As IXlColumn = sheet.CreateColumn()
                    column.WidthInPixels = 117
                    column.ApplyFormatting(numberFormat)
                End Using
            Next i

            ' Finalize the group creation. 
            sheet.EndGroup()

            ' Create the summary column "G", adjust its width and set the specific number format for its cells.  
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 117
                column.ApplyFormatting(numberFormat)
            End Using
        End Sub

        Private Sub GenerateTitle(ByVal sheet As IXlSheet)
            ' Specify formatting settings for the document title.
            Dim formatting As New XlCellFormatting()
            formatting.Font = New XlFont()
            formatting.Font.Name = "Calibri Light"
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None
            formatting.Font.Size = 24
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.5)
            formatting.Border = New XlBorder()
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.5)
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium

            ' Add the document title.
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the cell "A1".  
                row.SkipCells(1)
                ' Create the cell "B1" containing the document title.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "SALES ANALYSIS 2014"
                    cell.Formatting = formatting
                End Using
                ' Create five empty cells with the title formatting.
                row.BlankCells(5, formatting)
            End Using

            ' Skip one row before starting to generate data rows.
            sheet.SkipRows(1)

            ' Insert a picture from a file and anchor it to the cell "G1".
            Dim startupPath As String = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
            Using picture As IXlPicture = sheet.CreatePicture()
                picture.Image = Image.FromFile(Path.Combine(startupPath, "Logo.png"))
                picture.SetOneCellAnchor(New XlAnchorPoint(6, 0, 8, 4), 105, 30)
            End Using

        End Sub

        Private Sub GenerateData(ByVal sheet As IXlSheet, ByVal nameOfState As String)
            ' Create the header row for the state sales.
            GenerateHeaderRow(sheet, nameOfState)

            Dim firstDataRowIndex As Integer = sheet.CurrentRowIndex

            ' Begin to group worksheet rows (create the inner group of rows containing sales data for the specific state).
            sheet.BeginGroup(False)

            ' Create the query expression to retrieve sales data for the specified State. Then, sort data by the Product key in ascending order. 
            Dim salesQuery = From data In sales _
                             Where data.State = nameOfState _
                             Order By data.Product _
                             Select data

            ' Create the data row to display sales information for each product. 
            For Each data As SalesData In salesQuery
                GenerateDataRow(sheet, data)
            Next data

            ' Finalize the group creation. 
            sheet.EndGroup()

            ' Create the summary row for the group. 
            GenerateTotalRow(sheet, firstDataRowIndex)
        End Sub

        Private Sub GenerateHeaderRow(ByVal sheet As IXlSheet, ByVal nameOfState As String)
            ' Create the header row for sales data in the specific state.
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the first cell in the row.
                row.SkipCells(1)

                ' Create the cell that displays the state name and specify its format settings. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = nameOfState
                    cell.ApplyFormatting(headerRowFormatting)
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0)))
                    cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.General, XlVerticalAlignment.Bottom))
                End Using

                ' Create four successive cells with values "Q1", "Q2", "Q3" and "Q4". 
                ' Apply specific formatting settings to the created cells.  
                For i As Integer = 0 To 3
                    Using cell As IXlCell = row.CreateCell()
                        cell.Value = String.Format("Q{0}", i + 1)
                        cell.ApplyFormatting(headerRowFormatting)
                    End Using
                Next i

                ' Create the "Yearly total" cell and specify its format settings.  
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "Yearly total"
                    cell.ApplyFormatting(headerRowFormatting)
                End Using
            End Using
        End Sub

        Private Sub GenerateDataRow(ByVal sheet As IXlSheet, ByVal data As SalesData)
            ' Create the row to display sales information for each sale item.  
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the first row in the cell.
                row.SkipCells(1)

                ' Create the cell to display the product name and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Product
                    cell.ApplyFormatting(dataRowFormatting)
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8))
                End Using

                ' Create the cell to display sales amount in the first quarter and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Q1
                    cell.ApplyFormatting(dataRowFormatting)
                End Using

                ' Create the cell to display sales amount in the second quarter and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Q2
                    cell.ApplyFormatting(dataRowFormatting)
                End Using

                ' Create the cell to display sales amount in the third quarter and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Q3
                    cell.ApplyFormatting(dataRowFormatting)
                End Using

                ' Create the cell to display sales amount in the fourth quarter and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Q4
                    cell.ApplyFormatting(dataRowFormatting)
                End Using

                ' Create the cell to display annual sales for the product. Use the SUM function to add product sales in each quarter.   
                Using cell As IXlCell = row.CreateCell()
                    cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(2, row.RowIndex, 5, row.RowIndex)))
                    cell.ApplyFormatting(dataRowFormatting)
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)))
                End Using
            End Using
        End Sub

        Private Sub GenerateTotalRow(ByVal sheet As IXlSheet, ByVal firstDataRowIndex As Integer)
            ' Create the total row for each inner group of sales in the specific state.
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the first cell in the row.
                row.SkipCells(1)

                ' Create the "Total" cell and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "Total"
                    cell.ApplyFormatting(totalRowFormatting)
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6))
                End Using

                ' Create four successive cells displaying total sales for each quarter individually. Use the SUBTOTAL function to add quarterly sales. 
                For j As Integer = 0 To 3
                    Using cell As IXlCell = row.CreateCell()
                        cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 2, firstDataRowIndex, j + 2, row.RowIndex - 1), XlSummary.Sum, False))
                        cell.ApplyFormatting(totalRowFormatting)
                    End Using
                Next j

                ' Create the cell that displays yearly sales for the state. Use the SUBTOTAL function to add yearly sales in the current state for each product. 
                Using cell As IXlCell = row.CreateCell()
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(6, firstDataRowIndex, 6, row.RowIndex - 1), XlSummary.Sum, False))
                    cell.ApplyFormatting(totalRowFormatting)
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.1)))
                End Using
            End Using
        End Sub

        Private Sub GenerateGrandTotalRow(ByVal sheet As IXlSheet)
            ' Create the grand total row.
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the first cell in the row.
                row.SkipCells(1)

                ' Create the "Grand Total" cell and specify its format settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "Grand Total"
                    cell.ApplyFormatting(grandTotalRowFormatting)
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.4))
                End Using

                ' Create five successive cells displaying quarterly total sales and annual sales for all states. The SUBTOTAL function is used to calculate subtotals for the related rows in each column. 
                For j As Integer = 0 To 4
                    Using cell As IXlCell = row.CreateCell()
                        cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 2, 3, j + 2, row.RowIndex - 1), XlSummary.Sum, False))
                        cell.ApplyFormatting(grandTotalRowFormatting)
                    End Using
                Next j
            End Using
        End Sub

        Private Sub SetupPageParameters(ByVal sheet As IXlSheet)
            ' Specify the header and footer for the odd-numbered pages.
            sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.Bold & "DevAV", Nothing, XlHeaderFooter.Date)
            sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR("Sales report", Nothing, XlHeaderFooter.PageNumber & " of " & XlHeaderFooter.PageTotal)

            ' Specify page settings.
            sheet.PageSetup = New XlPageSetup()
            ' Scale the print area to fit to one page wide.
            sheet.PageSetup.FitToPage = True
            sheet.PageSetup.FitToWidth = 1
            sheet.PageSetup.FitToHeight = 0
        End Sub
    End Class
End Namespace
