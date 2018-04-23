using DevExpress.Export.Xl;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace XLExportExampleSalesReport {
    public partial class Form1 : Form {
        // Obtain the data source.
        List<SalesData> sales = SalesDataRepository.CreateSalesData();
        XlCellFormatting headerRowFormatting;
        XlCellFormatting dataRowFormatting;
        XlCellFormatting totalRowFormatting;
        XlCellFormatting grandTotalRowFormatting;

        public Form1() {
            InitializeComponent();
            InitializeFormatting();
        }

        void InitializeFormatting() {
            // Specify formatting settings for the header rows.
            headerRowFormatting = new XlCellFormatting();
            headerRowFormatting.Font = XlFont.BodyFont();
            headerRowFormatting.Font.Bold = true;
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
            headerRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom);

            // Specify formatting settings for the data rows.
            dataRowFormatting = new XlCellFormatting();
            dataRowFormatting.Font = XlFont.BodyFont();
            dataRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0));

            // Specify formatting settings for the total rows.
            totalRowFormatting = new XlCellFormatting();
            totalRowFormatting.Font = XlFont.BodyFont();
            totalRowFormatting.Font.Bold = true;
            totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0));

            // Specify formatting settings for the grand total row.
            grandTotalRowFormatting = new XlCellFormatting();
            grandTotalRowFormatting.Font = XlFont.BodyFont();
            grandTotalRowFormatting.Font.Bold = true;
            grandTotalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.2));
        }

        // Export the document to XLSX format.
        void btnExportToXLSX_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xlsx))
                ShowFile(fileName);
        }

        // Export the document to XLS format.
        void btnExportToXLS_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xls))
                ShowFile(fileName);
        }

        // Export the document to CSV format.
        void btnExportToCSV_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Csv))
                ShowFile(fileName);
        }

        string GetSaveFileName(string filter, string defaulName) {
            saveFileDialog1.Filter = filter;
            saveFileDialog1.FileName = defaulName;
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return null;
            return saveFileDialog1.FileName;
        }

        void ShowFile(string fileName) {
            if (!File.Exists(fileName))
                return;
            DialogResult dResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName),
                this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dResult == DialogResult.Yes)
                Process.Start(fileName);
        }

        bool ExportToFile(string fileName, XlDocumentFormat documentFormat) {
            try {
                using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                    // Create an exporter instance.
                    IXlExporter exporter = XlExport.CreateExporter(documentFormat);
                    // Create a new document and begin to write it to the specified stream.
                    using (IXlDocument document = exporter.CreateDocument(stream)) {
                        // Generate the document content.
                        GenerateDocument(document);
                    }
                }
                return true;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        void GenerateDocument(IXlDocument document) {
            // Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture;

            // Add a new worksheet to the document.
            using (IXlSheet sheet = document.CreateSheet()) {
                // Specify the worksheet name.
                sheet.Name = "Sales Report";

                // Specify print settings for the worksheet. 
                SetupPageParameters(sheet);

                // Specify the summary row and summary column location for the grouped data.
                sheet.OutlineProperties.SummaryBelow = true;
                sheet.OutlineProperties.SummaryRight = true;

                // Generate worksheet columns.
                GenerateColumns(sheet);

                // Add the document title.
                GenerateTitle(sheet);

                // Begin to group worksheet rows (create the outer group of rows).
                sheet.BeginGroup(false);

                // Create the query expression to retrieve data from the sales list and group data by the State.
                // Query variable is an IEnumerable<IGrouping<string, SalesData>>. 
                var statesQuery = from data in sales
                                  group data by data.State into dataGroup
                                  orderby dataGroup.Key
                                  select dataGroup.Key;

                // Create data rows to display sales for each state.  
                foreach (string state in statesQuery)
                    GenerateData(sheet, state);

                // Finalize the group creation.
                sheet.EndGroup();                                       

                // Create the grand total row.
                GenerateGrandTotalRow(sheet);

                // Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange;
            }
        }

        void GenerateColumns(IXlSheet sheet) {
            // Create the column "A" and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 18;

            // Create the column "B" and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 166;

            XlNumberFormat numberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";

            // Begin to group worksheet columns starting from the column "C" to the column "F".
            sheet.BeginGroup(false);

            // Create four successive columns ("C", "D", "E" and "F") and set the specific number format for their cells.
            for (int i = 0; i < 4; i++) {
                using (IXlColumn column = sheet.CreateColumn()) {
                    column.WidthInPixels = 117;
                    column.ApplyFormatting(numberFormat);
                }
            }

            // Finalize the group creation. 
            sheet.EndGroup();

            // Create the summary column "G", adjust its width and set the specific number format for its cells.  
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 117;
                column.ApplyFormatting(numberFormat);
            }
        }

        void GenerateTitle(IXlSheet sheet) {
            // Specify formatting settings for the document title.
            XlCellFormatting formatting = new XlCellFormatting();
            formatting.Font = new XlFont();
            formatting.Font.Name = "Calibri Light";
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            formatting.Font.Size = 24;
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.5);
            formatting.Border = new XlBorder();
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.5);
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;

            // Add the document title.
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the cell "A1".  
                row.SkipCells(1);
                // Create the cell "B1" containing the document title.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "SALES ANALYSIS 2014";
                    cell.Formatting = formatting;
                }
                // Create five empty cells with the title formatting.
                row.BlankCells(5, formatting);
            }

            // Skip one row before starting to generate data rows.
            sheet.SkipRows(1);

            // Insert a picture from a file and anchor it to the cell "G1".
            string startupPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            using (IXlPicture picture = sheet.CreatePicture()) {
                picture.Image = Image.FromFile(Path.Combine(startupPath, "Logo.png"));
                picture.SetOneCellAnchor(new XlAnchorPoint(6, 0, 8, 4), 105, 30);
            }

        }

        void GenerateData(IXlSheet sheet, string nameOfState) {
            // Create the header row for the state sales.
            GenerateHeaderRow(sheet, nameOfState);

            int firstDataRowIndex = sheet.CurrentRowIndex;

            // Begin to group worksheet rows (create the inner group of rows containing sales data for the specific state).
            sheet.BeginGroup(false);

            // Create the query expression to retrieve sales data for the specified State. Then, sort data by the Product key in ascending order. 
            var salesQuery = from data in sales
                              where data.State == nameOfState
                              orderby data.Product
                              select data;

            // Create the data row to display sales information for each product. 
            foreach (SalesData data in salesQuery)
                GenerateDataRow(sheet, data);

            // Finalize the group creation. 
            sheet.EndGroup();

            // Create the summary row for the group. 
            GenerateTotalRow(sheet, firstDataRowIndex);
        }

        void GenerateHeaderRow(IXlSheet sheet, string nameOfState) {
            // Create the header row for sales data in the specific state.
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the first cell in the row.
                row.SkipCells(1);

                // Create the cell that displays the state name and specify its format settings. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = nameOfState;
                    cell.ApplyFormatting(headerRowFormatting);
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0)));
                    cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.General, XlVerticalAlignment.Bottom));
                }

                // Create four successive cells with values "Q1", "Q2", "Q3" and "Q4". 
                // Apply specific formatting settings to the created cells.  
                for (int i = 0; i < 4; i++) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.Value = string.Format("Q{0}", i + 1);
                        cell.ApplyFormatting(headerRowFormatting);
                    }
                }

                // Create the "Yearly total" cell and specify its format settings.  
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "Yearly total";
                    cell.ApplyFormatting(headerRowFormatting);
                }
            }
        }

        void GenerateDataRow(IXlSheet sheet, SalesData data) {
            // Create the row to display sales information for each sale item.  
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the first row in the cell.
                row.SkipCells(1);

                // Create the cell to display the product name and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Product;
                    cell.ApplyFormatting(dataRowFormatting);
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8));
                }

                // Create the cell to display sales amount in the first quarter and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Q1;
                    cell.ApplyFormatting(dataRowFormatting);
                }

                // Create the cell to display sales amount in the second quarter and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Q2;
                    cell.ApplyFormatting(dataRowFormatting);
                }

                // Create the cell to display sales amount in the third quarter and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Q3;
                    cell.ApplyFormatting(dataRowFormatting);
                }

                // Create the cell to display sales amount in the fourth quarter and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Q4;
                    cell.ApplyFormatting(dataRowFormatting);
                }

                // Create the cell to display annual sales for the product. Use the SUM function to add product sales in each quarter.   
                using (IXlCell cell = row.CreateCell()) {
                    cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(2, row.RowIndex, 5, row.RowIndex)));
                    cell.ApplyFormatting(dataRowFormatting);
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                }
            }
        }

        void GenerateTotalRow(IXlSheet sheet, int firstDataRowIndex) {
            // Create the total row for each inner group of sales in the specific state.
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the first cell in the row.
                row.SkipCells(1);

                // Create the "Total" cell and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "Total";
                    cell.ApplyFormatting(totalRowFormatting);
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                }

                // Create four successive cells displaying total sales for each quarter individually. Use the SUBTOTAL function to add quarterly sales. 
                for (int j = 0; j < 4; j++) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 2, firstDataRowIndex, j + 2, row.RowIndex - 1), XlSummary.Sum, false));
                        cell.ApplyFormatting(totalRowFormatting);
                    }
                }

                // Create the cell that displays yearly sales for the state. Use the SUBTOTAL function to add yearly sales in the current state for each product. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(6, firstDataRowIndex, 6, row.RowIndex - 1), XlSummary.Sum, false));
                    cell.ApplyFormatting(totalRowFormatting);
                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.1)));
                }
            }
        }

        void GenerateGrandTotalRow(IXlSheet sheet) {
            // Create the grand total row.
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the first cell in the row.
                row.SkipCells(1);

                // Create the "Grand Total" cell and specify its format settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "Grand Total";
                    cell.ApplyFormatting(grandTotalRowFormatting);
                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.4));
                }

                // Create five successive cells displaying quarterly total sales and annual sales for all states. The SUBTOTAL function is used to calculate subtotals for the related rows in each column. 
                for (int j = 0; j < 5; j++) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 2, 3, j + 2, row.RowIndex - 1), XlSummary.Sum, false));
                        cell.ApplyFormatting(grandTotalRowFormatting);
                    }
                }
            }
        }

        void SetupPageParameters(IXlSheet sheet) {
            // Specify the header and footer for the odd-numbered pages.
            sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.Bold + "DevAV", null, XlHeaderFooter.Date);
            sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR("Sales report", null, XlHeaderFooter.PageNumber + " of " + XlHeaderFooter.PageTotal);

            // Specify page settings.
            sheet.PageSetup = new XlPageSetup();
            // Scale the print area to fit to one page wide.
            sheet.PageSetup.FitToPage = true;
            sheet.PageSetup.FitToWidth = 1;
            sheet.PageSetup.FitToHeight = 0;
        }
    }
}
