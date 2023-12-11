using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDRS_Metadata
{
    class FormatExcel
    {

        public static void FormatTableOverviewCells(Worksheet tableOverviewWorksheet, int count)
        {
            Range tempRng = tableOverviewWorksheet.Cells[2, 1];
            tempRng.Activate();
            tempRng.Application.ActiveWindow.FreezePanes = true;

            tempRng = tableOverviewWorksheet.Range["A1", "I1"];
            tempRng.Characters.Font.Bold = true;

            // Border lines
            for (int n = 1; n < 10; n++)
            {
                if (n < 6)
                {
                    tempRng = tableOverviewWorksheet.Cells[1, n];
                    tempRng.Interior.Color = Color.LightGray;
                }

                for (int m = 1; m < count; m++)
                {
                    tempRng = tableOverviewWorksheet.Cells[m, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }
            }

            // Cell background color
            for (int m = 1; m < count; m++)
            {
                tempRng = tableOverviewWorksheet.Cells[m, 6];
                tempRng.Interior.Color = Color.LightYellow;

                tempRng = tableOverviewWorksheet.Cells[m, 7];
                tempRng.Interior.Color = Color.LightGreen;

                tempRng = tableOverviewWorksheet.Cells[m, 8];
                tempRng.Interior.Color = Color.LightSkyBlue;

                tempRng = tableOverviewWorksheet.Cells[m, 9];
                tempRng.Interior.Color = Color.LightGray;
            }

            // Alignment
            tableOverviewWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableOverviewWorksheet.Columns.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Column widths
            tableOverviewWorksheet.Columns["A:A"].AutoFit();
            tableOverviewWorksheet.Columns["B:B"].AutoFit();  // .ColumnWidth = 8;
            tableOverviewWorksheet.Columns["C:C"].AutoFit();  // .ColumnWidth = 8;

            tableOverviewWorksheet.Columns["D:D"].AutoFit();
            tableOverviewWorksheet.Columns["D:D"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["E:E"].AutoFit();
            tableOverviewWorksheet.Columns["E:E"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["F:F"].ColumnWidth = 10;
            tableOverviewWorksheet.Columns["F:F"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["G:G"].ColumnWidth = 20;
            tableOverviewWorksheet.Columns["G:G"].WrapText = true;

            tableOverviewWorksheet.Columns["H:H"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["H:H"].WrapText = true;

            tableOverviewWorksheet.Columns["I:I"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["I:I"].WrapText = true;

            // Column sorting
            tableOverviewWorksheet.Sort.SortFields.Clear();

            tableOverviewWorksheet.Sort.SortFields.Add(tableOverviewWorksheet.Range["F:F"], XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, "HIGH, MEDIUM, LOW, SYSTEM, STATS, EMPTY, DUMMY", XlSortDataOption.xlSortNormal);
            tableOverviewWorksheet.Sort.SetRange(tableOverviewWorksheet.UsedRange);
            tableOverviewWorksheet.Sort.Header = XlYesNoGuess.xlYes;
            tableOverviewWorksheet.Sort.Apply();
        }
    }
}
