using System;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SpreadsheetGear;

namespace myRazorPages1.Pages.Support.Samples.RazorPages.Reporting
{
    public partial class ExcelWorkbookConsolidationModel : PageModel
    {
        [BindProperty]
        public required string Region { get; set; }

        public required IRange DataRange { get; set; }


        public void OnGet() { }


        public void OnPostRenderInTable()
        {
            if (Region == null)
                return;

            IWorkbook workbook = GetSalesReportWorkbook();

            if (Region == "All")
            {
                DataRange = workbook.Worksheets[0].UsedRange;
            }
            else
            {
                DataRange = workbook.Names["YearSales"].RefersToRange;
            }
        }

        public FileResult OnPostDownloadWorkbook()
        {
            IWorkbook workbook = GetSalesReportWorkbook();

            System.IO.Stream workbookStream = workbook.SaveToStream(FileFormat.OpenXMLWorkbook);

            workbookStream.Seek(0, System.IO.SeekOrigin.Begin);

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileName = $"Sales-{Region}.xlsx";
            return File(workbookStream, contentType, fileName);
        }


        private IWorkbook GetSalesReportWorkbook()
        {
            if (Region == "All")
            {
                return GetWorkbookConsolidated();
            }
            else
            {
                return GetWorkbookForRegion(Region);
            }
        }

        private static IWorkbook GetWorkbookForRegion(string region)
        {
            string filename = region switch
            {
                "South" => "spicesouth.xlsx",
                "East" => "spiceeast.xlsx",
                "West" => "spicewest.xlsx",
                _ => "spicenorth.xlsx",
            };
            return Factory.GetWorkbook("files/" + filename);
        }


        private static IWorkbook GetWorkbookConsolidated()
        {
            IWorkbook workbook = Factory.GetWorkbook();
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Total Sales";

            CopyRegion(worksheet, "North", PasteOperation.None);
            CopyRegion(worksheet, "South", PasteOperation.Add);
            CopyRegion(worksheet, "East", PasteOperation.Add);
            CopyRegion(worksheet, "West", PasteOperation.Add);

            worksheet.UsedRange.Columns.AutoFit();

            return workbook;
        }

        private static void CopyRegion(IWorksheet dstWorksheet, string region, PasteOperation pasteOperation)
        {
            IWorkbook srcWorkbook = GetWorkbookForRegion(region);
            IRange srcRange = srcWorkbook.Names["YearSales"].RefersToRange;

            string address = srcRange.Address;
            IRange dstRange = dstWorksheet.Cells[address];

            srcRange.Copy(dstRange,
                PasteType.Values,
                pasteOperation, true, false);

            srcRange.Copy(dstRange,
                PasteType.Formats,
                pasteOperation, true, false);
        }
    }
}