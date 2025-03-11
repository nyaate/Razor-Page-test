using System.Collections.Generic;
using System.Text;
using Microsoft.AspNetCore.Razor.TagHelpers;

namespace myRazorPages1.Pages.Support.Samples.RazorPages
{
    [HtmlTargetElement("range-to-table")]
    public class RangeToTableTagHelper : TagHelper
    {
      
        public required SpreadsheetGear.IRange Range { get; set; }

        public bool FirstRowIsHeader { get; set; } = true;

        public override void Process(TagHelperContext context, TagHelperOutput output)
        {
            if (Range == null)
            {
                output.TagName = "div";
                output.Content.SetHtmlContent(@"<div class='alert alert-info'><i>No range data available.</i></div>");
                return;
            }

            output.TagName = "table";
            output.Attributes.Add("class", "table table-striped table-bordered table-hover table-sm");

            StringBuilder sb = new();

            SpreadsheetGear.IRange dataRange = Range;
            if (FirstRowIsHeader)
            {
                SpreadsheetGear.IRange headerRow = dataRange[0, 0, 0, dataRange.ColumnCount - 1];
                sb.Append("<thead class='table-dark'><tr>");

                foreach (SpreadsheetGear.IRange cell in headerRow)
                {
                    string classes = GetClassAttribute(cell);

                    sb.Append("<th" + (classes.Length > 0 ? $" class='{classes}'" : "") + ">").Append(cell.Text).Append("</th>");
                }
                sb.Append("</tr></thead>");

                dataRange = dataRange.Subtract(headerRow);
            }

            sb.Append("<tbody>");
            if (dataRange != null)
            {
                foreach (SpreadsheetGear.IRange row in dataRange.Rows)
                {
                    sb.Append("<tr>");
                   
                    foreach (SpreadsheetGear.IRange cell in row.Columns)
                    {
                        string classes = GetClassAttribute(cell);

                        sb.Append("<td" + (classes.Length > 0 ? $" class='{classes}'" : "") + ">").Append(cell.Text).Append("</td>");
                    }
                    sb.Append("</tr>");
                }
            }
            else
            {
                sb.Append($"<tr><td colspan='{Range.ColumnCount}' class='text-center text-muted'>No Data Available</td></tr>");
            }
            sb.Append("</tbody>");

            output.Content.SetHtmlContent(sb.ToString());
        }

        /// <param name="cell">A single cell, for which formatting classes will be based off.</param>
        private static string GetClassAttribute(SpreadsheetGear.IRange cell)
        {
            List<string> value = new();
            List<string> classes = value;

            if (cell.HorizontalAlignment == SpreadsheetGear.HAlign.Center)
                classes.Add("text-center");
            else if (cell.HorizontalAlignment == SpreadsheetGear.HAlign.Right)
                classes.Add("text-end");

            if (cell.Font.Bold)
                classes.Add("fw-bold");
            if (cell.Font.Italic)
                classes.Add("fst-italic");
            if (cell.Font.Underline != SpreadsheetGear.UnderlineStyle.None)
                classes.Add("text-underline");

            if (classes.Count > 0)
                return string.Join(' ', classes);
            return "";
        }
    }
}