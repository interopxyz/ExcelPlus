using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rg = Rhino.Geometry;

namespace ExcelPlus
{
    public static class GhExtensions
    {

        public static bool TryGetCell(this IGH_Goo input, out ExCell cell)
        {

            if (input == null)
            {
                cell = new ExCell();
                return false;
            }

            if (input.CastTo<ExCell>(out ExCell cl))
            {
                cell = new ExCell(cl);
                return true;
            }
            else if (input.CastTo<GH_Point>(out GH_Point gpt))
            {
                cell = new ExCell((int)gpt.Value.X, (int)gpt.Value.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point3d>(out Rg.Point3d p3d))
            {
                cell = new ExCell((int)p3d.X, (int)p3d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point2d>(out Rg.Point2d p2d))
            {
                cell = new ExCell((int)p2d.X, (int)p2d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Interval>(out Rg.Interval domain))
            {
                cell = new ExCell((int)domain.T0, (int)domain.T1);
                return true;
            }
            else if (input.CastTo<string>(out string address))
            {
                cell = new ExCell(address);
                return true;
            }
            else
            {
                cell = null;
                return true;
            }
        }

        public static bool TryGetRange(this IGH_Goo input, out ExRange range)
        {

            if (input == null)
            {
                range = new ExRange();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out ExWorkbook book))
            {
                book.TryGetSheet(0, out ExWorksheet sheetOut);
                sheetOut.TryGetRange(0, out ExRange rangeOut);
                range = rangeOut;
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet.TryGetRange(0, out ExRange rangeOut);
                range = rangeOut;
                return true;
            }
            else if (input.CastTo<ExRange>(out ExRange rng))
            {
                range = new ExRange(rng);
                return true;
            }
            else if (input.CastTo<ExCell>(out ExCell cell))
            {
                range = new ExRange(cell);
                return true;
            }
            else
            {
                range = new ExRange();
                return true;
            }
        }

        public static bool TryGetWorksheet(this IGH_Goo input, out ExWorksheet worksheet)
        {

            if (input == null)
            {
                worksheet = new ExWorksheet();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out ExWorkbook book))
            {
                book.TryGetSheet(0, out ExWorksheet result);
                worksheet = result;
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                worksheet = new ExWorksheet(sheet);
                return true;
            }
            else if (input.CastTo<ExRange>(out ExRange range))
            {
                worksheet = new ExWorksheet(range);
                return true;
            }
            else if (input.CastTo<ExCell>(out ExCell cell))
            {
                worksheet = new ExWorksheet(new ExRange(cell));
                return true;
            }
            else if (input.CastTo<string>(out string name))
            {
                worksheet = new ExWorksheet(name);
                return true;
            }
            else
            {
                worksheet = new ExWorksheet();
                return true;
            }
        }

        public static bool TryGetWorkbook(this IGH_Goo input, out ExWorkbook workbook)
        {

            if (input == null)
            {
                workbook = new ExWorkbook();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out ExWorkbook book))
            {
                workbook = new ExWorkbook(book);
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                workbook = new ExWorkbook( sheet);
                return true;
            }
            else if (input.CastTo<ExRange>(out ExRange range))
            {
                workbook = new ExWorkbook(new ExWorksheet(range));
                return true;
            }
            else if (input.CastTo<ExCell>(out ExCell cell))
            {
                workbook = new ExWorkbook( new ExWorksheet(new ExRange(cell)));
                return true;
            }
            else if (input.CastTo<string>(out string name))
            {
                workbook = new ExWorkbook(name);
                return true;
            }
            else
            {
                workbook = new ExWorkbook();
                return true;
            }
        }

        public static bool TryCompileWorkbook(this List<IGH_Goo> inputs, out ExWorkbook workbook)
        {
            ExWorkbook book = new ExWorkbook();
            ExWorksheet sheet = new ExWorksheet();
            ExRange range = new ExRange();

                workbook = book;
            if (inputs == null)  return true;
            if (inputs.Count<1) return true;
            
            bool hasWb = false;
            foreach (IGH_Goo input in inputs)
            {
                if (input.CastTo<ExWorkbook>(out ExWorkbook outWb))
                {
                    if (hasWb)
                    {
                        book.Sheets.AddRange(outWb.Sheets);
                    }
                    else
                    {
                        book = new ExWorkbook(outWb);
                        hasWb = true;
                    }
                }
            }
            bool hasRange = false;
            bool hasCell = false;
            foreach (IGH_Goo input in inputs)
            {
                if (input.CastTo<ExWorksheet>(out ExWorksheet outSh))
                {
                    book.Sheets.Add(outSh);
                }
                else if (input.CastTo<ExRange>(out ExRange outRn))
                {
                    sheet.Ranges.Add(outRn);
                    hasRange = true;
                }
                else if (input.CastTo<ExCell>(out ExCell outCl))
                {
                    range.SetCells(outCl);
                    hasCell = true;
                    hasRange = true;
                }
            }

            if (hasCell) sheet.Ranges.Add(range);
            if (hasRange) book.Sheets.Add(sheet);
            workbook = book;
            return true;
        }

    }
}
