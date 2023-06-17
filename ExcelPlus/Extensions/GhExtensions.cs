using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

namespace ExcelPlus
{
    public static class GhExtensions
    {
        public static bool TryGetBitmap(this IGH_Goo goo, out Sd.Bitmap bitmap, out string path)
        {

            string filePath = string.Empty;
            goo.CastTo<string>(out filePath);
            Sd.Bitmap bmp = null;

            path = filePath;
            bitmap = bmp;

            if (goo.CastTo<Sd.Bitmap>(out bmp))
            {
                bitmap = new Sd.Bitmap(bmp);
                return true;
            }
            else if (File.Exists(filePath))
            {
                if (filePath.GetBitmapFromFile(out bmp))
                {
                    path = filePath;
                    bitmap = bmp;
                    return true;
                }
                return false;
            }
            return false;
        }
        public static bool GetBitmapFromFile(this string FilePath, out Sd.Bitmap bitmap)
        {
            bitmap = null;
            if (Path.HasExtension(FilePath))
            {
                string extension = Path.GetExtension(FilePath);
                switch (extension)
                {
                    default:
                        return (false);
                    case ".bmp":
                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    case ".jfif":
                    case ".gif":
                    case ".tif":
                    case ".tiff":
                        bitmap = (Sd.Bitmap)Sd.Bitmap.FromFile(FilePath);
                        return (bitmap != null);
                }

            }

            return (false);
        }

        public static bool TryGetCell(this IGH_Goo input, out ExCell cell)
        {

            if (input == null)
            {
                cell = new ExCell("A1");
                return true;
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
            else if (input.CastTo<GH_Vector>(out GH_Vector vpt))
            {
                cell = new ExCell((int)vpt.Value.X, (int)vpt.Value.Y);
                return true;
            }
            else if (input.CastTo<Rg.Vector3d>(out Rg.Vector3d v3d))
            {
                cell = new ExCell((int)v3d.X, (int)v3d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Vector2d>(out Rg.Vector2d v2d))
            {
                cell = new ExCell((int)v2d.X, (int)v2d.Y);
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
            range = new ExRange();

            if (input == null)
            {
                range = new ExRange();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out ExWorkbook book))
            {
                book.TryGetSheet(0, out ExWorksheet sheetOut);
                sheetOut.TryGetRange(0, out ExRange rangeOut);
                range = new ExRange(rangeOut);
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet.TryGetRange(0, out ExRange rangeOut);
                range = new ExRange(rangeOut);
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
            else if (input.CastTo<string>(out string address))
            {
                if (!address.Contains(":")) return true;
                string[] addresses = address.Split(':');
                if (addresses.Count() < 2) return true;
                range = new ExRange(new ExCell(addresses[0]), new ExCell(addresses[1]));
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
                worksheet = new ExWorksheet(result);
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
                        book.Sheets.AddRange(outWb.GetSheets());
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
                    book.Sheets.Add(new ExWorksheet( outSh));
                }
                else if (input.CastTo<ExRange>(out ExRange outRn))
                {
                    sheet.Ranges.Add(new ExRange(outRn));
                    hasRange = true;
                }
                else if (input.CastTo<ExCell>(out ExCell outCl))
                {
                    range.SetCells(new ExCell(outCl));
                    hasCell = true;
                    hasRange = true;
                }
            }

            if (hasCell) sheet.Ranges.Add(range);
            if (hasRange) book.Sheets.Add(sheet);
            workbook = book;
            return true;
        }

        public static bool TryCompileWorksheet(this List<IGH_Goo> inputs, out ExWorksheet worksheet)
        {
            ExWorksheet sheet = new ExWorksheet();
            ExRange range = new ExRange();

            worksheet = sheet;
            if (inputs == null) return true;
            if (inputs.Count < 1) return true;

            bool hasCell = false;
            foreach (IGH_Goo input in inputs)
            {
                if (input.CastTo<ExWorkbook>(out ExWorkbook outWb))
                {
                    foreach(ExWorksheet sht in outWb.Sheets) sheet.Ranges.AddRange(sht.GetRanges());
                }
                if (input.CastTo<ExWorksheet>(out ExWorksheet outSh))
                {
                    sheet.Ranges.AddRange(outSh.GetRanges());
                }
                else if (input.CastTo<ExRange>(out ExRange outRn))
                {
                    sheet.Ranges.Add(new ExRange(outRn));
                }
                else if (input.CastTo<ExCell>(out ExCell outCl))
                {
                    range.SetCells(new ExCell(outCl));
                    hasCell = true;
                }
            }

            if (hasCell) sheet.Ranges.Add(range);
            worksheet = sheet;
            return true;
        }

        public static bool TryCompileRange(this List<IGH_Goo> inputs, out ExRange range)
        {
            ExRange rng = new ExRange();
            List<ExCell> cells = new List<ExCell>();

            range = rng;
            if (inputs == null) return true;
            if (inputs.Count < 1) return true;

            foreach (IGH_Goo input in inputs)
            {
                if (input.CastTo<ExCell>(out ExCell cell))
                {
                    cells.Add(new ExCell(cell));
                }
                else if(input.CastTo<ExRange>(out ExRange rnge))
                {
                    cells.AddRange(rnge.ActiveCells);
                }
                else if (input.CastTo<ExWorksheet>(out ExWorksheet sheet))
                {
                    cells.AddRange(sheet.ActiveCells);
                }
                else if (input.CastTo<ExWorkbook>(out ExWorkbook book))
                {
                    foreach(ExWorksheet workSheet in book.Sheets)
                    {
                    cells.AddRange(workSheet.ActiveCells);
                    }
                }
            }

            rng.SetCells(cells);
            range = rng;
            return true;
        }

    }
}
