﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExWorkbook
    {

        #region members

        public XL.IXLWorkbook ComObj = null;

        protected string name = "";

        public List<ExWorksheet> Sheets = new List<ExWorksheet>();

        #endregion

        #region constructors

        public ExWorkbook()
        {
        }

        public ExWorkbook(string name)
        {
            this.name = name;
        }

        public ExWorkbook(ExWorkbook workbook)
        {
            this.name = workbook.Name;
            foreach (ExWorksheet sheet in workbook.Sheets)
            {
                this.Sheets.Add(new ExWorksheet(sheet));
            }
        }

        public ExWorkbook(List<ExWorksheet> sheets)
        {
            foreach (ExWorksheet sheet in sheets)
            {
                this.Sheets.Add(new ExWorksheet(sheet));
            }
        }

        public ExWorkbook(ExWorksheet sheet)
        {
                this.Sheets.Add(new ExWorksheet(sheet));
        }

        public ExWorkbook(List<ExRange> ranges)
        {
                this.Sheets.Add(new ExWorksheet(ranges));
        }

        public ExWorkbook(List<ExCell> cells)
        {
            this.Sheets.Add(new ExWorksheet(new List<ExRange> { new ExRange(cells)}));
        }


        #endregion

        #region properties

        public virtual string Name
        {
            get { return this.name; }
            set 
            { 
                this.name = value;
            }
        }

        public virtual double ColumnWidth
        {
            set { foreach (ExWorksheet sheet in this.Sheets) sheet.ColumnWidth = value; }
        }

        public virtual double RowHeight
        {
            set { foreach (ExWorksheet sheet in this.Sheets) sheet.RowHeight = value; }
        }

        #endregion

        #region methods

        public virtual List<ExWorksheet> GetSheets()
        {
            List<ExWorksheet> output = new List<ExWorksheet>();
            foreach (ExWorksheet sheet in this.Sheets) output.Add(new ExWorksheet(sheet));
            return output;
        }

        public void ClearValues()
        {
            foreach (ExWorksheet sheet in this.Sheets) sheet.ClearValues();
        }

        public void ClearFormatting()
        {
            foreach (ExWorksheet sheet in this.Sheets) sheet.ClearFormatting();
        }

        public List<ExWorksheet> GetWorksheets()
        {
            List<ExWorksheet> sheets = new List<ExWorksheet>();
            foreach (ExWorksheet sheet in this.Sheets) sheets.Add(new ExWorksheet(sheet));

            return sheets;
        }

        public bool TryGetSheet(int index, out ExWorksheet worksheet)
        {
            if (index > this.Sheets.Count)
            {
                worksheet = new ExWorksheet();
                return true;
            }
            else
            {
                worksheet = new ExWorksheet(this.Sheets[index]);
                return true;
            }
        }

        public bool TryGetSheet(string name, out ExWorksheet worksheet)
        {

            foreach (ExWorksheet sheet in this.Sheets)
            {
                if (sheet.Name == name)
                {
                    worksheet = new ExWorksheet(sheet);
                    return true;
                }
            }

            worksheet = null;
            return false;
        }

        public string Save(string directory, Extensions extension)
        {
            if(name == string.Empty)
            {
                return this.Save(directory, Constants.UniqueName, extension);
            }
            else
            {
                return this.Save(directory, name, extension);
            }
        }

        public string Save(string directory, string filename, Extensions extension )
        {
            string folder = Path.GetDirectoryName(directory);
            string name = Path.GetFileNameWithoutExtension(filename);
            string filepath = folder + "/" + name + "."+extension.ToString();

            this.name = name;
            this.CompileWorkbook();

            this.ComObj.SaveAs(filepath);

            return filepath;
        }

        public void Refresh()
        {
            this.CompileWorkbook();

            this.ParseWorkbook();
        }

        public void Open(string directory, string filename, Extensions extension)
        {
            string folder = Path.GetDirectoryName(directory);
            string name = Path.GetFileNameWithoutExtension(filename);
            string filepath = folder + "/" + name + "." + extension.ToString();

            XL.LoadOptions options = new XL.LoadOptions();
            options.RecalculateAllFormulas = true;
            FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            this.ComObj = new XL.XLWorkbook(fileStream);
            fileStream.Dispose();

            this.ParseWorkbook();
            this.name = name;
        }

        public void Open(string filepath)
        {
            XL.LoadOptions options = new XL.LoadOptions();
            options.RecalculateAllFormulas = true;
            FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            this.ComObj = new XL.XLWorkbook(fileStream);
            fileStream.Dispose();

            this.ParseWorkbook();
            this.name = Path.GetFileNameWithoutExtension(filepath);
        }

        public string Write()
        {
            this.CompileWorkbook();
            Stream stream = new MemoryStream();
            this.ComObj.SaveAs(stream,true);
            stream.Position = 0;
            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, (int)stream.Length);

            return System.Text.Encoding.Default.GetString(buffer);
        }

        public void Read(string byteArrayString)
        {
            byte[] buffer = Encoding.Default.GetBytes(byteArrayString);
            Stream stream = new MemoryStream(buffer);
            stream.Position = 0;
            this.ComObj = new XL.XLWorkbook(stream, new XL.LoadOptions());
            this.ParseWorkbook();
        }

        private void CompileWorkbook()
        {
            this.ComObj = new XL.XLWorkbook();
            foreach (ExWorksheet sheet in Sheets)
            {

                XL.IXLWorksheet xlSheet = this.ComObj.AddWorksheet();
                if (sheet.Name != string.Empty) xlSheet.Name = sheet.Name;
                if (sheet.Active) xlSheet.SetTabActive();

                sheet.ApplyFormatting(xlSheet);

                if (sheet.ColumnWidth > 0) xlSheet.ColumnWidth = sheet.ColumnWidth;
                if (sheet.RowHeight > 0) xlSheet.RowHeight= sheet.RowHeight;

                foreach (ExRange range in sheet.Ranges)
                {
                    XL.IXLRange xlRange = xlSheet.Range(range.Min.Row, range.Min.Column, range.Max.Row, range.Max.Column);
                    range.ApplyFormatting(xlRange);

                    if (range.ColumnWidth > 0) xlSheet.Columns(range.Min.Column, range.Max.Column).Width = range.ColumnWidth;
                    if (range.RowHeight > 0) xlSheet.Rows(range.Min.Row, range.Max.Row).Height = range.RowHeight;
                }

                List<ExCell> cells = sheet.ActiveCells;

                foreach (ExCell cell in cells)
                {
                    XL.IXLCell xlCell = xlSheet.Cell(cell.Row, cell.Column);

                    cell.SetValue(xlCell);
                    cell.ApplyFormatting(xlCell);
                }

                foreach (ExRange range in sheet.Ranges)
                {

                    if (range.IsMerged)
                    {
                        XL.IXLRange xlRange = xlSheet.Range(range.Min.Row, range.Min.Column, range.Max.Row, range.Max.Column);
                        xlRange.Merge(true);
                    }
                }

            }

            this.ComObj.RecalculateAllFormulas();
        }

        private void ParseWorkbook()
        {
            List<ExWorksheet> sheets = new List<ExWorksheet>();

            foreach(XL.IXLWorksheet sheet in this.ComObj.Worksheets.ToList())
            {
                sheets.Add(new ExWorksheet(sheet));
            }

            this.Sheets = sheets;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Workbook | " + this.Name;
        }

        #endregion

    }
}
