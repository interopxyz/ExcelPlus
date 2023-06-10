using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExWorksheet
    {

        #region members

        public XL.IXLWorksheet ComObj = null;

        protected string name = string.Empty;

        protected ExRange baseRange = new ExRange();

        public List<ExRange> Ranges = new List<ExRange>();

        protected bool active = false;

        public ExGraphic Graphic = new ExGraphic();
        public ExFont Font = new ExFont();

        #endregion

        #region constructors

        public ExWorksheet()
        {
        }

        public ExWorksheet(XL.IXLWorksheet sheet)
        {
            this.name = sheet.Name;
            this.Ranges.Add(new ExRange(sheet.CellsUsed().ToList()));
            this.baseRange = new ExRange(sheet.Cells().ToList());
            this.active = sheet.TabActive;
        }

        public ExWorksheet(string name)
        {
            this.name = name;
        }

        public ExWorksheet(ExRange range)
        {
            this.Ranges.Add(new ExRange(range));
        }

        public ExWorksheet(List<ExRange> ranges)
        {
            foreach (ExRange range in ranges)
            {
                this.Ranges.Add(new ExRange(range));
            }
        }

        public ExWorksheet(List<ExCell> cells)
        {
            this.Ranges.Add(new ExRange(cells));
        }

        public ExWorksheet(ExWorksheet worksheet)
        {
            this.name = worksheet.name;
            this.active = worksheet.active;
            foreach(ExRange range in worksheet.Ranges)
            {
                this.Ranges.Add(new ExRange(range));
            }
            this.baseRange = worksheet.baseRange;

            this.Graphic = new ExGraphic(worksheet.Graphic);
            this.Font = new ExFont(worksheet.Font);
        }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return name; }
            set
            {
                this.name = value;
                if (this.ComObj != null) this.ComObj.Name = this.name;
            }
        }

        public virtual bool Active
        {
            get { return active; }
            set { active = value; }
        }

        public virtual List<ExCell> ActiveCells
        {
            get
            {
                List<ExCell> cells = new List<ExCell>();
                foreach (ExRange range in this.Ranges) cells.AddRange(range.ActiveCells);
                return cells;
            }
        }

        public virtual double ColumnWidth
        {
            get { return this.baseRange.ColumnWidth; }
            set { this.baseRange.ColumnWidth = value; }
        }

        public virtual double RowHeight
        {
            get { return this.baseRange.RowHeight; }
            set { this.baseRange.RowHeight = value; }
        }

        #endregion

        #region methods

        public void ClearValues()
        {
            foreach(ExRange range in Ranges) range.ClearValues();
        }

        public void ClearFormatting()
        {
            this.Graphic = new ExGraphic();
            this.Font = new ExFont();
            foreach (ExRange range in Ranges) range.ClearFormatting();
        }

        public List<ExRange> GetRanges(bool collapse = false)
        {
            List<ExRange> ranges = new List<ExRange>();
            if (collapse)
            {
                ranges.Add(new ExRange(this.ActiveCells));
            }
            else
            {
                foreach (ExRange range in this.Ranges) ranges.Add(new ExRange(range));
            }

            return ranges;
        }

        public bool TryGetRange(int index, out ExRange result)
        {
            if (index>this.Ranges.Count)
            {
                result = new ExRange();
                return true;
            }
            else
            {
                result = new ExRange(this.Ranges[index]);
                return true;
            }
        }

        #endregion

        #region application

        public void ApplyGraphics(XL.IXLWorksheet input)
        {
            if (this.Graphic.Active)
            {
                if (this.Graphic.HasFillColor) input.Style.Fill.SetBackgroundColor(this.Graphic.FillColor.ToExcel());

                if (this.Graphic.BorderBottom.Active)
                {
                    input.Style.Border.SetBottomBorder(this.Graphic.BorderBottom.LineType.ToExcel());
                    input.Style.Border.SetBottomBorderColor(this.Graphic.BorderBottom.Color.ToExcel());
                }

                if (this.Graphic.BorderTop.Active)
                {
                    input.Style.Border.SetTopBorder(this.Graphic.BorderTop.LineType.ToExcel());
                    input.Style.Border.SetTopBorderColor(this.Graphic.BorderTop.Color.ToExcel());
                }

                if (this.Graphic.BorderLeft.Active)
                {
                    input.Style.Border.SetLeftBorder(this.Graphic.BorderLeft.LineType.ToExcel());
                    input.Style.Border.SetLeftBorderColor(this.Graphic.BorderLeft.Color.ToExcel());
                }

                if (this.Graphic.BorderRight.Active)
                {
                    input.Style.Border.SetRightBorder(this.Graphic.BorderRight.LineType.ToExcel());
                    input.Style.Border.SetRightBorderColor(this.Graphic.BorderRight.Color.ToExcel());
                }

                if (this.Graphic.BorderInside.Active)
                {
                    input.Style.Border.SetInsideBorder(this.Graphic.BorderInside.LineType.ToExcel());
                    input.Style.Border.SetInsideBorderColor(this.Graphic.BorderInside.Color.ToExcel());
                }

                if (this.Graphic.BorderOutside.Active)
                {
                    input.Style.Border.SetOutsideBorder(this.Graphic.BorderOutside.LineType.ToExcel());
                    input.Style.Border.SetOutsideBorderColor(this.Graphic.BorderOutside.Color.ToExcel());
                }
            }
        }

        public void ApplyFont(XL.IXLWorksheet input)
        {
            if (this.Font.Active)
            {
                if (this.Font.HasColor) input.Style.Font.SetFontColor(this.Font.Color.ToExcel());
                if (this.Font.HasFamily) input.Style.Font.SetFontName(this.Font.Family);
                if (this.Font.HasSize) input.Style.Font.SetFontSize(this.Font.Size);
                if (this.Font.HasJustification)
                {
                    input.Style.Alignment.Horizontal = this.Font.Justification.ToExcelHAlign();
                    input.Style.Alignment.Vertical = this.Font.Justification.ToExcelVAlign();
                }
                input.Style.Font.SetBold(this.Font.IsBold);
                input.Style.Font.SetItalic(this.Font.IsItalic);
                if (this.Font.IsUnderlined)
                {
                    input.Style.Font.SetUnderline(XL.XLFontUnderlineValues.Single);
                }
                else
                {
                    input.Style.Font.SetUnderline(XL.XLFontUnderlineValues.None);
                }
            }
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Worksheet | " + this.Name;
        }

        #endregion

    }
}
