using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExCell
    {

        #region members

        protected int column = 1;
        protected int row = 1;
        protected bool isColumnAbsolute = false;
        protected bool isRowAbsolute = false;
        protected string value = string.Empty;
        protected string format = "None";
        protected ContentTypes contentType = ContentTypes.Value;

        protected double width = -1;
        protected double height = -1;

        public ExGraphic Graphic = new ExGraphic();
        public ExFont Font = new ExFont();

        #endregion

        #region constructors

        public ExCell()
        {

        }

        public ExCell(XL.IXLCell cell)
        {
            this.column = cell.Address.ColumnNumber;
            this.row = cell.Address.RowNumber;
            this.isColumnAbsolute = cell.Address.FixedColumn;
            this.isRowAbsolute = cell.Address.FixedRow;

            switch (cell.Value.Type)
            {
                case XL.XLDataType.Text:
                    this.value = cell.Value.GetText();
                    break;
                case XL.XLDataType.Number:
                    this.value = Convert.ToString(cell.Value.GetNumber());
                    break;
                case XL.XLDataType.Boolean:
                    this.value = Convert.ToString(cell.Value.GetBoolean());
                    break;
                case XL.XLDataType.DateTime:
                    this.value = Convert.ToString(cell.Value.GetDateTime());
                    break;
            }

            this.Graphic = new ExGraphic(cell.Style);
            this.Font = new ExFont(cell.Style);

            this.height = cell.WorksheetRow().Height;
            this.width = cell.WorksheetColumn().Width;
        }

        public ExCell(ExCell cell)
        {
            this.column = cell.column;
            this.row = cell.row;

            this.isColumnAbsolute = cell.isColumnAbsolute;
            this.isRowAbsolute = cell.isRowAbsolute;

            this.value = cell.value;
            this.format = cell.format;

            this.Graphic = new ExGraphic( cell.Graphic);
            this.Font = new ExFont(cell.Font);

            this.width = cell.width;
            this.height = cell.height;

            this.contentType = cell.contentType;
        }

        public ExCell(string address)
        {
            this.Address = address;
        }

        public ExCell(int column, int row)
        {
            this.column = column;
            this.row = row;
        }

        public ExCell(int column, int row, bool absColumn, bool absRow)
        {
            this.column = column;
            this.row = row;
            this.isColumnAbsolute = absColumn;
            this.isRowAbsolute = absRow;
        }

        #endregion

        #region properties

        public virtual bool IsFormula
        {
            get 
            { 
                if (this.value == string.Empty) return false;
                if (this.value == "") return false;
                if (this.value[0] == '=') return true;
                return false;
            }
        }

        public virtual string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }

        public virtual string Format
        {
            get { return this.format; }
            set { this.format = value; }
        }

        public virtual double Width
        {
            get { return width; }
            set { this.width = value; }
        }

        public virtual double Height
        {
            get { return height; }
            set { this.height = value; }
        }

        public virtual int Column
        {
            get { return column; }
            set { this.column = value; }
        }

        public virtual int Row
        {
            get { return row; }
            set { this.row = value; }
        }

        public virtual string Address
        {
            get { return Helper.GetCellAddress(column, row, IsColumnAbsolute, isRowAbsolute); }
            set
            {
                Tuple<int, int> loc = Helper.GetCellLocation(value);
                Tuple<bool, bool> abs = Helper.GetCellAbsolute(value);

                this.column = loc.Item1;
                this.row = loc.Item2;

                this.isColumnAbsolute = abs.Item1;
                this.isRowAbsolute = abs.Item2;
            }
        }

        public virtual bool IsColumnAbsolute
        {
            get { return isColumnAbsolute; }
            set { isColumnAbsolute = value; }
        }

        public virtual bool IsRowAbsolute
        {
            get { return isRowAbsolute; }
            set { isRowAbsolute = value; }
        }


        #endregion

        #region methods

        public void ClearValue()
        {
            this.value = string.Empty;
            this.format = "General";
        }

        public void ClearFormatting()
        {

            this.width = -1;
            this.height = -1;

            this.Graphic = new ExGraphic();
            this.Font = new ExFont();
        }

        #endregion

        #region application

        public void ApplyGraphics(XL.IXLCell input)
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

        public void ApplyFont(XL.IXLCell input)
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
            return "Cell | " + this.Address;
        }

        #endregion

    }
}
