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

            this.Graphic = cell.Graphic;
            this.Font = cell.Font;

            this.width = cell.width;
            this.height = cell.height;
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
            this.Graphic = new ExGraphic();
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
