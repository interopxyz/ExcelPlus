using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        protected DataType type = DataType.String;

        #endregion

        #region constructors

        public ExCell()
        {

        }

        public ExCell(ExCell cell)
        {
            this.column = cell.column;
            this.row = cell.row;

            this.isColumnAbsolute = cell.isColumnAbsolute;
            this.isRowAbsolute = cell.isRowAbsolute;

            this.value = cell.value;
            this.type = cell.type;
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
            set
            {
                this.value = value;
                this.type = DataType.String;
            }
        }
            

        public virtual int Column
        {
            get { return column; }
            set
            {
                this.column = value;
            }
        }

        public virtual int Row
        {
            get { return row; }
            set
            {
                this.row = value;
            }
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
        }

        public virtual bool IsRowAbsolute
        {
            get { return isRowAbsolute; }
        }


        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            return "Cell | " + this.Address;
        }

        #endregion

    }
}
