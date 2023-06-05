using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExRange
    {
        # region members

        public XL.IXLRange ComObj = null;

        protected Dictionary<string,ExCell> cells = new Dictionary<string, ExCell>();
        protected ExCell min = new ExCell();
        protected ExCell max = new ExCell();

        #endregion

        #region constructors

        public ExRange()
        {

        }

        public ExRange(List<XL.IXLCell> cells)
        {
            List<ExCell> newCells = new List<ExCell>();
            foreach(XL.IXLCell cell in cells)
            {
                newCells.Add(new ExCell(cell));
            }
            this.SetCells(newCells);
        }

        public ExRange(ExRange range)
        {
            this.ComObj = range.ComObj;
            this.cells = range.cells;
            this.min = range.min;
            this.max = range.max;
        }

        public ExRange(ExCell cell)
        {
            this.SetCells(cell);
        }

        public ExRange(List<ExCell> cells)
        {
            this.SetCells(cells);
        }

        public ExRange(ExCell min, ExCell max)
        {
            this.SetCells(new ExCell[] { min, max });
        }

        public ExRange(ExCell source, List<List<GH_String>> data)
        {
            int x = data.Count;
            List<ExCell> cells = new List<ExCell>();

            for (int i = 0; i < x; i++)
            {
                int y = data[i].Count;
                for (int j = 0; j < y; j++)
                {
                    ExCell cell = new ExCell(source.Column + i, source.Row + j);
                    cell.Value = data[i][j].Value;
                    cells.Add(cell);
                }
            }
            this.SetCells(cells);
        }

        #endregion

        #region properties

        public virtual List<ExCell> Cells
        {
            get { return cells.Values.ToList(); }
        }

        public virtual ExCell Min
        {
            get { return new ExCell(min); }
        }

        public virtual ExCell Max
        {
            get { return new ExCell(max); }
        }

        public virtual bool IsSingle
        {
            get { return ((min.Column == max.Column) & (min.Row == max.Row)); }
        }

        public virtual int Width
        {
            get { return (max.Column - min.Column); }
        }

        public virtual int Height
        {
            get { return (max.Row - min.Row); }
        }

        public virtual string Address
        {
            get { return min.Address + ":" + max.Address; }
                }

        #endregion

        #region methods

        protected void ReCalculateBounds()
        {
            List<ExCell> temp = cells.Values.ToList();
            ExCell t = temp[0];
            int minX = t.Column;
            int maxX = t.Column;
            int minY = t.Row;
            int maxY = t.Row;

            foreach(ExCell cell in temp)
            {
                minX = Math.Min(minX, cell.Column);
                maxX = Math.Max(maxX, cell.Column);

                minY = Math.Min(minY, cell.Row);
                maxY = Math.Max(maxY, cell.Row);
            }

            this.min = new ExCell(minX, minY);
            this.max = new ExCell(maxX, maxY);

            if (cells.ContainsKey(min.Address)) min = new ExCell(cells[min.Address]);
            if (cells.ContainsKey(max.Address)) max = new ExCell(cells[max.Address]);
        }

        public void SetCells(ExCell cell, bool overwrite = true)
        {
            if (!this.cells.ContainsKey(cell.Address))
            {
                this.cells.Add(cell.Address, new ExCell(cell));
                this.ReCalculateBounds();
            }
            else
            {
                if (overwrite) this.cells[cell.Address] = new ExCell(cell);
            }
        }

        public void SetCells(List<ExCell> cells, bool overwrite = true)
        {
            bool rangeChanged = false;
                foreach(ExCell cell in cells) { 
            if (!this.cells.ContainsKey(cell.Address))
            {
                this.cells.Add(cell.Address, new ExCell(cell));
                    rangeChanged = true;
            }
            else
            {
                if (overwrite) this.cells[cell.Address] = new ExCell(cell);
            }
            }
                if(rangeChanged) this.ReCalculateBounds();
        }

        public void SetCells(ExCell[] cells, bool overwrite = true)
        {
            bool rangeChanged = false;
            foreach (ExCell cell in cells)
            {
                if (!this.cells.ContainsKey(cell.Address))
                {
                    this.cells.Add(cell.Address, new ExCell(cell));
                    rangeChanged = true;
                }
                else
                {
                    if (overwrite) this.cells[cell.Address] = new ExCell(cell);
                }
            }
            if (rangeChanged) this.ReCalculateBounds();
        }

        public bool TryGetCell(ExCell cell, out ExCell result)
        {
            result = new ExCell();
            if (cell.Column > max.Column) return false;
            if (cell.Column < min.Column) return false;
            if (cell.Row > max.Row) return false;
            if (cell.Row < max.Row) return false;

            if (cells.ContainsKey(cell.Address))
            {
                result = new ExCell( cells[cell.Address]);
                return true;
            }
            else
            {
                result = new ExCell(cell);
                return true;
            }
        }

        public bool TryGetCell(int column, int row, out ExCell result)
        {
            bool status = this.TryGetCell(new ExCell(column, row), out ExCell output);
            result = output;
            return status;
        }

        public bool TryGetCell(string address, out ExCell result)
        {
            bool status = this.TryGetCell(new ExCell(address), out ExCell output);
            result = output;
            return status;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Range | " + this.Address;
        }

        #endregion

    }
}
