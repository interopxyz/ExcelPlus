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

        public ExGraphic Graphic = new ExGraphic();
        public ExFont Font = new ExFont();

        protected double columnWidth = -1;
        protected double rowHeight = -1;

        protected bool merge = false;

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
            this.Graphic = new ExGraphic(range.Graphic);
            this.Font = new ExFont(range.Font);
            this.merge = range.merge;

            this.columnWidth = range.columnWidth;
            this.rowHeight = range.rowHeight;
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

        public ExRange(ExCell source, List<List<GH_String>> data, bool byColumn = true)
        {
            int x = data.Count;
            List<ExCell> cells = new List<ExCell>();

            if (byColumn)
            {
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
            }
            else
            {
                for (int i = 0; i < x; i++)
                {
                    int y = data[i].Count;
                    for (int j = 0; j < y; j++)
                    {
                        ExCell cell = new ExCell(source.Column + j, source.Row + i);
                        cell.Value = data[i][j].Value;
                        cells.Add(cell);
                    }
                }
            }
            this.SetCells(cells);
        }

        #endregion

        #region properties

        public virtual List<ExCell> ActiveCells
        {
            get { return cells.Values.ToList(); }
        }

        public virtual List<ExCell> Cells(bool flip = false)
        {
                List<ExCell> output = new List<ExCell>();
            if(flip)
            {
                for (int i = this.min.Column; i < this.max.Column+1; i++)
                {
                    for (int j = this.min.Row; j < this.max.Row+1; j++)
                    {
                        ExCell cell = new ExCell(i, j);
                        if (cells.ContainsKey(cell.Address)) cell = new ExCell(cells[cell.Address]);
                        output.Add(cell);
                    }
                }
            }
            else
            {
                for (int j = this.min.Row; j < this.max.Row+1; j++)
                {
                    for (int i = this.min.Column; i < this.max.Column+1; i++)
                    {
                        ExCell cell = new ExCell(i, j);
                        if (cells.ContainsKey(cell.Address)) cell = new ExCell(cells[cell.Address]);
                        output.Add(cell);
                    }
                }
            }

                return output; 
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

        public virtual double ColumnWidth
        {
            get { return columnWidth; }
            set { columnWidth = value; }
        }

        public virtual double RowHeight
        {
            get { return rowHeight; }
            set { rowHeight = value; }
        }

        public virtual string Address
        {
            get { return min.Address + ":" + max.Address; }
        }

        public virtual bool IsMerged
        {
            get { return this.merge; }
            set { this.merge = value; }
        }

        #endregion

        #region methods

        public void ClearValues()
        {
            foreach (string key in cells.Keys) cells[key].ClearValue();
        }

        public void ClearFormatting()
        {
            this.Graphic = new ExGraphic();
            this.Font = new ExFont();

            this.columnWidth = -1;
            this.rowHeight = -1;

            foreach (string key in cells.Keys) cells[key].ClearFormatting();
        }

        public List<ExRange> SubRanges(bool byColumn)
        {
            List<ExRange> ranges = new List<ExRange>();

            if(byColumn)
            {
                for (int j = this.min.Row; j < this.max.Row+1; j++)
                {
                    List<ExCell> tempCells = new List<ExCell>();
                    for (int i = this.min.Column; i < this.max.Column+1; i++)
                    {
                        ExCell cell = new ExCell(i, j);
                        if (this.cells.ContainsKey(cell.Address)) cell = new ExCell(this.cells[cell.Address]);
                        tempCells.Add(cell);
                    }
                    ranges.Add(new ExRange(tempCells));
                }
            }
            else
            {
                for(int i = this.min.Column; i < this.max.Column+1; i++)
                {
                    List<ExCell> tempCells = new List<ExCell>();
                    for (int j = this.min.Row; j < this.max.Row+1; j++)
                    {
                        ExCell cell = new ExCell(i, j);
                        if (this.cells.ContainsKey(cell.Address)) cell = new ExCell(this.cells[cell.Address]);
                        tempCells.Add(cell);
                    }
                    ranges.Add(new ExRange(tempCells));
                }
            }
            return ranges;
        }

        public List<ExRange> Explode(bool active = false)
        {
            List<ExCell> tempCells = new List<ExCell>();

            if (!active) tempCells = this.Cells();
            if (active) tempCells = this.ActiveCells;

            List<ExRange> ranges = new List<ExRange>();
            foreach (ExCell cell in tempCells) ranges.Add(new ExRange(cell));
            return ranges;
        }

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

        #region application

        public void ApplyGraphics(XL.IXLRange input)
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

        public void ApplyFont(XL.IXLRange input)
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
            return "Range | " + this.Address;
        }

        #endregion

    }
}
