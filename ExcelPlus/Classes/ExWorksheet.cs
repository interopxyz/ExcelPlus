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

        public List<ExRange> Ranges = new List<ExRange>();

        #endregion

        #region constructors

        public ExWorksheet()
        {
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
            foreach(ExRange range in worksheet.Ranges)
            {
                this.Ranges.Add(new ExRange(range));
            }
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

        public virtual List<ExCell> ActiveCells
        {
            get
            {
                List<ExCell> cells = new List<ExCell>();
                foreach (ExRange range in this.Ranges) cells.AddRange(range.Cells);
                return cells;
            }
        }

        #endregion

        #region methods

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

        #region overrides

        public override string ToString()
        {
            return "Worksheet | " + Name;
        }

        #endregion

    }
}
