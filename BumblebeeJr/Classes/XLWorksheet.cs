using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace BumblebeeJr
{
    public class XLWorksheet
    {

        #region members

        public XL.IXLWorksheet ComObj = null;

        protected string name = "";

        #endregion

        #region constructors

        public XLWorksheet()
        {
        }

        //public XLWorksheet(ExRange range)
        //{
        //    this.ComObj = range.ComObj.Worksheet;
        //    this.name = this.ComObj.Name;
        //}

        public XLWorksheet(XL.IXLWorksheet comObj)
        {
            this.ComObj = comObj;
            this.name = comObj.Name;
        }

        public XLWorksheet(XLWorksheet worksheet)
        {
            this.ComObj = worksheet.ComObj;
            this.name = worksheet.Name;
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

        #endregion

        #region data



        public XLRange WriteData(List<List<GH_String>> data, XLCell source)
        {
            int y = data[0].Count;
            int x = data.Count;

            string[,] values = new string[y, x];
            double[,] numbers = new double[y, x];
            double num = 0;
            bool isNumeric = true;

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[j, i] = data[i][j].Value;
                    if (double.TryParse(values[j, i], out num)) numbers[j, i] = num; else isNumeric = false;
                }
            }

            string target = Helper.GetCellAddress(source.Column + x - 1, source.Row + y - 1);
            XL.IXLRange rng = this.ComObj.Range(source.ToString());

            if (isNumeric) SetData(rng, numbers);
            if (!isNumeric) SetData(rng, values);

            return new XLRange();
        }

        protected void SetData(XL.IXLRange rng, string[,] values)
        {
        }

        protected void SetData(XL.IXLRange rng, double[,] values)
        {
        }

        #endregion

    }
}
