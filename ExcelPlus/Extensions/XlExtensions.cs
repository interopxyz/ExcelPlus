using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using XL = ClosedXML.Excel;


namespace ExcelPlus
{
    public static class XlExtensions
    {
        public static XL.XLColor ToExcel(this Sd.Color input)
        {
            return XL.XLColor.FromColor(input);
        }


    }
}
