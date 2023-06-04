using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace BumblebeeJr
{
    public class XLRange
    {
        #region members

        public XL.IXLRange ComObj = null;

        #endregion

        #region constructors

        public XLRange()
        {
        }

        public XLRange(XL.IXLRange comObj)
        {
            this.ComObj = comObj;
        }

        #endregion


    }
}
