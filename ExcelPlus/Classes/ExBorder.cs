using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExBorder
    {
        #region members
        protected LineTypes lineType = LineTypes.Medium;
        protected Sd.Color color = Sd.Color.Transparent;
        protected bool active = false;
        #endregion

        #region constructor

        public ExBorder()
        {

        }

        public ExBorder(ExBorder border)
        {
            this.lineType = border.lineType;
            this.color = border.color;
            this.active = border.active;
        }

        public ExBorder(XL.XLBorderStyleValues lineType, XL.XLColor color)
        {
            this.lineType = lineType.ToPlus();
            this.color = color.ToColor();
            this.active = !(this.lineType == LineTypes.None);
        }

        #endregion

        #region properties

        public virtual bool Active
        {
            get { return this.active; }
        }

        public LineTypes LineType
        {
            get { return this.lineType; }
            set { this.lineType = value; this.active = true; }
        }

        public Sd.Color Color
        {
            get { return this.color; }
            set { this.color = value; this.active = true; }
        }

        #endregion
    }
}
