using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

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
        }

        #endregion

        #region properties

        public virtual bool Active
        {
            get { return active; }
        }

        public LineTypes LineType
        {
            get { return lineType; }
            set { lineType = value; active = true; }
        }

        public Sd.Color Color
        {
            get { return color; }
            set { color = value; active = true; }
        }

        #endregion
    }
}
