using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace ExcelPlus
{
    public class ExGraphic
    {

        #region members

        protected ExBorder borderTop = new ExBorder();
        protected ExBorder borderBottom = new ExBorder();
        protected ExBorder borderRight = new ExBorder();
        protected ExBorder borderLeft = new ExBorder();
        protected ExBorder borderInside = new ExBorder();
        protected ExBorder borderOutside = new ExBorder();

        protected Sd.Color fillColor = Sd.Color.Transparent;

        protected bool active = false;

        #endregion

        #region constructors

        public ExGraphic()
        {

        }

        public ExGraphic(ExGraphic graphic)
        {
            this.borderTop = graphic.borderTop;
            this.borderBottom = graphic.borderBottom;
            this.borderRight = graphic.borderRight;
            this.borderLeft = graphic.borderLeft;
            this.borderInside = graphic.borderInside;
            this.borderOutside = graphic.borderOutside;

            this.fillColor = graphic.fillColor;

            this.active = graphic.active;
        }


        #endregion

        #region properties

        public virtual bool Active
        {
            get 
            { 
                if(active)return true;
                if (borderTop.Active) return true;
                if (borderBottom.Active) return true;
                if (borderRight.Active) return true;
                if (borderLeft.Active) return true;
                if (borderInside.Active) return true;
                if (borderOutside.Active) return true;
                return false;
            }
        }

        public virtual ExBorder BorderTop
        {
            get { return this.borderTop; }
            set { this.borderTop = value;}
        }

        public virtual ExBorder BorderBottom
        {
            get { return this.borderBottom; }
            set { this.borderBottom = value; }
        }

        public virtual ExBorder BorderLeft
        {
            get { return this.borderLeft; }
            set { this.borderLeft = value; }
        }

        public virtual ExBorder BorderRight
        {
            get { return this.borderRight; }
            set { this.borderRight = value; }
        }

        public virtual ExBorder BorderOutside
        {
            get { return this.borderOutside; }
            set { this.borderOutside = value; }
        }

        public virtual ExBorder BorderInside
        {
            get { return this.borderInside; }
            set { this.borderInside = value; }
        }

        public virtual Sd.Color FillColor
        {
            get { return this.fillColor; }
            set { this.fillColor = value; this.active = true; }
        }

        public virtual bool HasFillColor
        {
            get { return this.fillColor != Sd.Color.Transparent; }
        }

        #endregion

        #region methods



        #endregion
    }
}
