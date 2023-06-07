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

        protected  HorizontalBorder hBorder = HorizontalBorder.None;
        protected  LineType hType = LineType.Continuous;
        protected  BorderWeight hWeight = BorderWeight.Medium;
        protected  Sd.Color hStrokeColor = Sd.Color.Black;

        protected  VerticalBorder vBorder = VerticalBorder.None;
        protected  LineType vType = LineType.Continuous;
        protected  BorderWeight vWeight = BorderWeight.Medium;
        protected  Sd.Color vStrokeColor = Sd.Color.Black;

        protected  Sd.Color fillColor = Sd.Color.Transparent;

        #endregion

        #region constructors

        public ExGraphic()
        {

        }

        public ExGraphic(ExGraphic graphic)
        {
            this.hBorder = graphic.hBorder;
            this.hType = graphic.hType;
            this.hWeight = graphic.hWeight;
            this.hStrokeColor = graphic.hStrokeColor;

            this.vBorder = graphic.vBorder;
            this.vType = graphic.vType;
            this.vWeight = graphic.vWeight;
            this.vStrokeColor = graphic.vStrokeColor;

            this.fillColor = graphic.fillColor;
        }


        #endregion

        #region properties

        public virtual HorizontalBorder HorizontalBorder
        {
            get { return hBorder; }
            set { hBorder = value; }
        }
        public virtual LineType HorizontalType
        {
            get { return hType; }
            set { hType = value; }
        }
        public virtual BorderWeight HorizontalWeight
        {
            get { return hWeight; }
            set { hWeight = value; }
        }
        public virtual Sd.Color HorizontalStrokeColor
        {
            get { return hStrokeColor; }
            set { hStrokeColor = value; }
        }

        public virtual VerticalBorder VerticalBorder
        {
            get { return vBorder; }
            set { vBorder = value; }
        }
        public virtual LineType VerticalType
        {
            get { return vType; }
            set { vType = value; }
        }
        public virtual BorderWeight VerticalWeight
        {
            get { return vWeight; }
            set { vWeight = value; }
        }
        public virtual Sd.Color VerticalStrokeColor
        {
            get { return vStrokeColor; }
            set { vStrokeColor = value; }
        }

        public virtual Sd.Color FillColor
        {
            get { return fillColor; }
            set { fillColor = value; }
        }

        #endregion

        #region methods



        #endregion
    }
}
