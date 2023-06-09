﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using XL = ClosedXML.Excel;

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
            this.borderTop = new ExBorder(graphic.borderTop);
            this.borderBottom = new ExBorder(graphic.borderBottom);
            this.borderRight = new ExBorder(graphic.borderRight);
            this.borderLeft = new ExBorder(graphic.borderLeft);
            this.borderInside = new ExBorder(graphic.borderInside);
            this.borderOutside = new ExBorder(graphic.borderOutside);

            this.fillColor = graphic.fillColor;

            this.active = graphic.active;
        }

        public ExGraphic(XL.IXLStyle style, XL.IXLWorkbook workbook)
        {
            this.borderTop = new ExBorder(style.Border.TopBorder, style.Border.TopBorderColor, workbook);
            this.borderBottom = new ExBorder(style.Border.BottomBorder, style.Border.BottomBorderColor, workbook);
            this.borderRight = new ExBorder(style.Border.RightBorder, style.Border.RightBorderColor, workbook);
            this.borderLeft = new ExBorder(style.Border.LeftBorder, style.Border.LeftBorderColor, workbook);

            this.fillColor = style.Fill.BackgroundColor.ToColor(workbook);
            this.active = true;
        }

        #endregion

        #region properties

        public virtual bool Active
        {
            get
            {
                if (active) return true;
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
            get { return new ExBorder(this.borderTop); }
            set { this.borderTop = value; }
        }

        public virtual ExBorder BorderBottom
        {
            get { return new ExBorder(this.borderBottom); }
            set { this.borderBottom = value; }
        }

        public virtual ExBorder BorderLeft
        {
            get { return new ExBorder(this.borderLeft); }
            set { this.borderLeft = value; }
        }

        public virtual ExBorder BorderRight
        {
            get { return new ExBorder(this.borderRight); }
            set { this.borderRight = value; }
        }

        public virtual ExBorder BorderOutside
        {
            get { return new ExBorder(this.borderOutside); }
            set { this.borderOutside = value; }
        }

        public virtual ExBorder BorderInside
        {
            get { return new ExBorder(this.borderInside); }
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

        public void Apply(XL.IXLStyle input)
        {
            if (this.Active)
            {
                if (this.HasFillColor) input.Fill.SetBackgroundColor(this.FillColor.ToExcel());

                if (this.BorderBottom.Active)
                {
                    input.Border.SetBottomBorder(this.BorderBottom.LineType.ToExcel());
                    input.Border.SetBottomBorderColor(this.BorderBottom.Color.ToExcel());
                }

                if (this.BorderTop.Active)
                {
                    input.Border.SetTopBorder(this.BorderTop.LineType.ToExcel());
                    input.Border.SetTopBorderColor(this.BorderTop.Color.ToExcel());
                }

                if (this.BorderLeft.Active)
                {
                    input.Border.SetLeftBorder(this.BorderLeft.LineType.ToExcel());
                    input.Border.SetLeftBorderColor(this.BorderLeft.Color.ToExcel());
                }

                if (this.BorderRight.Active)
                {
                    input.Border.SetRightBorder(this.BorderRight.LineType.ToExcel());
                    input.Border.SetRightBorderColor(this.BorderRight.Color.ToExcel());
                }

                if (this.BorderInside.Active)
                {
                    input.Border.SetInsideBorder(this.BorderInside.LineType.ToExcel());
                    input.Border.SetInsideBorderColor(this.BorderInside.Color.ToExcel());
                }

                if (this.BorderOutside.Active)
                {
                    input.Border.SetOutsideBorder(this.BorderOutside.LineType.ToExcel());
                    input.Border.SetOutsideBorderColor(this.BorderOutside.Color.ToExcel());
                }
            }
        }

        #endregion

    }
}
