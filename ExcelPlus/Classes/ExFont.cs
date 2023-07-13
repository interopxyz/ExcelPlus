using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExFont
    {

        #region members

        protected string family = "None";
        protected double size = -1;
        protected Sd.Color color = Sd.Color.Transparent;
        protected Justifications justification = Justifications.BottomLeft;
        protected bool isBold = false;
        protected bool isItalic = false;
        protected bool isUnderlined = false;
        protected bool active = false;

        #endregion

        #region constructors

        public ExFont()
        {

        }

        public ExFont(ExFont font)
        {
            this.family = font.family;
            this.size = font.size;
            this.color = font.color;
            this.justification = font.justification;
            this.isBold = font.isBold;
            this.isItalic = font.isItalic;
            this.isUnderlined = font.isUnderlined;
            this.active = font.active;
        }

        public ExFont(XL.IXLStyle style, XL.IXLWorkbook workbook)
        {
            this.color = style.Font.FontColor.ToColor(workbook);
            this.family = style.Font.FontName;
            this.size = style.Font.FontSize;
            this.isBold = style.Font.Bold;
            this.isItalic= style.Font.Italic;
            this.isUnderlined = !(style.Font.Underline == XL.XLFontUnderlineValues.None);
            this.justification = style.Alignment.Horizontal.ToJustification(style.Alignment.Vertical);
            this.active = true;
        }


        #endregion

        #region properties

        public virtual bool Active
        {
            get { return active; }
        }

        public virtual string Family
        {
            get { return this.family; }
            set { this.family = value; this.active = true; }
        }

        public virtual bool HasFamily
        {
            get { return this.family != "None"; }
        }

        public virtual double Size
        {
            get { return this.size; }
            set { this.size = value; active = true; }
        }

        public virtual bool HasSize
        {
            get { return this.size > 0; }
        }

        public virtual Sd.Color Color
        {
            get { return this.color; }
            set { this.color = value; }
        }

        public virtual bool HasColor
        {
            get { return this.color != Sd.Color.Transparent; }
        }

        public virtual Justifications Justification
        {
            get { return this.justification; }
            set { this.justification = value; }
        }

        public virtual bool HasJustification
        {
            get { return this.justification != Justifications.None; }
        }

        public virtual bool IsBold
        {
            get { return this.isBold; }
            set { this.isBold = value; }
        }

        public virtual bool IsItalic
        {
            get { return this.isItalic; }
            set { this.isItalic = value; }
        }

        public virtual bool IsUnderlined
        {
            get { return this.isUnderlined; }
            set { this.isUnderlined = value; }
        }

        #endregion

        #region methods

        public void Apply(XL.IXLStyle input)
        {
            if (this.Active)
            {
                if (this.HasColor) input.Font.SetFontColor(this.Color.ToExcel());
                if (this.HasFamily) input.Font.SetFontName(this.Family);
                if (this.HasSize) input.Font.SetFontSize(this.Size);
                if (this.HasJustification)
                {
                    input.Alignment.Horizontal = this.Justification.ToExcelHAlign();
                    input.Alignment.Vertical = this.Justification.ToExcelVAlign();
                }
                input.Font.SetBold(this.IsBold);
                input.Font.SetItalic(this.IsItalic);
                if (this.IsUnderlined)
                {
                    input.Font.SetUnderline(XL.XLFontUnderlineValues.Single);
                }
                else
                {
                    input.Font.SetUnderline(XL.XLFontUnderlineValues.None);
                }
            }
        }

        #endregion

    }
}