using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace ExcelPlus
{
    public class ExFont
    {

        #region members

        protected string family = "Arial";
        protected double size = 12;
        protected Sd.Color color = Sd.Color.Black;
        protected Justification justification = Justification.BottomLeft;
        protected bool isBold = false;
        protected bool isItalic = false;
        protected bool isUnderlined = false;

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
        }


        #endregion

        #region properties

        public virtual string Family
        {
            get { return family; }
            set { family = value; }
        }

        public virtual double Size
        {
            get { return size; }
            set { size = value; }
        }

        public virtual Sd.Color Color
        {
            get { return color; }
            set { color = value; }
        }

        public virtual Justification Justification
        {
            get { return justification; }
            set { justification = value; }
        }

        public virtual bool IsBold
        {
            get { return isBold; }
            set { isBold = value; }
        }

        public virtual bool IsItalic
        {
            get { return isItalic; }
            set { isItalic = value; }
        }

        public virtual bool IsUnderlined
        {
            get { return isUnderlined; }
            set { isUnderlined = value; }
        }

        #endregion

        #region methods



        #endregion

    }
}