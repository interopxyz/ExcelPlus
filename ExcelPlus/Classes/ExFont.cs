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



        #endregion

    }
}