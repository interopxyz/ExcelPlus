using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace ExcelPlus
{
    public class ExSpark
    {
        #region members

        protected ExRange location = new ExRange();
        protected ExRange range = new ExRange();

        protected Sd.Color color = Constants.StartColor;
        protected double weight = 1.0;

        protected bool isColumn = false;

        #endregion

        #region constructors

        public ExSpark()
        {
        }

        public ExSpark(ExSpark spark)
        {
            this.range = spark.Range;
            this.location = spark.Location;

            this.color = spark.color;
            this.weight = spark.weight;
            this.isColumn = spark.isColumn;
        }

        public static ExSpark ConstructLine(ExRange location, ExRange range, Sd.Color color, double weight)
        {
            ExSpark spark = new ExSpark();

            spark.Location = location;
            spark.Range = range;
            spark.color = color;
            spark.weight = weight;
            spark.isColumn = false;

            return spark;
        }

        public static ExSpark ConstructBar(ExRange location, ExRange range, Sd.Color color, double weight)
        {
            ExSpark spark = new ExSpark();

            spark.Location = location;
            spark.Range = range;
            spark.color = color;
            spark.weight = weight;
            spark.isColumn = true;

            return spark;
        }

        #endregion

        #region properties

        public virtual ExRange Location
        {
            get { return new ExRange(this.location); }
            set { this.location = new ExRange(value); }
        }

        public virtual ExRange Range
        {
            get { return new ExRange(this.range); }
            set { this.range = new ExRange(value); }
        }

        public virtual Sd.Color Color
        {
            get { return this.color; }
            set { this.color = value; }
        }

        public virtual double Weight
        {
            get { return this.weight; }
            set { this.weight = value; }
        }

        public virtual bool IsColumn
        {
            get { return this.isColumn; }
            set { this.isColumn = value; }
        }

        #endregion

        #region methods



        #endregion


    }
}
