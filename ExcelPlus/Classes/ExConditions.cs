using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExCondition
    {
        #region members

        protected ConditionalTypes type = ConditionalTypes.None;

        protected ValueCondition valueType = ValueCondition.Equal;

        protected AverageCondition averageType = AverageCondition.AboveAverage;

        protected TextCondition textType = TextCondition.Contains;

        protected Sd.Color color1 = Constants.StartColor;
        protected Sd.Color color2 = Constants.MidColor;
        protected Sd.Color color3 = Constants.EndColor;

        protected string text = "";
        protected double value1 = 0.5;
        protected double value2 = 1.0;
        protected int count = 10;
        protected bool toggle = false;

        #endregion

        #region constructors

        public ExCondition()
        {

        }

        public ExCondition(ExCondition condition)
        {
            this.type = condition.type;
            this.valueType = condition.valueType;
            this.averageType = condition.averageType;

            this.color1 = condition.color1;
            this.color2 = condition.color2;
            this.color3 = condition.color3;

            this.text = condition.text;
            this.value1 = condition.value1;
            this.value2 = condition.value2;
            this.count = condition.count;
            this.toggle = condition.toggle;
        }

        public static ExCondition CreateUniqueCondition(bool flip, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Unique;
            condition.toggle = flip;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateTopCountCondition(int count, bool flip, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Count;
            condition.count = count;
            condition.toggle = flip;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateTextCondition(string text, TextCondition type, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Percent;
            condition.text = text;
            condition.textType = type;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateValueCondition(double value, ValueCondition type, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Percent;
            condition.value1 = value;
            condition.valueType = type;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateTopPercentCondition(double value, bool flip, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Percent;
            condition.value1 = value;
            condition.toggle = flip;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateBetweenCondition(double value1,double value2, bool flip, Sd.Color color1)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Between;
            condition.value1 = value1;
            condition.value2 = value2;
            condition.toggle = flip;
            condition.color1 = color1;

            return new ExCondition();
        }

        public static ExCondition CreateScalarCondition(Sd.Color color1, Sd.Color color2)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Scale;
            condition.color1 = color1;
            condition.color2 = color2;

            return new ExCondition();
        }

        public static ExCondition CreateScalarCondition(Sd.Color color1, Sd.Color color2, double param, Sd.Color color3)
        {
            ExCondition condition = new ExCondition();
            condition.type = ConditionalTypes.Scale;
            condition.value1 = param;
            condition.color1 = color1;
            condition.color2 = color2;
            condition.color3 = color3;

            return new ExCondition();
        }

        #endregion

        #region properties


        #endregion

        #region methods

        public void ApplyCondition(XL.IXLRange xlRange)
        {
            XL.IXLConditionalFormat format = xlRange.AddConditionalFormat();

            switch(type)
            {
                default:

                    break;
                case ConditionalTypes.Scale:
                    xlRange
                        .AddConditionalFormat()
                        .ColorScale()
                        .Minimum(XL.XLCFContentType.Percentile, 0, this.color1.ToExcel())
                        .Midpoint(XL.XLCFContentType.Percentile, 50, this.color2.ToExcel())
                        .Maximum(XL.XLCFContentType.Percentile, 100, this.color3.ToExcel());
                    break;
                case ConditionalTypes.Blanks:
                    if (this.toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenNotBlank().Fill.BackgroundColor = color1.ToExcel();
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsBlank().Fill.BackgroundColor = color1.ToExcel();
                    }
                    break;
                case ConditionalTypes.Value:
                    switch (this.valueType)
                    {
                        case ValueCondition.Equal:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEquals(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case ValueCondition.NotEqual:
                            xlRange
                            .AddConditionalFormat()
                            .WhenNotEquals(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case ValueCondition.Greater:
                            xlRange
                            .AddConditionalFormat()
                            .WhenGreaterThan(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case ValueCondition.GreaterEqual:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEqualOrGreaterThan(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case ValueCondition.Less:
                            xlRange
                            .AddConditionalFormat()
                            .WhenLessThan(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case ValueCondition.LessEqual:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEqualOrLessThan(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                    }
                    break;
                case ConditionalTypes.Text:
                    switch (this.textType)
                    {
                        case TextCondition.Begins:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEqualOrLessThan(value1).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case TextCondition.Contains:
                            xlRange
                            .AddConditionalFormat()
                            .WhenContains(text).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case TextCondition.NotContains:
                            xlRange
                            .AddConditionalFormat()
                            .WhenNotContains(text).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case TextCondition.Ends:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEndsWith(text).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case TextCondition.Equal:
                            xlRange
                            .AddConditionalFormat()
                            .WhenEquals(text).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                        case TextCondition.NotEqual:
                            xlRange
                            .AddConditionalFormat()
                            .WhenNotEquals(text).Fill.BackgroundColor = this.color1.ToExcel();
                            break;
                    }
                    break;
                case ConditionalTypes.Between:
                    if (toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenNotBetween(Math.Min(this.value1,this.value2),Math.Max(this.value1,this.value2)).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenBetween(Math.Min(this.value1, this.value2), Math.Max(this.value1, this.value2)).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    break;
                case ConditionalTypes.Unique:
                    if (toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsUnique().Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsDuplicate().Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    break;
                case ConditionalTypes.Count:
                    if (toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsBottom(this.count, XL.XLTopBottomType.Items).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsTop(this.count, XL.XLTopBottomType.Items).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    break;
                case ConditionalTypes.Percent:
                    if (toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsBottom((int)(this.value1*100.0), XL.XLTopBottomType.Percent).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .WhenIsTop((int)(this.value1 * 100.0), XL.XLTopBottomType.Percent).Fill.BackgroundColor = this.color1.ToExcel();
                    }
                    break;
                case ConditionalTypes.Bars:
                    if (toggle)
                    {
                        xlRange
                        .AddConditionalFormat()
                        .DataBar(this.color1.ToExcel(),this.color2.ToExcel());
                    }
                    else
                    {
                        xlRange
                        .AddConditionalFormat()
                        .DataBar(this.color1.ToExcel());
                    }
                    break;
            }
        }

        #endregion

    }

}
