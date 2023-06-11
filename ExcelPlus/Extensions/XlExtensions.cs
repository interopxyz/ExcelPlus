using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using XL = ClosedXML.Excel;


namespace ExcelPlus
{
    public static class XlExtensions
    {
        public static XL.XLColor ToExcel(this Sd.Color input)
        {
            return XL.XLColor.FromColor(input);
        }
        public static Sd.Color ToColor(this XL.XLColor input)
        {
            return Sd.Color.FromArgb(input.Color.A,input.Color.R,input.Color.G,input.Color.B);
        }

        public static XL.XLAlignmentHorizontalValues ToExcelHAlign(this Justifications input)
        {
            switch (input)
            {
                case Justifications.BottomMiddle:
                case Justifications.CenterMiddle:
                case Justifications.TopMiddle:
                    return XL.XLAlignmentHorizontalValues.Center;
                case Justifications.BottomRight:
                case Justifications.CenterRight:
                case Justifications.TopRight:
                    return XL.XLAlignmentHorizontalValues.Center;
                default:
                    return XL.XLAlignmentHorizontalValues.Left;
            }
        }

        public static XL.XLAlignmentVerticalValues ToExcelVAlign(this Justifications input)
        {
            switch (input)
            {
                case Justifications.TopLeft:
                case Justifications.TopMiddle:
                case Justifications.TopRight:
                    return XL.XLAlignmentVerticalValues.Top;
                case Justifications.BottomLeft:
                case Justifications.BottomMiddle:
                case Justifications.BottomRight:
                    return XL.XLAlignmentVerticalValues.Bottom;
                default:
                    return XL.XLAlignmentVerticalValues.Center;
            }
        }

        public static Justifications ToJustification( this XL.XLAlignmentHorizontalValues horizontal, XL.XLAlignmentVerticalValues vertical)
        {
            switch (horizontal)
            {
                default:
                    switch (vertical)
                    {
                        default:
                            return Justifications.CenterLeft;
                        case XL.XLAlignmentVerticalValues.Bottom:
                            return Justifications.BottomLeft;
                        case XL.XLAlignmentVerticalValues.Top:
                            return Justifications.TopLeft;
                    }
                case XL.XLAlignmentHorizontalValues.Right:
                    switch (vertical)
                    {
                        default:
                            return Justifications.CenterRight;
                        case XL.XLAlignmentVerticalValues.Bottom:
                            return Justifications.BottomRight;
                        case XL.XLAlignmentVerticalValues.Top:
                            return Justifications.TopRight;
                    }
                case XL.XLAlignmentHorizontalValues.Center:
                    switch (vertical)
                    {
                        default:
                            return Justifications.CenterMiddle;
                        case XL.XLAlignmentVerticalValues.Bottom:
                            return Justifications.BottomMiddle;
                        case XL.XLAlignmentVerticalValues.Top:
                            return Justifications.TopMiddle;
                    }
            }
        }

        public static XL.XLBorderStyleValues ToExcel(this LineTypes input)
        {
            switch (input)
            {
                default:
                    return XL.XLBorderStyleValues.Medium;
                case LineTypes.DashDot:
                    return XL.XLBorderStyleValues.DashDot;
                case LineTypes.DashDotDot:
                    return XL.XLBorderStyleValues.DashDotDot;
                case LineTypes.Dashed:
                    return XL.XLBorderStyleValues.Dashed;
                case LineTypes.Dotted:
                    return XL.XLBorderStyleValues.Dotted;
                case LineTypes.Double:
                    return XL.XLBorderStyleValues.Double;
                case LineTypes.Hair:
                    return XL.XLBorderStyleValues.Hair;
                case LineTypes.MediumDashDot:
                    return XL.XLBorderStyleValues.MediumDashDot;
                case LineTypes.MediumDashDotDot:
                    return XL.XLBorderStyleValues.MediumDashDotDot;
                case LineTypes.MediumDashed:
                    return XL.XLBorderStyleValues.MediumDashed;
                case LineTypes.None:
                    return XL.XLBorderStyleValues.None;
                case LineTypes.SlantDashDot:
                    return XL.XLBorderStyleValues.SlantDashDot;
                case LineTypes.Thick:
                    return XL.XLBorderStyleValues.Thick;
                case LineTypes.Thin:
                    return XL.XLBorderStyleValues.Thin;
            }
        }

        public static LineTypes ToPlus(this XL.XLBorderStyleValues input)
        {
            switch (input)
            {
                default:
                    return LineTypes.Medium;
                case XL.XLBorderStyleValues.DashDot:
                    return LineTypes.DashDot;
                case XL.XLBorderStyleValues.DashDotDot:
                    return LineTypes.DashDotDot;
                case XL.XLBorderStyleValues.Dashed:
                    return LineTypes.Dashed;
                case XL.XLBorderStyleValues.Dotted:
                    return LineTypes.Dotted;
                case XL.XLBorderStyleValues.Double:
                    return LineTypes.Double;
                case XL.XLBorderStyleValues.Hair:
                    return LineTypes.Hair;
                case XL.XLBorderStyleValues.MediumDashDot:
                    return LineTypes.MediumDashDot;
                case XL.XLBorderStyleValues.MediumDashDotDot:
                    return LineTypes.MediumDashDotDot;
                case XL.XLBorderStyleValues.MediumDashed:
                    return LineTypes.MediumDashed;
                case XL.XLBorderStyleValues.None:
                    return LineTypes.None;
                case XL.XLBorderStyleValues.SlantDashDot:
                    return LineTypes.SlantDashDot;
                case XL.XLBorderStyleValues.Thick:
                    return LineTypes.Thick;
                case XL.XLBorderStyleValues.Thin:
                    return LineTypes.Thin;
            }
        }

    }
}
