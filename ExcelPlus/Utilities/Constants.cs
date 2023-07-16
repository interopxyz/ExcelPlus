using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace ExcelPlus
{
    public class Constants
    {

        #region naming

        public static string UniqueName
        {
            get { return ShortName+"_" + DateTime.UtcNow.ToString("yyyy-dd-M_HH-mm-ss"); }
        }

        public static string LongName
        {
            get { return ShortName + " v" + Major + "." + Minor; }
        }

        public static string ShortName
        {
            get { return "ExcelPlus"; }
        }

        private static string Minor
        {
            get { return typeof(Constants).Assembly.GetName().Version.Minor.ToString(); }
        }
        private static string Major
        {
            get { return typeof(Constants).Assembly.GetName().Version.Major.ToString(); }
        }

        public static string SubRange
        {
            get { return "Range"; }
        }

        public static string SubObject
        {
            get { return "Shape"; }
        }

        public static string SubCell
        {
            get { return "Cell"; }
        }

        public static string SubApp
        {
            get { return "App"; }
        }

        public static string SubData
        {
            get { return "Data"; }
        }

        public static string SubFormat
        {
            get { return "Format"; }
        }

        public static string SubChart
        {
            get { return "Chart"; }
        }

        public static string SubAnalysis
        {
            get { return "Analysis"; }
        }

        public static string SubWorkBooks
        {
            get { return "Workbook"; }
        }

        public static string SubWorkSheets
        {
            get { return "Worksheet"; }
        }

        public static Descriptor Location
        {
            get { return new Descriptor("Cell Location", "L", "Location of the cell as:" 
                + Environment.NewLine +"String address ex. 'A1'"
                + Environment.NewLine+ "Domain interval ex. 1 to 1"
                + Environment.NewLine + "Point3d ex. {1,1,0}",
                 "Cell location string", "Cell location string"); }
        }

        public static Descriptor Format
        {
            get {
                return new Descriptor("Format", "F", "MS Office Number Format"
              + Environment.NewLine + "Examples (\"General\", \"hh: mm:ss\", \"$#,##0.0\" )"
              + Environment.NewLine + "https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-us&rs=en-us&ad=us"
                 ,"Cell format string", "Cell format string");
            }
        }

        public static Descriptor App
        {
            get { return new Descriptor("Excel Application", "App", "A Worksheet, Workbook, Range Object, Excel Application, or Text Workbook Name", "An Excel Application Object", "Excel Application Objects"); }
        }

        public static Descriptor Workbook
        {
            get { return new Descriptor("Workbook", "Wbk", "A Worksheet, Workbook, Range, or Workbook Name", "An Excel Workbook Object", "Excel Workbook Objects"); }
        }

        public static Descriptor Worksheet
        {
            get { return new Descriptor("Worksheet", "Wks", "A Worksheet, Workbook, Range, or Worksheet Name", "An Excel Worksheet Object", "Excel Worksheet Objects"); }
        }

        public static Descriptor Range
        {
            get { return new Descriptor("Range", "Rng", "A Range or Address (ex. A1:B1)", "An Excel Range Object", "Excel Range Objects"); }
        }

        public static Descriptor Cell
        {
            get { return new Descriptor("Cell", "Cel", "A Cell or Address (ex. A1)", "An Excel Cell Object", "Excel Cell Objects"); }
        }

        public static Descriptor Shape
        {
            get { return new Descriptor("Shape", "Shp", "A Smart Art, Control, or Graphical Shape", "An Excel Shape Object", "Excel Shape Objects"); }
        }

        public static Descriptor Chart
        {
            get { return new Descriptor("Chart", "Cht", "A Chart object", "A Chart object", "Chart objects"); }
        }

        public static Descriptor ShapeType
        {
            get { return new Descriptor("Shape Type", "T", "The Shape type", "The Shape type", "The Shape types"); }
        }

        public static Descriptor Boundary
        {
            get { return new Descriptor("Boundary", "B", "The Shape bounding rectangle", "The Shape bounding rectangle", "The Shape bounding rectangles"); }
        }

        public static Descriptor Activate
        {
            get { return new Descriptor("Activate", "_A", "If true, the component will be activated", "The active status of the component", "The active statuses of the component"); }
        }

        public static Descriptor DataSet
        {
            get { return new Descriptor("Data Set", "Dst", "A Data Set object", "A Data Set object", "Data Set objects"); }
        }

        #endregion

        #region color

        public static Sd.Color StartColor
        {
            get { return Sd.Color.FromArgb(99, 190, 123); }
        }

        public static Sd.Color MidColor
        {
            get { return Sd.Color.FromArgb(255, 235, 132); }
        }

        public static Sd.Color EndColor
        {
            get { return Sd.Color.FromArgb(248, 105, 107); }
        }

        #endregion
    }

    public class Descriptor
    {
        private string name = string.Empty;
        private string nickname = string.Empty;
        private string input = string.Empty;
        private string output = string.Empty;
        private string outputs = string.Empty;

        public Descriptor(string name, string nickname, string input, string output, string outputs)
        {
            this.name = name;
            this.nickname = nickname;
            this.input = input;
            this.output = output;
            this.outputs = outputs;
        }

        public virtual string Name
        {
            get { return name; }
        }

        public virtual string NickName
        {
            get { return nickname; }
        }

        public virtual string Input
        {
            get { return input; }
        }

        public virtual string Output
        {
            get { return output; }
        }

        public virtual string Outputs
        {
            get { return outputs; }
        }
    }
}
