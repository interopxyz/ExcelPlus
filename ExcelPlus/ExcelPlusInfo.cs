using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace ExcelPlus
{
    public class ExcelPlusInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "ExcelPlus";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return Properties.Resources.ExcelPlus_24;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "A Grasshopper 3d plugin for Excel File Read & Write";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("963a7dc5-77f2-4125-a995-561ee3106f83");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "David Mans";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "interopxyz@gmail.com";
            }
        }

        public override string AssemblyVersion
        {
            get
            {
                return "1.0.0.2";
            }
        }
    }

    public class ExcelPlusCategoryIcon : GH_AssemblyPriority
    {
        public object Properties { get; private set; }

        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon(Constants.ShortName, ExcelPlus.Properties.Resources.ExcelPlus_Tab_16);
            Instances.ComponentServer.AddCategorySymbolName(Constants.ShortName, 'E');
            return GH_LoadingInstruction.Proceed;
        }
    }

}
