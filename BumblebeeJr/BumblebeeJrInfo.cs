using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace BumblebeeJr
{
    public class BumblebeeJrInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "BumblebeeJr";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return Properties.Resources.BumblebeeJr_24px;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "Excel NPOI plugin for Grasshopper 3d";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("e83974d5-15d1-4fbd-818d-0aa8fec21ba0");
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
                return "1.0.0.0";
            }
        }
    }

    public class BitmapJrPlusCategoryIcon : GH_AssemblyPriority
    {
        public object Properties { get; private set; }

        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon(Constants.ShortName, BumblebeeJr.Properties.Resources.BBJr_TabIcon_04);
            Instances.ComponentServer.AddCategorySymbolName(Constants.ShortName, 'B');
            return GH_LoadingInstruction.Proceed;
        }
    }
}
