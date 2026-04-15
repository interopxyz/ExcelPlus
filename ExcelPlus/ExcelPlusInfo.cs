using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace ExcelPlus
{
    public class ExcelPlusInfo : GH_AssemblyInfo
    {
        public override string Name => "ExcelPlus";
        public override Bitmap Icon => Properties.Resources.ExcelPlus_24;
        public override string Description => "Grasshopper 3d plugin for Excel File Read & Write";
        public override Guid Id => new Guid("963a7dc5-77f2-4125-a995-561ee3106f83");
        public override string AuthorName => "David Mans";
        public override string AuthorContact => "interopxyz@gmail.com";
        public override string AssemblyVersion => GetType().Assembly.GetName().Version.ToString();
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
