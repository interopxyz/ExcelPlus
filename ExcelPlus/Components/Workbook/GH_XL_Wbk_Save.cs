using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelPlus.Components
{
    public class GH_XL_Wbk_Save : GH_XL_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_Save class.
        /// </summary>
        public GH_XL_Wbk_Save()
          : base("Save Workbook", "Save Wbk",
              "Save a Workbook to an Excel file",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.senary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddTextParameter("Folder Path", "F", "The path to the workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("File Name", "N", "The Workbook name", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Extensions", "E", "The file type extension", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (Extensions value in Enum.GetValues(typeof(Extensions)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("FilePath", "P", "The full filepath to the saved file", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            bool activate = false;
            DA.GetData(4, ref activate);
            if (activate)
            {
                IGH_Goo gooB = null;
                DA.GetData(0, ref gooB);
                if (!gooB.TryGetWorkbook(out ExWorkbook workbook)) return;

                string path = "C:\\Users\\Public\\Documents\\";
                bool hasPath = DA.GetData(1, ref path);

                if (!hasPath)
                {
                    if (this.OnPingDocument().FilePath != null)
                    {
                        path = Path.GetDirectoryName(this.OnPingDocument().FilePath) + "\\";
                    }
                }

                if (!Directory.Exists(path))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The file provided path does not exist. Please verify this is a valid file path.");
                    return;
                }

                string name = Constants.UniqueName;
                DA.GetData(2, ref name);

                int ext = 0;
                DA.GetData(3, ref ext);

                string filepath = workbook.Save(path,name, (Extensions)ext);

                DA.SetData(0, workbook);
                DA.SetData(1, filepath);
            }


        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.XL_Wbk_Save;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("2083eddc-f00d-4773-8d49-81875571b5ed"); }
        }
    }
}