using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace ExcelPlus.Components.Worksheet
{
    public class GH_XL_Wks_PlaceImage : GH_XL_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wks_PlaceImage class.
        /// </summary>
        public GH_XL_Wks_PlaceImage()
          : base("Place Image", "Sht Img",
              "Places a bitmap image in a worksheet",
              Constants.ShortName, Constants.SubWorkSheets)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager[0].Optional = true;
            pManager.AddGenericParameter("Image", "I", "The System.Drawing.Bitmap or image file path to display", GH_ParamAccess.item);
            pManager.AddPointParameter("Location", "L", "The location of the Shape", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Scale", "S", "A unitized scale factor (0-1) for the image", GH_ParamAccess.item, 1.0);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            DA.GetData(0, ref gooS);
            if (!gooS.TryGetWorksheet(out ExWorksheet worksheet)) return;

            IGH_Goo gooB = null;
            DA.GetData(1, ref gooB);
            if (!gooB.TryGetBitmap(out Sd.Bitmap bitmap, out string path)) return;

            Point3d location = new Point3d(0,0,0);
            DA.GetData(2, ref location);

            double scale = 1.0;
            DA.GetData(3, ref scale);

            worksheet.AddShapes(ExShape.ConstructImage(bitmap, location, scale));

            DA.SetData(0, worksheet);
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
                return Properties.Resources.XL_Shp_Image;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("de722794-758d-4ee5-bafb-6b1cd9eb785e"); }
        }
    }
}