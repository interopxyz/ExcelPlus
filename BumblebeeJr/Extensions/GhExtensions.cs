using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rg = Rhino.Geometry;

namespace BumblebeeJr
{
    public static class GhExtensions
    {

        public static bool TryGetCell(this IGH_Goo input, ref XLCell cell)
        {
            XLCell cl;
            GH_Point gpt;
            Rg.Point3d p3d;
            Rg.Point2d p2d;
            string address;
            Rg.Interval domain;

            if (input == null) return false;

            if (input.CastTo<XLCell>(out cl))
            {
                cell = new XLCell(cl);
                return true;
            }
            else if (input.CastTo<GH_Point>(out gpt))
            {
                cell = new XLCell((int)gpt.Value.X, (int)gpt.Value.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point3d>(out p3d))
            {
                cell = new XLCell((int)p3d.X, (int)p3d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point2d>(out p2d))
            {
                cell = new XLCell((int)p2d.X, (int)p2d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Interval>(out domain))
            {
                cell = new XLCell((int)domain.T0, (int)domain.T1);
                return true;
            }
            else if (input.CastTo<string>(out address))
            {
                cell = new XLCell(address);
                return true;
            }
            else
            {
                cell = null;
                return false;
            }
        }

    }
}
