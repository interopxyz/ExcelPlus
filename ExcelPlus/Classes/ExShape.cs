using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;
using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExShape
    {
        #region members

        protected ShapeTypes shapeType = ShapeTypes.Figure;

        protected enum FigureTypes { Arrow, Figure, Flow, Geometry, List, Star, Symbol };
        protected FigureTypes figureType = FigureTypes.Arrow;
        protected ShapeArrow arrow = ShapeArrow.Bent;
        protected ShapeFigure figure = ShapeFigure.Arc;
        protected ShapeFlowChart flow = ShapeFlowChart.AlternateProcess;
        protected ShapeGeometry geometry = ShapeGeometry.BlockArc;
        protected ShapeList list = ShapeList.AlternatingFlow;
        protected ShapeStar star = ShapeStar.Pt10;
        protected ShapeSymbol symbol = ShapeSymbol.Divide;

        protected Sd.Color fill = Sd.Color.Transparent;
        protected Sd.Color stroke = Constants.StartColor;
        protected double weight = 1.0;

        protected Rg.Point3d location = new Rg.Point3d();
        protected Rg.Line line = new Rg.Line();
        protected Rg.Rectangle3d boundary = new Rg.Rectangle3d();

        protected double scale = 1.0;
        protected Sd.Bitmap bitmap = new Sd.Bitmap(1,1);

        protected ArrowStyle startArrow = ArrowStyle.None;
        protected ArrowStyle endArrow = ArrowStyle.None;

        #endregion

        #region constructors

        public ExShape()
        {

        }

        public ExShape(ExShape shape)
        {
            this.shapeType = shape.shapeType;

            this.figureType = shape.figureType;
            this.arrow = shape.arrow;
            this.figure = shape.figure;
            this.flow = shape.flow;
            this.geometry = shape.geometry;
            this.list = shape.list;
            this.star = shape.star;
            this.symbol = shape.symbol;

            this.fill = shape.fill;
            this.stroke = shape.stroke;
            this.weight = shape.weight;

            this.location = shape.Location;
            this.line = shape.Line;
            this.boundary = shape.Boundary;

            this.scale = shape.Scale;

            this.bitmap = shape.Bitmap;

            this.startArrow = shape.startArrow;
            this.endArrow = shape.endArrow;
        }

        public static ExShape ConstructArrow(ShapeArrow type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Arrow;

            output.arrow = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructFigure(ShapeFigure type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Figure;

            output.figure = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructFlowChart(ShapeFlowChart type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Flow;

            output.flow = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructGeometry(ShapeGeometry type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Geometry;

            output.geometry = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructList(ShapeList type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.List;

            output.list = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructStar(ShapeStar type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Star;

            output.star = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructSymbol(ShapeSymbol type, Rg.Rectangle3d boundary)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Figure;
            output.figureType = FigureTypes.Symbol;

            output.symbol = type;
            output.Boundary = boundary;

            return output;
        }

        public static ExShape ConstructLine(Rg.Line line, ArrowStyle startSymbol, ArrowStyle endSymbol)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Line;

            output.Line = line;
            output.startArrow = startSymbol;
            output.endArrow = endSymbol;

            return output;
        }

        public static ExShape ConstructImage(Sd.Bitmap bitmap, Rg.Point3d location, double scale = 1.0)
        {
            ExShape output = new ExShape();

            output.shapeType = ShapeTypes.Image;

            output.Bitmap = bitmap;
            output.Location = location;
            output.scale = scale;

            return output;
        }

        #endregion

        #region properties

        public virtual ShapeTypes ShapeType
        {
            get { return shapeType; }
        }

        public virtual Rg.Rectangle3d Boundary
        {
            get { return new Rg.Rectangle3d(this.boundary.Plane, this.boundary.Corner(0), this.boundary.Corner(2)); }
            set { this.boundary = new Rg.Rectangle3d(value.Plane, value.Corner(0), value.Corner(2)); }
        }

        public virtual Rg.Line Line
        {
            get { return new Rg.Line(this.line.From, this.line.To); }
            set { this.line = new Rg.Line(value.From, value.To); }
        }

        public virtual Rg.Point3d Location
        {
            get { return new Rg.Point3d(this.location.X, this.location.Y, this.location.Z); }
            set { this.location = new Rg.Point3d(value.X, value.Y, value.Z); }
        }

        public virtual double Scale
        {
            get { return this.scale; }
            set { this.scale = value; }
        }

        public virtual Sd.Color Fill
        {
            get { return this.fill; }
            set { this.fill = value; }
        }

        public virtual Sd.Color Stroke
        {
            get { return this.stroke; }
            set { this.stroke = value; }
        }

        public virtual double Weight
        {
            get { return this.weight; }
            set { this.weight = value; }
        }

        public virtual Sd.Bitmap Bitmap
        {
            get { return new Sd.Bitmap(bitmap); }
            set { this.bitmap = new Sd.Bitmap(value); }
        }

        #endregion

        #region methods

        public void Apply(XL.IXLWorksheet xlSheet)
        {
            switch (this.shapeType)
            {
                case ShapeTypes.Image:
                    MemoryStream bmpStream = new MemoryStream();
                    this.bitmap.Save(bmpStream, Sd.Imaging.ImageFormat.Png);
                    XL.Drawings.IXLPicture picture = xlSheet.AddPicture(bmpStream);
                    picture.Placement = XL.Drawings.XLPicturePlacement.FreeFloating;
                    picture.Scale(this.scale);
                    picture.Left = (int)location.X;
                    picture.Top = (int)location.Y;
                    bmpStream.Dispose();
                    break;
            }

        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Shape | " + ShapeType.ToString();
        }

        #endregion

    }
}