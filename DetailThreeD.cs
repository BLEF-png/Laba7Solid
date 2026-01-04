using SolidWorks.Interop.cosworks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.LinkLabel;

namespace Lab5_Kaluzhny
{
    public class DetailThreeD
    {
        private double size = 0.1;
        private double x = 0, y = 0, z = 0;
        private double height = 3.2;
        private double width = 2.2;
        private double deep = 1;

        private void selectPlane(ModelDoc2 md, string name)//select a plane
        {
            string obj = "PLANE";
            md.Extension.SelectByID2(name, obj, 0, 0, 0, false, 0, null, 0);
        }

        private Feature featureExtrusion(ModelDoc2 md, double size)
        {
            bool dir = false;
            return md.FeatureManager.FeatureExtrusion2(true, false, dir, (int)swEndConditions_e.swEndCondBlind,
                (int)swEndConditions_e.swEndCondBlind, size, 0, false, false, false, false, 0, 0, false, false, false,
                false, true, true, true, 0, 0, false);
        }

        public Feature DrawStep1(SketchManager sm, ModelDoc2 md)
        {
            string top = "Top Plane";
            selectPlane(md, top);

            md.SketchManager.InsertSketch(false);

            SketchPoint pointRectTop = md.SketchManager.CreatePoint(x - width / 2, y + height / 2, z);
            SketchPoint pointRectTopRighter = md.SketchManager.CreatePoint(x + width / 2, y + height / 2, z);
            SketchPoint pointRectBottom = md.SketchManager.CreatePoint(x + width / 2, y - height / 2, z);


            md.SketchManager.Create3PointCornerRectangle(pointRectTop.X, pointRectTop.Y, pointRectTop.Z,
                                                                         pointRectTopRighter.X, pointRectTopRighter.Y, pointRectTopRighter.Z,
                                                                         pointRectBottom.X, pointRectBottom.Y, pointRectBottom.Z);

            pointRectTop.Select(false);
            pointRectTopRighter.Select(true);
            md.IAddHorizontalDimension2(pointRectTop.X - (pointRectTop.X - pointRectTopRighter.X) / 2, z, pointRectTop.Y + size);
            md.ClearSelection();

            pointRectTop.Select(false);
            pointRectBottom.Select(true);
            md.IAddVerticalDimension2(pointRectTop.X - size, y, pointRectBottom.Y + (pointRectTop.Y - pointRectBottom.Y) / 2);


            var feature = featureExtrusion(md, deep);
            md.ClearSelection();
            return feature;
        }
    }
}
    