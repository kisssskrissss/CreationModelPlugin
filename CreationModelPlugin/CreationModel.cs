using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CreationModel
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            GetLevels(doc,out var level1,out var level2);
            var walls = CreateWalls(doc,level1,level2);

            AddDoor(doc, level1, walls.First());

            var windowType= GetWindowFamilySymbol(doc);
            for (int i=1;i<walls.Count;i++)
            {
                AddWindow(doc, level1, walls[i], windowType);
            }

            AddRoof(doc, level2, walls);

            return Result.Succeeded;
        }

        private void GetLevels(Document doc, out Level level1, out Level level2)
        {
            var levelsList = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            level1 = levelsList
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();

            level2 = levelsList
               .Where(x => x.Name.Equals("Уровень 2"))
               .FirstOrDefault();
        }

        private List<Wall> CreateWalls(Document doc, Level level1,Level level2)
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc);
            transaction.Start("Построение стен");

            for (int i = 0; i < points.Count - 1; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }

            transaction.Commit();

            return walls;
        }

        private FamilySymbol GetWindowFamilySymbol(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0610 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
        }

        private void AddWindow(Document doc,Level level, Wall wall, FamilySymbol windowType)
        {
            Transaction transaction = new Transaction(doc, "Создание окна");
            transaction.Start();
            if (!windowType.IsActive)
            {
                windowType.Activate();
            }

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            point1 += new XYZ(0, 0, 0.5);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            point2 += new XYZ(0, 0, 0.5);
            XYZ point = (point1 + point2) / 2;
            doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            transaction.Commit();
        }

        private void AddDoor(Document doc,Level level,Wall wall)
        {
            Transaction transaction= new Transaction(doc,"Создание двери");
            transaction.Start();
            FamilySymbol doorType= new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            if (!doorType.IsActive)
            {
                doorType.Activate();
            }

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;
            doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            transaction.Commit();
        }

        private void AddRoof(Document doc, Level level, List<Wall> walls)
        {
           
            ElementId id = doc.GetDefaultElementTypeId(ElementTypeGroup.RoofType);
            RoofType roofType = doc.GetElement(id) as RoofType;

            double maxLenght = walls.Max(x=>x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble());

            Application application = doc.Application;
            CurveArray footPrint = application.Create.NewCurveArray();
            LocationCurve curve = walls[0].Location as LocationCurve;
            LocationCurve curve2 = walls[2].Location as LocationCurve;
            var height = walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
            XYZ p1 = curve.Curve.GetEndPoint(0)+new XYZ(0,0,height);
            XYZ p2 = curve2.Curve.GetEndPoint(0) + new XYZ(0, 0, height);
            XYZ middlePoint = ((p1 + p2) / 2)+ new XYZ(0,0,5);
            footPrint.Append(Line.CreateBound(p1, middlePoint));
            footPrint.Append(Line.CreateBound(middlePoint, p2));

            Transaction transaction = new Transaction(doc, "Создание крыши");
            transaction.Start();
            ReferencePlane plane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), doc.ActiveView);
            try
            {
                ExtrusionRoof footPrintRoof = doc.Create.NewExtrusionRoof(footPrint, plane, level, roofType, p1.X, UnitUtils.ConvertToInternalUnits(5000,UnitTypeId.Millimeters));
            }
            catch (Exception ex)
            {
                var q = ex.Message;
            }
            transaction.Commit();
        }
    }
}
