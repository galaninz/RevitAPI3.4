using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI3._4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categorySet = new CategorySet();
            //categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Walls));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));
            //categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Doors));
            //categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Windows));

            using (Transaction ts = new Transaction(doc, "Add parameter"))
            {
                ts.Start();
                CreateShareParameter(uiapp.Application, doc, "Наименование", categorySet, BuiltInParameterGroup.PG_GEOMETRY, true);
                ts.Commit();
            }

            //var selectedRef = uidoc.Selection.PickObject(ObjectType.Element, "Выберете элемент");
            //var selectedElement = doc.GetElement(selectedRef);

            //IList<Reference> selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, new PipeFilter(), "Выберете элементы труб");
            //var pipeList = new List<Pipe>();
            //var len = new List<double>();
            //double sum = 0;
            //string Length = string.Empty;

            var pipeList = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            //foreach (var selectedElement in selectedRef)
            //{
            //    Pipe oPipe = doc.GetElement(selectedElement) as Pipe;
            //    pipeList.Add(oPipe);
            //}

            foreach (Pipe pipe in pipeList)
            {
                Parameter inDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                Parameter outDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                string innerDiameter = UnitUtils.ConvertFromInternalUnits(inDiameter.AsDouble(), UnitTypeId.Millimeters).ToString();
                string outerDiameter = UnitUtils.ConvertFromInternalUnits(outDiameter.AsDouble(), UnitTypeId.Millimeters).ToString();

                using (Transaction ts = new Transaction(doc, "Set parameters"))
                {
                    ts.Start();
                    Parameter commentParameter = pipe.LookupParameter("Наименование");
                    commentParameter.Set($"Труба {innerDiameter}/{outerDiameter}");
                    ts.Commit();
                }
            }

            string Diameter = "Диаметры труб успешно заполнены";

            TaskDialog.Show("Selection", Diameter);

            return Result.Succeeded;
        }

        private void CreateShareParameter(Application application,
            Document doc, string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile defFile = application.OpenSharedParameterFile();
            if (defFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = defFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        } 
    }
}
