using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;

namespace ReminderToFillParameter
{
    public class ReminderTxt : IExternalDBApplication
    {

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            application.DocumentSynchronizingWithCentral -= SyncReminder;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            application.DocumentSynchronizingWithCentral += new EventHandler<DocumentSynchronizingWithCentralEventArgs>(SyncReminder);

            return ExternalDBApplicationResult.Succeeded;
        }

        public void SyncReminder(object sender, DocumentSynchronizingWithCentralEventArgs args)
        {
            
            List<BuiltInCategory> categories = new List<BuiltInCategory>()
            {
              BuiltInCategory.OST_Views,
              BuiltInCategory.OST_Sheets,
              BuiltInCategory.OST_Walls,
              BuiltInCategory.OST_PipeFitting
            };

            List<string> views_sheets_names = new List<string>();
            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);
            Document doc = args.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType();

            foreach (Element el in collector)
            {
                Autodesk.Revit.DB.Parameter param = el.LookupParameter("paramz");
                string parameterValue = param.AsString();

                if (!param.HasValue || parameterValue == "")

                {
                    views_sheets_names.Add(el.Name + " ID: " + el.Id);

                }


                if (param.HasValue || parameterValue != "")

                {

                }
            }

            if (views_sheets_names.Count != 0)

            {
                string Elements = string.Join(", " + System.Environment.NewLine, views_sheets_names);


                File.WriteAllText(@"D:\elements.txt", Elements);
                TaskDialog.Show("Уведомление", "Элементы экспортированы, файл находится по адресу D:/elements.txt");

            }



        }


    }


}