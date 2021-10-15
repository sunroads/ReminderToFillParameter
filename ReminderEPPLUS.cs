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
using OfficeOpenXml;
using OfficeOpenXml.Table;

//THIS IS DBAPPLICATION - (TaskDialog) after Synchronization

namespace ReminderToFillParameter
{
    class ReminderEPPLUS : IExternalDBApplication
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

            
            ExcelPackage package = new ExcelPackage();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet");

            //WHERE TO SAVE
            var file = new FileInfo(@"D:\elements.xlsx");

            List<string> views_sheets_names = new List<string>();
            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);
            Document doc = args.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType();


            using (ExcelRange Rng = ws.Cells["A1:E2"])
            {
                //Indirectly access ExcelTableCollection class
                ExcelTable table = ws.Tables.Add(Rng, "Table");


                //Set Columns position & name
                table.Columns[0].Name = "Наименование помещения";
                table.Columns[1].Name = "ID";
                table.Columns[2].Name = "Площадь";
                table.Columns[3].Name = "Объем";
                table.Columns[4].Name = "Номер комнаты";

            }


            foreach (Element el in collector)
            {
                Parameter param = el.LookupParameter("paramz");
                string parameterValue = param.AsString();

                if (!param.HasValue || parameterValue == "")

                {
                    views_sheets_names.Add(el.Name);

                }


                if (param.HasValue || parameterValue != "")

                {

                }
            }

            if (views_sheets_names.Count != 0)

            {
                int col = 1;
                for (int row = 2; row <= views_sheets_names.Count; row++)
                {
                    ws.Cells[row, col].Value = views_sheets_names[row - 2];
                }

            }

            package.SaveAs(new FileInfo(file.ToString()));
            TaskDialog.Show("Уведомление", "Элементы экспортированы, файл находится по адресу D:/elements.txt");


        }



    }


}
