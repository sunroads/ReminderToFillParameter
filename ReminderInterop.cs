//using Autodesk.Revit.ApplicationServices;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Events;
//using Autodesk.Revit.UI;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.IO;
//using Microsoft.Office.Interop.Excel;

//namespace ReminderToFillParameter
//{
//    class ReminderInterop : IExternalDBApplication
//    {
//        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
//        {
//            application.DocumentSynchronizingWithCentral -= SyncReminder;
//            return ExternalDBApplicationResult.Succeeded;
//        }

//        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
//        {
//            application.DocumentSynchronizingWithCentral += new EventHandler<DocumentSynchronizingWithCentralEventArgs>(SyncReminder);
//            return ExternalDBApplicationResult.Succeeded;
//        }

//        public void SyncReminder(object sender, DocumentSynchronizingWithCentralEventArgs args)
//        {
//            //TaskDialog.Show("ВНИМАНИЕ", "Проверьте заполненность параметра ***");


//            List<BuiltInCategory> categories = new List<BuiltInCategory>()
//            {
//              BuiltInCategory.OST_Views,
//              BuiltInCategory.OST_Sheets,
//              BuiltInCategory.OST_Walls,
//              BuiltInCategory.OST_PipeFitting
//            };

//            List<string> views_sheets_names = new List<string>();
//            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);
//            Document doc = args.Document;
//            FilteredElementCollector collector = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType();

//            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

//            if (xlApp == null)
//            {
//                TaskDialog.Show("Excel", "Excel is not properly installed!!");

//            }

//            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
//            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
//            object misValue = System.Reflection.Missing.Value;
//            xlWorkBook = xlApp.Workbooks.Add(misValue);
//            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

//            xlWorkSheet.Cells[1, 1] = "ID";
//            xlWorkSheet.Cells[1, 2] = "Name";
//            xlWorkSheet.Cells[2, 1] = "1";
//            xlWorkSheet.Cells[2, 2] = "One";
//            xlWorkSheet.Cells[3, 1] = "2";
//            xlWorkSheet.Cells[3, 2] = "Two";

//            xlWorkBook.SaveAs("d:\\test1.xlsx");
//            xlWorkBook.Close();


//            TaskDialog.Show("Excel", "Excel file created , you can find the file d:\\csharp-Excel.xlsx");




//        }


//    }

//}
