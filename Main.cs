using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIExportToDWG
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            using (var ts = new Transaction(doc, "export dwg"))
            {
                ts.Start();
                ViewPlan viewPlan = new FilteredElementCollector(doc) //экспортируем в dwg план первого этажа
                                    .OfClass(typeof(ViewPlan))
                                    .Cast<ViewPlan>()
                                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan &&
                                                          v.Name.Equals("Level 1"));//печатаем только план первого этажа
                var dwgOption = new DWGExportOptions();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.dwg",
                            new List<ElementId> { viewPlan.Id }, dwgOption);
                ts.Commit();
            }
            return Result.Succeeded;
                       
        }
        public void BatchPrint(Document doc)
        {
            var sheets = new FilteredElementCollector(doc)//создаем переменную, которая собирет нам все листы существующих в данном файле, делаем через FilteredElementCollector
                        .WhereElementIsNotElementType() //будем собирать все листы, которые не являются типами, нам нужно конкретно экземпляры листов
                        .OfClass(typeof(ViewSheet))//указываем, что нам нужны элементы типа ViewSheet
                        .Cast<ViewSheet>()//преобразуем в ViewSheet
                        .ToList(); //cоздаем список листов

            //групируем листы по Title-блокам А1, А2,А3,А4 (группировка по формату), делаем это с помощью группы:
            var groupedSheets = sheets.GroupBy(sheet => doc.GetElement(new FilteredElementCollector(doc, sheet.Id)
                                                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                        .FirstElementId()).Name);//выбираем самый первый элемент

            var viewSets = new List<ViewSet>();//список набора листов

            PrintManager printManager = doc.PrintManager;//обращаемся
            printManager.SelectNewPrintDriver("PDFCreator");//задаем использованный принтер будем использовать печать PDF, это бесплатный принтер можно его скачать и использовть
            printManager.PrintRange = PrintRange.Select;

            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;//обращаемся к настройкам, для того чтобы сохранить свои листы

            foreach (var groupedSheet in groupedSheets) //проходимся по всем сгруппированным листам
            {
                if (groupedSheet.Key == null)
                    continue;

                var viewSet = new ViewSet(); //создаем пустой абор листов, который будем заполнять и который добавим в список

                var sheetsOfGroup = groupedSheet.Select(s => s).ToList();//из текущей группы собираем все листы

                foreach (var sheet in sheetsOfGroup)//проходимся по всем листам добавляя их в пустой набор
                {
                    viewSet.Insert(sheet);
                }

                viewSets.Add(viewSet);//добавили в наш набор листов, но пока он виртуально существует

                printManager.PrintRange = PrintRange.Select;//сохраняем, указываем диапазон который будем печатать, указываем что будем выбирать самостоятельно
                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;//указываем что именно будем печатать - набор листов

                using (var ts = new Transaction(doc, "Create view set"))//используем конструкцию using для транзакции
                {
                    ts.Start();
                    viewSheetSetting.SaveAs($"{groupedSheet.Key}_{Guid.NewGuid()}");

                    ts.Commit();

                }
                //создаем переменную, которая позволит понять, что нужный формат был уже выбран, по умолчанию будет false
                bool isFormatSelected = false;


                //сохраняемый набор листов будем печатать, для этого проходимся по всем форматам листов, которые существуют в выбранном принтере
                foreach (PaperSize paperSize in printManager.PaperSizes)
                {
                    if (string.Equals(groupedSheet.Key, "А4К") && //если сформированы листы форматом А4К и также в текущей выборке формат листа тоже А4
                        string.Equals(paperSize.Name, "A4"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;//задаем настройки для печати
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Portrait; //указываем формат листа портретный
                        isFormatSelected = true;

                    }

                    else if (string.Equals(groupedSheet.Key, "А3А") && //если сформированы листы форматом А4К и также в текущей выборке формат листа тоже А4
                        string.Equals(paperSize.Name, "A3"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;//задаем настройки для печати
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Landscape; //указываем формат листа альбомный
                        isFormatSelected = true;

                    }

                }

                //если вдруг формат не был найден
                if (!isFormatSelected)
                {
                    TaskDialog.Show("Ошибка", "Не найден формат");
                    return;
                }

                //если все впорядке
                printManager.CombinedFile = false;
                printManager.SubmitPrint();

            }


            return;
        }
    }
    
}
