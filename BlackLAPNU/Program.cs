using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BlackLAPNU
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var excel = new Application();
            var bookTRP = excel.Workbooks.Open(@"D:\Учёба\ТПУ\Магистерская диссертация\ИТ\Примеры\ТРП.xlsx");
            var bookTPNBU = excel.Workbooks.Open(@"D:\Учёба\ТПУ\Магистерская диссертация\ИТ\Примеры\ТПНБУ.xlsx");
            var newBookTPNBU = bookTPNBU;

            try
            {
                var sheetTRP = bookTRP.Worksheets["Настройка ПО"];
                var sheetTPNBU = newBookTPNBU.Worksheets["ПО"];

                Console.WriteLine("Идёт работа программы");
                var node = new BlankNode();
                var nodeList = new List<BlankNode>();

                var counterLO = 0;
                var counterScheme = 0;

                var startingLine = node.FindStartingLine(sheetTRP);
                var currentLine = startingLine;
                var tempGroupList = node.GetListOfTemperatureGroups(sheetTRP, startingLine);

                while (counterLO < node.GetCountOfLaunchingOrgan(sheetTRP, startingLine))
                {
                    Console.WriteLine("Количество ПО - " + node.GetCountOfLaunchingOrgan(sheetTRP, startingLine));
                    node.GetNameOfLaunchingOrgan(sheetTRP, currentLine, newBookTPNBU.Worksheets["ПО и ПС"]);
                    Console.WriteLine("counterScheme = " + counterScheme);
                    Console.WriteLine(node.GetCountOfParams(sheetTRP, currentLine, (int)ColumnNumberInTRP.LaunchingOrgan, 1));

                    while (counterScheme < node.GetCountOfParams(sheetTRP, currentLine, (int)ColumnNumberInTRP.LaunchingOrgan, 1))
                    {
                        //Console.ReadKey();
                        Console.WriteLine("ИХАЕМ!!!");
                        var groupCount = node.GetTemperatureGroupCount(sheetTRP, currentLine/*startingLine - 2*/);
                        var counterTempGroup = 0;

                        node.CheckValuesGroups(sheetTRP, currentLine, node.GetMergeLineCount(sheetTRP, currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork), sheetTRP.Cells[currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork].Value);

                        Console.WriteLine("Количество групп - " + groupCount);
                        while (counterTempGroup < groupCount)
                        {
                            node.GetNetworkScheme(bookTRP, currentLine, newBookTPNBU, tempGroupList, nodeList, groupCount);

                            var counterControlledSection = 0;
                            var newCurrentLine = currentLine;

                            while (counterControlledSection < node.GetCountOfParams(sheetTRP, currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork, 1))
                            {
                                node.GetControlledSection(sheetTRP, newCurrentLine, newBookTPNBU.Worksheets["Сечения"], node);

                                var counterInfluencingFactors = 0;

                                Console.WriteLine($"\nКоличество ВФ - {node.GetCountOfParams(sheetTRP, newCurrentLine, (int)ColumnNumberInTRP.ControlledSection, 1)}\n");

                                while (counterInfluencingFactors < node.GetCountOfParams(sheetTRP, newCurrentLine, (int)ColumnNumberInTRP.ControlledSection, 1))
                                {
                                    node.GetInfluencingFactor(sheetTRP, newCurrentLine);

                                    Console.WriteLine("\n\n" + node.InfluencingFactor + "\n\n");

                                    var mergeCells = node.GetMergeLineCount(sheetTRP, newCurrentLine, (int)ColumnNumberInTRP.InfluencingFactors);
                                    node.GetValues(sheetTRP, newCurrentLine, mergeCells);

                                    node.GetControlActions(newCurrentLine, mergeCells, sheetTRP, newBookTPNBU.Worksheets["УВ"]);

                                    newCurrentLine = newCurrentLine + mergeCells;
                                    counterInfluencingFactors++;

                                    nodeList.Add(node);

                                    //node.CheckTemperatureGroup(sheetTRP, currentLine, nodeList, tempGroupList, bookTPNBU, groupCount);

                                    var POname = node.LaunchingOrganFullName;
                                    var POopname = node.LaunchingOrganOperationName;
                                    var scheme = node.SchemeOfNetwork;
                                    var section = node.ControlledSection;
                                    var factor = node.InfluencingFactor;
                                    var temp = node.TemperatureGroup;

                                    node = new BlankNode();
                                    node.LaunchingOrganFullName = POname;
                                    node.LaunchingOrganOperationName = POopname;
                                    node.SchemeOfNetwork = scheme;
                                    node.ControlledSection = section;
                                    node.InfluencingFactor = factor;
                                    node.TemperatureGroup = temp;
                                }

                                Console.WriteLine($"\nКоличество записей после фиксации сечения - {nodeList.Count}\n");

                                counterControlledSection++;
                            }

                            var copyList = new List<BlankNode>();
                            var columnIndex = sheetTRP.Cells.Find(node.TemperatureGroup).Column;
                            Console.WriteLine("columnIndex = " + columnIndex);
                            if (sheetTRP.Cells[currentLine, columnIndex].MergeCells && groupCount > 1)
                            {
                                var cnt = 0;
                                Console.WriteLine("Количество объединенных ячеек - " + sheetTRP.Cells[currentLine, columnIndex].MergeArea.Count);
                                while (cnt < sheetTRP.Cells[currentLine, columnIndex].MergeArea.Count - 1)
                                {
                                    //Console.WriteLine($"\n\ncurrentLine = {currentLine}\n\n");
                                    var countCells = nodeList.Count - node.GetCountOfParams(sheetTRP, currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork, 2);
                                    Console.WriteLine("\ncountCells = " + countCells);
                                    Console.WriteLine("nodeListCells = " + nodeList.Count);
                                    for (int i = countCells; i < nodeList.Count; i++)
                                    {
                                        Console.WriteLine("Запись №" + i);
                                        var blank = new BlankNode();
                                        blank.LaunchingOrganFullName = nodeList[i].LaunchingOrganFullName;
                                        blank.LaunchingOrganOperationName = nodeList[i].LaunchingOrganOperationName;
                                        blank.SchemeOfNetwork = nodeList[i].SchemeOfNetwork;
                                        blank.ControlledSection = nodeList[i].ControlledSection;
                                        blank.InfluencingFactor = nodeList[i].InfluencingFactor;
                                        blank.Values = nodeList[i].Values;
                                        blank.ControlActions = nodeList[i].ControlActions;
                                        copyList.Add(blank);
                                        Console.WriteLine("\ncopyListCells = " + copyList.Count);
                                    }

                                    Console.WriteLine("Кол-во скопированных записей: " + copyList.Count);

                                    var newLine = currentLine;

                                    Console.WriteLine("BEFORE - " + nodeList.Count);

                                    var blankNode = new BlankNode();
                                    blankNode.LaunchingOrganFullName = nodeList[nodeList.Count - 1].LaunchingOrganFullName;
                                    Console.WriteLine(blankNode.LaunchingOrganFullName);
                                    blankNode.LaunchingOrganOperationName = nodeList[nodeList.Count - 1].LaunchingOrganOperationName;
                                    Console.WriteLine(blankNode.LaunchingOrganOperationName);

                                    var groupIndex = tempGroupList.IndexOf(blankNode.TemperatureGroup) + 1;
                                    blankNode.TemperatureGroup = /*nodeList[nodeList.Count - 1].TemperatureGroup*/ tempGroupList[groupIndex];
                                    Console.WriteLine(blankNode.TemperatureGroup);
                                    blankNode.ControlledSection = nodeList[nodeList.Count - 1].ControlledSection;
                                    Console.WriteLine(blankNode.ControlledSection);
                                    blankNode.Values = nodeList[nodeList.Count - 1].Values;
                                    Console.WriteLine(blankNode.Values);
                                    blankNode.ControlActions = nodeList[nodeList.Count - 1].ControlActions;
                                    Console.WriteLine(blankNode.ControlActions);
                                    Console.WriteLine("ДО: " + nodeList[nodeList.Count - 1].SchemeOfNetwork.ToString());
                                    
                                    blankNode.GetNetworkScheme(bookTRP, currentLine, newBookTPNBU, tempGroupList, nodeList, groupCount);
                                    Console.WriteLine("ПОСЛЕ: " + nodeList[nodeList.Count - 1].SchemeOfNetwork.ToString());
                                    Console.WriteLine("Новая схема: " + blankNode.SchemeOfNetwork);
                                    Console.WriteLine("\ncopyListCount = " + copyList.Count);

                                    var newCountCells = copyList.Count - node.GetCountOfParams(sheetTRP, currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork, 2);

                                    for (int i = newCountCells; i < copyList.Count; i++)
                                    {
                                        /*Console.WriteLine("Пыхали со строкой - " + newLine);
                                        copyNode.TemperatureGroup = blankNode.TemperatureGroup;
                                        copyNode.SchemeOfNetwork = blankNode.SchemeOfNetwork;
                                        Console.WriteLine("Сразу после пыхали - " + copyNode.TemperatureGroup);
                                        nodeList.Add(copyNode);*/

                                        Console.WriteLine("Пыхали со строкой - " + newLine);
                                        copyList[i].TemperatureGroup = blankNode.TemperatureGroup;
                                        copyList[i].SchemeOfNetwork = blankNode.SchemeOfNetwork;
                                        Console.WriteLine("Сразу после пыхали - " + copyList[i].TemperatureGroup);
                                        nodeList.Add(copyList[i]);
                                        Console.WriteLine("nodeListCells = " + nodeList.Count);

                                        newLine = newLine + node.GetMergeLineCount(sheetTRP, newLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork);
                                    }
                                    Console.WriteLine("AFTER - " + nodeList.Count);
                                    columnIndex++;
                                    cnt++;
                                    //Console.ReadKey();
                                }
                            }

                            Console.WriteLine("Конец - " + nodeList[nodeList.Count - 1].TemperatureGroup);
                            counterTempGroup++;
                        }

                        currentLine = currentLine + node.GetMergeLineCount(sheetTRP, currentLine, (int)ColumnNumberInTRP.SchemeOfTheNetwork);

                        counterScheme++;
                    }
                    counterScheme = 0;
                    Console.WriteLine("currentLine = " + currentLine);
                    counterLO++;
                }

                Console.WriteLine("Количество записей - " + nodeList.Count());

                var line = 4;

                foreach (var nod in nodeList)
                {
                    WriteInTPNBU(nod.LaunchingOrganFullName, sheetTPNBU, line, nod.Values.Count - 1, GetColumn(ColumnNumber.LaunchingOrganFullName));
                    WriteInTPNBU(nod.LaunchingOrganOperationName, sheetTPNBU, line, nod.Values.Count - 1, GetColumn(ColumnNumber.LaunchingOrganOperationName));
                    WriteInTPNBU(nod.ControlledSection, sheetTPNBU, line, nod.Values.Count - 1, GetColumn(ColumnNumber.ControlledSection));
                    WriteInTPNBU(nod.SchemeOfNetwork, sheetTPNBU, line, nod.Values.Count - 1, GetColumn(ColumnNumber.SchemeOfTheNetwork));
                    WriteInTPNBU(nod.InfluencingFactor, sheetTPNBU, line, nod.Values.Count - 1, GetColumn(ColumnNumber.EquipmentCondition));
                    
                    var k = 0;

                    foreach (var value in nod.Values)
                    {
                        WriteInTPNBU(value, sheetTPNBU, line + k, 0, GetColumn(ColumnNumber.Values));
                        WriteInTPNBU("", sheetTPNBU, line + k, 0, GetColumn(ColumnNumber.ControlActionAdditional));
                        k++;
                    }
                    
                    k = 0;

                    for (int i = 0; i < nod.Values.Count; i++)
                    {
                        for (int j = 0; j < 2; j++)
                        {
                            if (j == 0)
                            {
                                WriteInTPNBU(nod.ControlActions[i, j], sheetTPNBU, line + k, 0, GetColumn(ColumnNumber.ControlActionGS));
                            }
                            else
                            {
                                WriteInTPNBU(nod.ControlActions[i, j], sheetTPNBU, line + k, 0, GetColumn(ColumnNumber.ControlActionLS));
                                k++;
                            }
                        }
                    }

                    line = line + nod.Values.Count;
                }

                JoinCells(sheetTPNBU, (int)ColumnNumber.LaunchingOrganFullName, XlVAlign.xlVAlignTop);
                JoinCells(sheetTPNBU, (int)ColumnNumber.LaunchingOrganOperationName, XlVAlign.xlVAlignTop);
                sheetTPNBU.Range["A1", "I1000"].EntireRow.AutoFit();

                excel.DisplayAlerts = false;
                newBookTPNBU.SaveAs(@"D:\Учёба\ТПУ\Магистерская диссертация\ИТ\Примеры\ТПНБУnew.xlsx");
                newBookTPNBU.Close();
                bookTRP.Close();

                Console.WriteLine("Работа завершена");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n{exception.Message + exception.StackTrace}");
                newBookTPNBU.Close();
                bookTRP.Close();

                Console.WriteLine("Работа завершена");
                Console.ReadKey();
            }

            void WriteInTPNBU(object text, Worksheet sheet, int row, int mergeLineCount, string column)
            {
                sheet.Cells[row, column].Value = text;
                sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].Merge(Type.Missing);
                sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].HorizontalAlignment = 
                    XlHAlign.xlHAlignCenter;
                sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].VerticalAlignment = 
                    XlVAlign.xlVAlignCenter;
                sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].WrapText = true;
                //sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].EntireRow.AutoFit();
                sheet.Range[$"{column}{row}", $"{column}{row + mergeLineCount}"].BorderAround2
                    (XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic);
            }

            void CellsBorderStyle(Worksheet sheet)
            {
                var lineCount = sheet.Rows.Count;

                /*sheet.Range[$"A2", $"I{lineCount}"].BorderAround2(XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic);*/

                foreach (var r in newBookTPNBU.Worksheets["ПО"].Range("A2", $"I{lineCount}").Cells)
                {
                    r.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin,
                        XlColorIndex.xlColorIndexAutomatic);
                }
            }

            
            void JoinCells(Worksheet sheet, int column, XlVAlign verticalAlignment)
            {
                excel.DisplayAlerts = false;

                var rowsCount = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                    false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                var startCount = sheet.Cells.Find("Схема сети").Row + 1;

                for (int i = startCount; i <= rowsCount;)
                {
                    int j;

                    while (true)
                    {
                        j = i;
                        Console.WriteLine("\nЗначение j = " + j);
                        var value = sheet.Cells[i, column].Value;

                        if (sheet.Cells[i, column].MergeCells)
                        {
                            Console.WriteLine($"Cells[{i}, {column}] объединена");
                            j = j + sheet.Cells[i, column].MergeArea.Count;
                            Console.WriteLine("Новое значение j = " + j);
                        }
                        else
                        {
                            Console.WriteLine($"\nCells[{i}, {column}] объединена");
                            j++;
                            Console.WriteLine("Новое значение j = " + j);
                        }

                        if (value == sheet.Cells[j, column].Value)
                        {
                            Console.WriteLine("Значения совпадают");
                            sheet.Range[sheet.Cells[i, column], sheet.Cells[j, column]].Merge(Type.Missing);
                            sheet.Range[sheet.Cells[i, column], sheet.Cells[j, column]].VerticalAlignment = verticalAlignment;
                        }
                        else
                        {
                            Console.WriteLine("Значения не совпадают");
                            break;
                        }
                    }

                    i = j;

                    /*var j = i;
                    var value = sheet.Cells[i, column].Value;

                    if (sheet.Cells[i, column].MergeCells)
                    {
                        j = j + sheet.Cells[i, column].MergeArea;
                    }
                    else
                    {
                        j++;
                    }

                    if (value == sheet.Cells[j, column].Value)
                    {
                        sheet.Range[sheet.Cells[i, column], sheet.Cells[j, column]].Merge(Type.Missing);
                    }

                    i = j;*/
                }
            }

            string GetColumn(ColumnNumber column)
            {
                var columnLetter = "";

                switch (column)
                {
                    case ColumnNumber.ControlledSection:
                        columnLetter = "A";
                        break;
                    case ColumnNumber.LaunchingOrganFullName:
                        columnLetter = "B";
                        break;
                    case ColumnNumber.LaunchingOrganOperationName:
                        columnLetter = "C";
                        break;
                    case ColumnNumber.SchemeOfTheNetwork:
                        columnLetter = "D";
                        break;
                    case ColumnNumber.EquipmentCondition:
                        columnLetter = "E";
                        break;
                    case ColumnNumber.Values:
                        columnLetter = "F";
                        break;
                    case ColumnNumber.ControlActionGS:
                        columnLetter = "G";
                        break;
                    case ColumnNumber.ControlActionLS:
                        columnLetter = "H";
                        break;
                    case ColumnNumber.ControlActionAdditional:
                        columnLetter = "I";
                        break;
                }

                return columnLetter;
            }

            void NodeRefresh(BlankNode node)
            {
                var POname = node.LaunchingOrganFullName;
                var POopname = node.LaunchingOrganOperationName;
                var scheme = node.SchemeOfNetwork;
                var section = node.ControlledSection;
                var temp = node.TemperatureGroup;

                node = new BlankNode();
                node.LaunchingOrganFullName = POname;
                node.LaunchingOrganOperationName = POopname;
                node.SchemeOfNetwork = scheme;
                node.ControlledSection = section;
                node.TemperatureGroup = temp;
            }
        }

        enum ColumnNumber
        {
            ControlledSection = 1,
            LaunchingOrganFullName = 2,
            LaunchingOrganOperationName = 3,
            SchemeOfTheNetwork = 4,
            EquipmentCondition = 5,
            Values = 6,
            ControlActionGS = 7,
            ControlActionLS = 8,
            ControlActionAdditional = 9
        }

        enum ColumnNumberInTRP
        {
            Number = 1,
            LaunchingOrgan = 2,
            SchemeOfTheNetwork = 3,
            ControlledSection = 4,
            InfluencingFactors = 5,
            Values = 6,
            ControlActionGS = 10,
            ControlActionLS = 11
        }
    }
}
