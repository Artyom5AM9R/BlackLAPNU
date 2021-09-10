using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BlackLAPNU
{
    public class BlankNode
    {
        /// <summary>
        /// Полное наименование пускового органа
        /// </summary>
        public string LaunchingOrganFullName { get; /*private */set; }

        /// <summary>
        /// Оперативное наименование пускового органа
        /// </summary>
        public string LaunchingOrganOperationName { get;/* private */set; }

        /// <summary>
        /// Контролируемое сечение
        /// </summary>
        public string ControlledSection { get; /*private*/ set; }

        /// <summary>
        /// Влияющий фактор
        /// </summary>
        public string InfluencingFactor { get; /*private*/ set; }

        /// <summary>
        /// Схема сети
        /// </summary>
        public string SchemeOfNetwork { get; /*private */set; }

        /// <summary>
        /// Группа уставок по температуре наружного воздуха
        /// </summary>
        public string TemperatureGroup { get; /*private*/ set; }

        /// <summary>
        /// Уставки КПР
        /// </summary>
        public List<int> Values { get; /*private */set; } = new List<int>();

        /// <summary>
        /// Управляющие воздействия
        /// </summary>
        public string[,] ControlActions { get; /*private */set; }


        /// <summary>
        /// Метод для добавления названий пускового органа
        /// </summary>
        /// <param name="worksheetTRP"></param>
        /// <param name="startingIndex"></param>
        /// <param name="worksheetTPNBU"></param>
        public void GetNameOfLaunchingOrgan(Worksheet worksheetTRP, int startingIndex, Worksheet worksheetTPNBU)
        {
            LaunchingOrganFullName = worksheetTRP.Cells[startingIndex, 2].Value;
            int indexOfEnd = LaunchingOrganFullName.IndexOf("(") - 1;
            Console.WriteLine("LaunchingOrganFullName - " + LaunchingOrganFullName);
            for (int i = 1; i <= worksheetTPNBU.Rows.Count; i++)
            {
                if (worksheetTPNBU.Cells[i, 1].Value == LaunchingOrganFullName.Substring(0, indexOfEnd))
                {
                    LaunchingOrganOperationName = worksheetTPNBU.Cells[i, 2].Value;
                    Console.WriteLine("LaunchingOrganOperationName - " + LaunchingOrganOperationName);
                    break;
                }
            }         
        }

        /// <summary>
        /// Метод для добавления контролируемого сечения
        /// </summary>
        /// <param name="worksheetTRP"></param>
        /// <param name="startingIndex"></param>
        /// <param name="worksheetTPNBU"></param>
        public void GetControlledSection(Worksheet worksheetTRP, int startingIndex, Worksheet worksheetTPNBU, BlankNode node) 
        {
            string controlledSectionName = worksheetTRP.Cells[startingIndex, 4].Value;

            for (int i = 3; i <= worksheetTPNBU.Rows.Count; i++)
            {
                if (worksheetTPNBU.Cells[i, 1].Value == controlledSectionName)
                {
                    ControlledSection = worksheetTPNBU.Cells[i, 2].Value;
                    break;
                }
            }
        }

        /// <summary>
        /// Метод для проверки корректности заполнения групп уставок в ТРП
        /// </summary>
        public void CheckValuesGroups(Worksheet sheet, int startLine, int countOfMergeCellsInNetworkScheme, string scheme)
        {
            var exception = new Exception("Неправильное формирование групп уставок для " +
                                "следующей схемы сети - " + scheme);

            for (int j = 6; j <= 9;)
            {
                var MergeCellsVolume = sheet.Cells[startLine, j].MergeArea.Count;

                if (MergeCellsVolume == 0)
                {
                    for (int i = startLine; i < startLine + countOfMergeCellsInNetworkScheme; i++)
                    {
                        if (sheet.Cells[i, j].MergeCells)
                        {
                            throw exception;
                        }
                    }

                    j++;
                }
                else
                {
                    for (int i = startLine; i < startLine + countOfMergeCellsInNetworkScheme; i++)
                    {
                        if (sheet.Cells[i, j].MergeArea.Count != MergeCellsVolume)
                        {
                            throw exception;
                        }
                    }

                    j = j + MergeCellsVolume;
                }
            }
        }

        /*public void CheckTemperatureGroup(Worksheet sheetTRP, int startingIndex, List<BlankNode> nodeList, List<string> tempGroupList, Workbook bookTPNBU, int groupCount)
        {
            var column = sheetTRP.Cells.Find(TemperatureGroup).Column;

            if (sheetTRP.Cells[startingIndex, column].MergeCells)
            {
                var newNode = new BlankNode();
                newNode = nodeList[nodeList.Count - 1];
                Console.WriteLine("\nГруппы полетели:");
                foreach (var value in newNode.Values)
                {
                    Console.WriteLine(value);
                }

                var index = tempGroupList.IndexOf(TemperatureGroup);

                newNode.TemperatureGroup = tempGroupList[index];
                newNode.GetNetworkScheme(sheetTRP, startingIndex, bookTPNBU, tempGroupList, nodeList, newNode, groupCount);
                //newNode.GetControlledSection(sheetTRP, startingIndex, bookTPNBU.Worksheets["Сечения"], newNode);
                nodeList.Add(newNode);
            }
        }*/

        /// <summary>
        /// Метод для получения значений уставок
        /// </summary>
        public void GetValues(Worksheet sheetTRP, int startingIndex, int countOfMergeCellsInInfluencingFactor)
        {
            Values.Clear();
            Console.WriteLine("tempGroup - " + TemperatureGroup);
            var column = sheetTRP.Cells.Find(TemperatureGroup).Column;
            Console.WriteLine("tempGroup - " + column);

            /*if (sheetTRP.Cells[startingIndex, column].MergeCells)
            {
                Console.WriteLine("\nБыло - " + groupCount);
                Console.WriteLine("groupCount++");
                groupCount++;
                Console.WriteLine("Стало - " + groupCount);
                Console.WriteLine();
            }*/

            /*while (sheetTRP.Cells[startingIndex, column].Value is null)
            {
                Console.WriteLine($"Начальный столбец - {column}");
                column = column - 1;
                //TemperatureGroup = list[list.IndexOf(TemperatureGroup) - 1];
                //groupCount++;
                Console.WriteLine($"Конечный столбец - {column}");
                Console.ReadKey();
            }*/

            Console.WriteLine($"\nНачальная строка - {startingIndex}");
            Console.WriteLine($"Количество объединенных строк - {countOfMergeCellsInInfluencingFactor}");
            Console.WriteLine($"Объединить по следующую строку - {startingIndex + countOfMergeCellsInInfluencingFactor - 1}\n");

            var tmpValues = sheetTRP.Range[sheetTRP.Cells[startingIndex, column], sheetTRP.Cells[startingIndex + countOfMergeCellsInInfluencingFactor - 1, column]].Value;

            foreach (var value in tmpValues)
            {
                Console.WriteLine($"{startingIndex}, {column}");
                Console.WriteLine("value - " + value);

                var intValue = 0;

                if (!int.TryParse(value.ToString(), out intValue))
                {
                    throw new Exception("Значение уставки должно быть представлено числом.");
                }

                Values.Add(intValue);
            }

            //return groupCount;

            /*if (!sheetTRP.Cells[startingIndex, column].MergeCells)
            {
                var tmpValues = sheetTRP.Range[sheetTRP.Cells[startingIndex, column], sheetTRP.Cells[startingIndex + countOfMergeCellsInNetworkScheme - 1, column]].Value;

                foreach (var value in tmpValues)
                {
                    Console.WriteLine($"{startingIndex}, {column}");
                    Console.WriteLine("value - " + value);
                    Values.Add(value.ToString());
                }
            }
            else
            {
                var tmpValues = new List<string>();

                for (int i = 0; i < countOfMergeCellsInNetworkScheme; i++)
                {
                    tmpValues.Add(" ");
                }
                
                foreach (var value in tmpValues)
                {                    
                    Values.Add(value);
                }
            }*/

            //Console.WriteLine($"Начало - Cells[{startingIndex}, {column}]");
            //Console.WriteLine($"Конец - Cells[{startingIndex + countOfMergeCellsInNetworkScheme - 1}, {column}]");
            /*var tmpValues = sheetTRP.Range[sheetTRP.Cells[startingIndex, column], sheetTRP.Cells[startingIndex + countOfMergeCellsInNetworkScheme - 1, column]].Value;

            foreach (var value in tmpValues)
            {
                //Console.WriteLine(value);
                Values.Add(value.ToString());
            }*/
        }

        /// <summary>
        /// Метод для получения команд УВ
        /// </summary>
        public void GetControlActions(int startRowIndex, int schemeRowsCount, Worksheet worksheetTRP, Worksheet worksheetTPNBU)
        {
            var controlActionsListTRP = new List<string>();

            for (int j = 10; j <= 11; j++)
            {
                for (int i = startRowIndex; i <= startRowIndex + schemeRowsCount - 1; i++)
                {
                    controlActionsListTRP.Add(worksheetTRP.Cells[i, j].Value);
                }
            }

            /*foreach (var node in controlActionsListTRP)
            {
                Console.WriteLine(node);
            }*/

            //Console.WriteLine();     

            var controlActionsListTPNBU = new List<string>();

            foreach (var node in controlActionsListTRP)
            {
                if (string.IsNullOrEmpty(node))
                {
                    break;
                }

                string newNode = node;

                if (newNode.Contains(","))
                {
                    string newCommand = "";

                    while (true)
                    {
                        int index = 0;
                        if (newNode.Contains(","))
                        {
                            index = newNode.IndexOf(",");
                        }
                        else
                        {
                            index = newNode.Length;
                        }
                        
                        string command = newNode.Substring(0, index);
                        newCommand = newCommand + FindCommand(command) + ", ";

                        if (!newNode.Contains(","))
                        {
                            newCommand = newCommand.Remove(newCommand.Length - 2, 2);
                            controlActionsListTPNBU.Add(newCommand);
                            break;
                        }
                        else
                        {
                            newNode = newNode.Remove(0, index + 2);
                        }
                    }
                }
                else if (node.Contains("–"))
                {
                    controlActionsListTPNBU.Add("");
                }
                else
                {
                    controlActionsListTPNBU.Add(FindCommand(node));
                }
            }
            
            string FindCommand(string command)
            {
                Console.WriteLine(command + " " + command.Length);
                var line = worksheetTPNBU.Cells.Find(command).Row; ;

                return worksheetTPNBU.Cells[line, 1].Value;
            }

            int k = 0;
            ControlActions = new string[controlActionsListTPNBU.Count / 2, 2];
            for (int j = 0; j <= 1; j++)
            {
                for (int i = 0; i < controlActionsListTPNBU.Count / 2; i++)
                {
                    ControlActions[i, j] = controlActionsListTPNBU[k];
                    k++;
                }
            }

            /*for (int i = 0; i < controlActionsListTPNBU.Count / 2; i++)
            {
                for (int j = 0; j <= 1; j++)
                {
                    Console.Write(ControlActions[i, j] + "\t");
                }
                Console.WriteLine();
            }*/
        }

        public void CopyControlActions(string[,] ControlActionsArray)
        {
            ControlActions = ControlActionsArray;
        }

        private string GetOperationConditions(Workbook bookTPNBU)
        {
            var index = LaunchingOrganFullName.IndexOf("ВЛ");
            var conditions = LaunchingOrganFullName.Remove(0, index);
            conditions = conditions.Remove(conditions.IndexOf("(") - 1);

            return $"Вкл({FindShortNameOfLine(conditions, bookTPNBU)})";
        }

        private string GetOperatingStatus(Workbook bookTPNBU)
        {
            var sheet = bookTPNBU.Worksheets["ПО и ПС"];
            var rowIndex = sheet.Cells.Find(LaunchingOrganOperationName).Row;
            var status = "";

            if (sheet.Cells[rowIndex, 3].Value.ToString().ToLower() == "да")
            {
                status = $" и Вкл({LaunchingOrganOperationName})";
            }

            return status;
        }
        
        public int GetMergeLineCount(Worksheet sheetTRP, int startIndex, int column)
        {
            return sheetTRP.Cells[startIndex, column].MergeArea.Count;
        }

        public int GetTemperatureGroupCount(Worksheet sheetTRP, int startIndex)
        {
            var countCells = 0;

            for (int j = 6; j <= 9;)
            {
                countCells++;

                if (sheetTRP.Cells[startIndex, j].MergeCells)
                {
                    j = j + sheetTRP.Cells[startIndex, j].MergeArea.Count;
                }
                else
                {
                    j++;
                }
            }

            return countCells;
        }

        /// <summary>
        /// Метод для получения списка групп уставок по температуре наружного воздуха, которые содержаться в ТРП
        /// </summary>
        /// <param name="sheetTRP">ТРП</param>
        /// <returns></returns>
        public List<string> GetListOfTemperatureGroups(Worksheet sheetTRP, int startLine)
        {
            var groupArray = sheetTRP.Range[$"F{startLine - 2}", $"I{startLine - 2}"].Value;
            var groupList = new List<string>();

            foreach (var group in groupArray)
            {
                groupList.Add(group.ToString());
            }

            return groupList;
        }

        /// <summary>
        /// Метод для получения названия считываемой в текущий момент группы уставок
        /// </summary>
        /// <param name="sheetTRP"></param>
        /// <param name="bookTPNBU"></param>
        /// <param name="currentLine"></param>
        /// <param name="list"></param>
        /// <param name="headingLine"></param>
        /// <returns></returns>
        private string GetTemperatureGroupName(List<string> groupList, List<BlankNode> nodeList, Workbook bookTPNBU, int tempGroup)
        {   
            if (tempGroup == 1)
            {
                TemperatureGroup = groupList[0];

                return "";
            }

            switch (nodeList.Count)
            {
                case 0:
                    TemperatureGroup = groupList[0];
                    break;
                default:
                    var index = groupList.IndexOf(nodeList[nodeList.Count - 1].TemperatureGroup);

                    if (index == groupList.Count - 1)
                    {
                        TemperatureGroup = groupList[0];
                    }
                    else
                    {
                        TemperatureGroup = groupList[index + 1];
                    }
                    Console.WriteLine("Группа - " + TemperatureGroup);
                    break;
            }

            var sheet = bookTPNBU.Worksheets["Группы"];
            var line = sheet.Cells.Find(TemperatureGroup).Row;
            //Console.WriteLine("line - " + line);
            return $" и Вкл({sheet.Cells[line, 1].Value.ToString()})";
        }

        /// <summary>
        /// Метод для получения сокращенных наименований линий
        /// </summary>
        /// <param name="fullNameOfLine">Полное наименование линии</param>
        /// <param name="bookTPNBU">Документ Excel для поиска сокращенного наименования линии</param>
        /// <returns></returns>
        private string FindShortNameOfLine(string fullNameOfLine, Workbook bookTPNBU)
        {
            string shortName;
            
            try
            {
                var sheet = bookTPNBU.Worksheets["Дисциплины"];
                var index = sheet.Cells.Find(fullNameOfLine.Replace("–", "-")).Row;

                shortName = sheet.Cells[index, 1].Value.ToString();
            }
            catch
            {
                throw new Exception($"В ТПНБУ нет сокращенного наименования для искомой линии: " +
                    $"{fullNameOfLine}");
            }

            return shortName;
        }

        /// <summary>
        /// Метод для проверки корректности задания схемы сети
        /// </summary>
        /// <param name="scheme"></param>
        private void CheckNetworkScheme(string scheme)
        {
            scheme = scheme.ToLower();
            var schemeTmp = @"^ремонт|работа [а-яёa-z0-9\–\№\ ]{1,100}$";/* [а-яёa-z0-9–№]{1,100}$*/
            var correctSchemeRegex = new Regex(schemeTmp);
            
            while (!string.IsNullOrEmpty(scheme))
            {
                string partOfScheme;

                if (scheme.Contains(" и "))
                {
                    partOfScheme = scheme.Substring(0, scheme.IndexOf(" и "));
                }
                else
                {
                    partOfScheme = scheme;
                }

                if (!correctSchemeRegex.IsMatch(partOfScheme))
                {
                    throw new Exception($"Ошибка при задании схемы сети: {partOfScheme}");
                }

                if (scheme.Contains(" и "))
                {
                    scheme = scheme.Remove(0, scheme.IndexOf(" и ") + 3);
                }
                else
                {
                    scheme = "";
                }
            }

        }

        public void GetNetworkScheme(Workbook bookTRP, int startingIndex, Workbook bookTPNBU, List<string> list, 
            List<BlankNode> nodeList, int groupCount)
        {
            var scheme = bookTRP.Worksheets["Настройка ПО"].Cells[startingIndex, 3].Value.ToString().ToLower();
            Console.WriteLine("scheme: " + scheme);

            switch(scheme)
            {
                case "любая":
                    SchemeOfNetwork = GetOperationConditions(bookTPNBU) + GetTemperatureGroupName(list, nodeList, bookTPNBU, groupCount)
                        + GetOperatingStatus(bookTPNBU);
                    break;
                case "нормальная":
                    var row = bookTRP.Worksheets["Нормальная схема"].Cells.Find("нормальная").Row;
                    scheme = bookTRP.Worksheets["Нормальная схема"].Cells[row, 2].Value;

                    CheckNetworkScheme(scheme);
                    var secondaryScheme2 = "";

                    while (!string.IsNullOrEmpty(scheme))
                    {
                        string partOfScheme;

                        if (scheme.Contains(" и "))
                        {
                            partOfScheme = scheme.Substring(0, scheme.IndexOf(" и "));
                        }
                        else
                        {
                            partOfScheme = scheme;
                        }

                        if (scheme.Contains(" и "))
                        {
                            scheme = scheme.Remove(0, scheme.IndexOf(" и ") + 3);
                        }
                        else
                        {
                            scheme = "";
                        }

                        secondaryScheme2 = secondaryScheme2 + GetNameOfWorkingOrRepairScheme(bookTPNBU, partOfScheme);
                    }

                    SchemeOfNetwork = secondaryScheme2 + GetOperationConditions(bookTPNBU) +
                        GetTemperatureGroupName(list, nodeList, bookTPNBU, groupCount) + GetOperatingStatus(bookTPNBU);
                    break;
                default:
                    CheckNetworkScheme(scheme);
                    var secondaryScheme = "";

                    while (!string.IsNullOrEmpty(scheme))
                    {
                        string partOfScheme;

                        if (scheme.Contains(" и "))
                        {
                            partOfScheme = scheme.Substring(0, scheme.IndexOf(" и "));
                        }
                        else
                        {
                            partOfScheme = scheme;
                        }

                        if (scheme.Contains(" и "))
                        {
                            scheme = scheme.Remove(0, scheme.IndexOf(" и ") + 3);
                        }
                        else
                        {
                            scheme = "";
                        }

                        secondaryScheme = secondaryScheme + GetNameOfWorkingOrRepairScheme(bookTPNBU, partOfScheme);
                    }

                    SchemeOfNetwork = secondaryScheme + GetOperationConditions(bookTPNBU) +
                        GetTemperatureGroupName(list, nodeList, bookTPNBU, groupCount) + GetOperatingStatus(bookTPNBU);
                    break;
            }
        }

        private string GetNameOfWorkingOrRepairScheme(Workbook bookNPNBU, string scheme)
        {
            string newScheme = "";
            Console.WriteLine("Схема сети на входе: " + scheme);
            while (scheme.Length > 0)
            {
                switch (scheme.ToLower().Substring(0, 6))
                {
                    case "ремонт":
                        newScheme = newScheme + "Откл";
                        scheme = scheme.Remove(0, 7);
                        Console.WriteLine("ремонт");
                        break;
                    case "работа":
                        newScheme = newScheme + "Вкл";
                        scheme = scheme.Remove(0, 7);
                        Console.WriteLine("работа");
                        break;
                    default:
                        throw new Exception("В описании схемы сети отсутствуют ключевые слова 'ремонт/работа'.");
                }
                Console.WriteLine("Схема сети начало - " + scheme);
                var conditionIndex = 0;
                
                if (scheme.IndexOf(" и ") >= 0)
                {
                    Console.WriteLine("Yes");
                    conditionIndex = scheme.IndexOf(" и ");
                    Console.WriteLine(conditionIndex);
                    Console.WriteLine("Поиск - " + scheme);
                    newScheme = newScheme + $"({FindShortNameOfLine(scheme.Substring(0, conditionIndex), bookNPNBU)}) и ";
                    Console.WriteLine("newScheme: " + newScheme);
                    scheme = scheme.Remove(0, conditionIndex + 3);
                    Console.WriteLine("scheme: " + scheme);
                }
                else
                {
                    Console.WriteLine("No");
                    Console.WriteLine("Поиск - " + scheme);
                    newScheme = newScheme + $"({FindShortNameOfLine(scheme, bookNPNBU)})";
                    scheme = "";
                }
            }
            Console.WriteLine("Схема сети конец - " + newScheme);
            return $"{newScheme} и ";
        }

        /// <summary>
        /// Метод для получения колчества пусковых органов, имеющихся в ТРП
        /// </summary>
        /// <param name="sheetTRP"></param>
        /// <param name="startLine"></param>
        /// <returns></returns>
        public int GetCountOfLaunchingOrgan(Worksheet sheetTRP, int startLine)
        {
            var rowCount = Find(sheetTRP) - startLine/*sheetTRP.Cells[startLine, 2].MergeArea.Count*/;
            var counter = 0;

            for (int i = startLine; i < startLine + rowCount;)
            {
                counter++;
                var mergeCells = sheetTRP.Cells[i, 2].MergeArea.Count;

                i = i + mergeCells;
            }

            return counter;
        }

        /// <summary>
        /// Метод для получения количества записей нижнего уровня, относящихся к записи верхнего уровня
        /// </summary>
        /// <param name="sheetTRP"></param>
        /// <param name="startLine"></param>
        /// <param name="column"></param>
        /// <param name="step"></param>
        /// <returns></returns>
        public int GetCountOfParams(Worksheet sheetTRP, int startLine, int column, int step)
        {
            var rowCount = sheetTRP.Cells[startLine, column].MergeArea.Count;
            var counter = 0;

            for (int i = startLine; i < startLine + rowCount;)
            {
                Console.WriteLine($"\n\ni = {i}\n\n");
                counter++;
                var countOfMergeCells = sheetTRP.Cells[i, column + step].MergeArea.Count;

                i = i + countOfMergeCells;
            }

            return counter;
        }

        /// <summary>
        /// Метод для получения индекса первой стройки с настройкой ПО
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public int FindStartingLine(Worksheet sheet)
        {
            var index = 0;

            var RowsCount = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                    false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            for (int i = 2; i <= RowsCount; i++)
            {
                if (sheet.Cells[i, 2].Value.ToString().Contains("Автоматика"))
                {
                    index = i;
                    break;
                }

                if (sheet.Cells[i, 2].MergeCells)
                {
                    i = i + sheet.Cells[i, 2].MergeArea.Count - 1;
                }
            }

            if (index == 0)
            {
                throw new Exception("Таблица не содержит записей с уставками для пусковых органов.");
            }

            return index;
        }

        public int Find(Worksheet sheet)
        {
            var RowsCount = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                    false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            return RowsCount + 1;
        }

        /// <summary>
        /// Метод для получения влияющих факторов
        /// </summary>
        /// <returns></returns>
        public void GetInfluencingFactor(Worksheet sheetTRP, int line)
        {
            var inputFactor = sheetTRP.Cells[line, 5].Value;

            var factorTmpList = new List<string>()
            {
                @"^ГГ.[а-яёА-ЯЁ]{1,13} [=,>,<,≥,≤]{1} [0-9]{1,2}$",
                @"^P.[а-яёА-ЯЁ]{1,13} [=,>,<,≥,≤]{1} [0-9]{3,5}$"
            };

            string partOfFactor;
            var fullFactor = "";

            while (!string.IsNullOrEmpty(inputFactor))
            {
                if (inputFactor.Contains(" и "))
                {
                    partOfFactor = inputFactor.Substring(0, inputFactor.IndexOf(" и "));
                }
                else
                {
                    partOfFactor = inputFactor;
                }

                switch (partOfFactor)
                {
                    case "–":
                        partOfFactor = "";
                        break;
                    default:
                        var fallsCounter = 0;

                        foreach (string template in factorTmpList)
                        {
                            var factorRegex = new Regex(template);

                            if (!factorRegex.IsMatch(partOfFactor))
                            {
                                fallsCounter++;
                            }
                        }

                        if (fallsCounter == factorTmpList.Count)
                        {
                            throw new Exception($"Ошибка при указании названия влияющего фактора: {partOfFactor}");
                        }

                        break;
                }

                if (inputFactor.Contains(" и "))
                {
                    fullFactor = fullFactor + partOfFactor + " и ";
                    inputFactor = inputFactor.Remove(0, inputFactor.IndexOf(" и ") + 3);
                }
                else
                {
                    fullFactor = fullFactor + partOfFactor;
                    inputFactor = "";
                }
            }

            InfluencingFactor = fullFactor;
        }
    }
}
