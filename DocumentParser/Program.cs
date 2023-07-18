using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentParser;

const string RED_COLOR = "ed6868";

Console.WriteLine("Начинаем обработку...");
var inputFiles = GetInputFiles();
var outputFile = GetOutputFile();
Console.WriteLine($"Найдены входные файлы: {string.Join(", ", inputFiles.Select(Path.GetFileName))}");

foreach (var inputFile in inputFiles)
{
    var inputData = new InputTable();
    Console.WriteLine($"Чтение данных из файла: {inputFile}...");
    ReadFromFile(inputData, inputFile);
    Console.WriteLine($"\nЧтение данных из файла: {inputFile} завершено");

    Console.WriteLine($"Запись данных из файла: {inputFile} в {outputFile}");
    WriteToFile(inputData, outputFile);
    Console.WriteLine($"\nЗапись данных из файла: {inputFile} в {outputFile} завершена");
}
Console.WriteLine("Обработка завершена.");
Console.ReadKey();

string[] GetInputFiles()
{
    return Directory.GetFiles("Content\\Input"); // получаем список файлов в директории
}

string GetOutputFile()
{
    return Directory.GetFiles("Content\\Output").FirstOrDefault(); // получаем список файлов в директории
}

List<double>? ChangeRange(TableCell cell)
{
    var text = string.Join("", cell.Select(x => x.InnerText));
    return text.Any(char.IsDigit)
        ? (text.Contains('±')
            ? new[]
            {
                $"-{Regex.Replace(text, "[^0-9,.]", "")}",
                Regex.Replace(text, "[^0-9,.]", "")
            }
            : Regex.Split(text, "[.]{2,}|…|÷")).Select(s => double.Parse(Regex.Replace(s, "[.]", ","))).ToList()
        : null;
}

List<DiscreteTable> GetDiscreteTables(WordprocessingDocument? document)
{
    var discreteTables = new List<DiscreteTable>();

    if (document == null) return discreteTables;
    
    foreach (var table in document.MainDocumentPart.Document.Body.Elements<Table>().Skip(1))
    {
        var discreteTable = new DiscreteTable();
        foreach (var row in table.Elements<TableRow>().Skip(2))
        {
            var cells = row.Elements<TableCell>().ToArray();
           
            if (cells[1].InnerText.Contains("Адрес слова"))
            {
                discreteTable.Address = Regex.Match(cells[1].InnerText, @"\d+").Value;
            }

            if (cells[2].InnerText.Contains('+'))
            {
                discreteTable.Parameters.Add(new DiscreteParameter
                {
                    Id = int.Parse(cells[0].InnerText),
                    Status = cells[1].TableCellProperties?.VerticalMerge == null ? cells[3].InnerText : null
                });
            }
        }
        discreteTables.Add(discreteTable);
    }

    return discreteTables;
}

int LogProcessCount(int processCount, int count, string? text)
{
    Console.SetCursorPosition(0, Console.CursorTop);
    processCount++;
    Console.Write($"{text}: {processCount} из {count}");
    return processCount;
}

void ReadFromFile(InputTable inputData1, string filePath)
{
    using var doc = WordprocessingDocument.Open(filePath, false);

    inputData1.Name = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().FirstOrDefault().InnerText;
    
    // Получаем первую таблицу из документа Word
    var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

    List<DiscreteTable> discreteTables = GetDiscreteTables(doc);

    int count = table.Elements<TableRow>().Skip(2).Count();
    int processCount = 0;
    // Обрабатываем таблицу
    foreach (var row in table.Elements<TableRow>().Skip(2))
    {
        var cells = row.Elements<TableCell>().Skip(1).ToArray();
        var parameter = new Parameter
        {
            Signal = cells[0].InnerText,
            Designation = cells[1].InnerText,
            TypeSignal = cells[2].InnerText,
            Unit = cells[3].InnerText,
            ChangeRange = ChangeRange(cells[4]),
            Address = cells[5].InnerText,
            HighDischargesPrice = cells[6].InnerText,
            QuantityMeaningDischarges = /*int.Parse(cells[7].InnerText)*/ cells[7].InnerText,
            FrequencyRegister = /*int.Parse(cells[7].InnerText)*/ cells[8].InnerText,
        };

        if (parameter.TypeSignal == "DW")
        {
            var discreteTable = discreteTables.Where(s => s.Address == parameter.Address).FirstOrDefault();

            foreach (var param in discreteTable.Parameters)
            {
                var par = new Parameter
                {
                    Signal = parameter.Signal,
                    Designation = parameter.Designation + "_b" + param.Id,
                    TypeSignal = parameter.TypeSignal,
                    Unit = parameter.Unit,
                    ChangeRange = parameter.ChangeRange,
                    Address = parameter.Address,
                    HighDischargesPrice = parameter.HighDischargesPrice,
                    QuantityMeaningDischarges = parameter.QuantityMeaningDischarges,
                    FrequencyRegister = parameter.FrequencyRegister,
                    LSB = param.Id.ToString(),
                    MSB = param.Id.ToString(),
                    ValueInZero = "",
                    ValueInOne = param.Status
                };

                inputData1.Parameters.Add(par);
            }
        }
        else
        {
            if (parameter.FrequencyRegister != "-") 
                inputData1.Parameters.Add(parameter);
        }
        processCount = LogProcessCount(processCount, count, $"Прочитанно строк из файла {filePath}");
    }
}

void WriteToFile(InputTable inputTable, string filePath)
{
    
    using (var doc = WordprocessingDocument.Open(filePath, true))
    {
        // Получаем таблицу 4.3.1 Определение интерфейсов
        
        Tuple<string, Table>[] documentTables = doc.MainDocumentPart.Document.Body
            .Elements<Table>()
            .Select(x=> new Tuple<string, Table>(x.PreviousSibling<Paragraph>().InnerText, x)).ToArray();
        
        
        var outputTableName = documentTables.FirstOrDefault(x=>x.Item1.Contains("Определение интерфейсов")).Item2
            .Descendants<TableRow>()
            .FirstOrDefault(t => inputTable.Name.Contains(t.Elements<TableCell>().ToArray()[6].InnerText)).ToArray()[1].InnerText;
            
        
        var table = documentTables.FirstOrDefault(x=>x.Item1.Contains(outputTableName)).Item2;

        TableRow[] rowsToDelete = table.Elements<TableRow>().Skip(1).ToArray();

        foreach (var row in rowsToDelete)
        {
            row.Remove();
        }

        var props = new TableProperties(
            new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1 
                }));
        
        table.AppendChild(props);

        int count = inputTable.Parameters.Count;
        int processCount = 1;
        for (var i = 1; i < inputTable.Parameters.Count; i++)
        {
            var newRow = new TableRow();
            if (inputTable.Parameters[i].TypeSignal == "BNR")
            {
                
                // ID параметра
                newRow.AppendChild(GetCell(outputTableName  + "-" + i.ToString("D3")));
               
                // Наименование параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Designation));
                
                // Расшифровка наименования параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Signal));
                
                //Label
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Address));
                
                // Тип параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].TypeSignal));
                
                // Единица измерения
                newRow.AppendChild(GetCell(GetOutputUnit(inputTable.Parameters[i].Unit)));

                //Тип матрицы состояния (SSM)
                newRow.AppendChild(GetCell(inputTable.Parameters[i].TypeSignal));
                
                //ИИП (SDI)
                newRow.AppendChild(GetCell("00"));
                
                //СЗР (MSB)
                newRow.AppendChild(inputTable.Parameters[i].QuantityMeaningDischarges switch
                {
                    "15" => GetCell("28"),
                    _ => GetCell("", RED_COLOR)
                });

                // МЗР (LSB)
                newRow.AppendChild(inputTable.Parameters[i].QuantityMeaningDischarges switch
                {
                    "15" => GetCell("14"),
                    _ => GetCell("", RED_COLOR)
                });

                // НСР (MSB Weight In Units
                newRow.AppendChild(GetCell(inputTable.Parameters[i].HighDischargesPrice));
                
                var usingSignBit = inputTable.Parameters[i].ChangeRange == null 
                ? "N/A"
                : inputTable.Parameters[i].ChangeRange[0] < 0 
                    ? "YES"
                    : "NO";
                // Использование знакового разряда
                newRow.AppendChild(GetCell(usingSignBit));
                
                // Физический диапазон
                var changeRange = inputTable.Parameters[i].ChangeRange;
                newRow.AppendChild(changeRange != null
                    ? GetCell(string.Join("…", changeRange))
                    : GetCell("N/A"));
                
                // Интервал обновления (refresh time), мс
                newRow.AppendChild(int.TryParse(inputTable.Parameters[i].FrequencyRegister, out var fr)
                    ? GetCell((1000 / fr).ToString())
                    : GetCell("TBD"));

                // Время задержки(Latency time), ms
                newRow.AppendChild(GetCell("TBD"));
                
                // Значение в «0»
                newRow.AppendChild(GetCell("N/A"));
                
                // Значение в «1»
                newRow.AppendChild(GetCell("N/A"));
                
                //Комментарии
                newRow.AppendChild(GetCell(""));

                table.AppendChild(newRow);
            }

            if (inputTable.Parameters[i].TypeSignal == "DW")
            {
                // ID параметра
                newRow.AppendChild(GetCell(outputTableName  + "-" + i.ToString("D3")));
                
                // Наименование параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Designation));
                
                // Расшифровка наименования параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Signal));
                
                //Label
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Address));
                
                // Тип параметра
                newRow.AppendChild(GetCell("DW"));
                
                // Единица измерения
                newRow.AppendChild(GetCell("N/A"));
                
                //Тип матрицы состояния (SSM)
                newRow.AppendChild(GetCell("DW"));
                
                //ИИП (SDI)
                newRow.AppendChild(GetCell("00"));
                
                //СЗР (MSB)
                newRow.AppendChild(GetCell(inputTable.Parameters[i].MSB));

                // МЗР (LSB)
                newRow.AppendChild(GetCell(inputTable.Parameters[i].LSB));

                // НСР (MSB Weight In Units
                newRow.AppendChild(GetCell(inputTable.Parameters[i].HighDischargesPrice));
                
                var usingSignBit = inputTable.Parameters[i].ChangeRange == null 
                    ? "N/A"
                    : inputTable.Parameters[i].ChangeRange[0] < 0 
                        ? "YES"
                        : "NO";
                // Использование знакового разряда
                newRow.AppendChild(GetCell(usingSignBit));
                
                // Физический диапазон
                var changeRange = inputTable.Parameters[i].ChangeRange;
                newRow.AppendChild(changeRange != null
                    ? GetCell(string.Join("…", changeRange))
                    : GetCell("N/A"));
                
                // Интервал обновления (refresh time), мс
                newRow.AppendChild(int.TryParse(inputTable.Parameters[i].FrequencyRegister, out var fr)
                    ? GetCell((1000 / fr).ToString())
                    : GetCell(""));

                // Время задержки(Latency time), ms
                newRow.AppendChild(GetCell("TBD"));
                
                // Значение в «0»
                newRow.AppendChild(GetCell("", RED_COLOR));
                
                // Значение в «1»
                newRow.AppendChild(GetCell(inputTable.Parameters[i].ValueInOne, String.IsNullOrEmpty(inputTable.Parameters[i].ValueInOne) ? RED_COLOR : null));
                
                //Комментарии
                newRow.AppendChild(GetCell(""));
                
                table.AppendChild(newRow);
            }
            
            processCount = LogProcessCount(processCount, count, $"Записанно строк в файл {outputFile}");
        }

        // Сохраняем изменения в документе Word
        doc.MainDocumentPart.Document.Save();
    }
}

string GetOutputUnit(string? unit)
{
    return unit switch
    {
        "град" => "Deg",
        "" => "",
        "мм" => "mm",
        "град/c" => "Deg/s",
        "оC" => "оC",
        "ед." => "1",
        "N/A" => "N/A",
        _ => ""
    };
}

TableCell GetCell(string value, string? backgroundColor = null)
{
    var tableCell = new TableCell();

    var tableCellProperties = new TableCellProperties();
    tableCellProperties.Append(new TableCellVerticalAlignment {Val = TableVerticalAlignmentValues.Center});
    tableCellProperties.Append(new NoWrap {Val = OnOffOnlyValues.On});
    if (backgroundColor != null)
        tableCellProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Shading()
        {
            Color = "auto",
            Fill = backgroundColor,
            Val = ShadingPatternValues.Clear
        });
    
    var paragraph = new Paragraph();
    
    var run = new Run();
    run.AppendChild(new Text(value));
    
    var runProp = new RunProperties();
 
    var runFont = new RunFonts
    {
        Ascii = "Arial",
        HighAnsi = "Arial"
    };
    
    var size = new FontSize { Val = new StringValue("14") };
    
    var justification = new Justification()
    {
        Val = JustificationValues.Right
    };
    
    var paragraphProperties = new ParagraphProperties();
    paragraphProperties.Justification = new Justification()
    {
        Val = JustificationValues.Center
    };

    //runProp.InsertBefore(justification, runProp.Elements().FirstOrDefault());
    runProp.Append(runFont);
    runProp.Append(size);

    run.PrependChild(runProp);
    
    paragraph.Append(paragraphProperties);
    paragraph.Append(run);

    tableCell.Append(tableCellProperties);
    tableCell.Append(paragraph);

    return tableCell;
}