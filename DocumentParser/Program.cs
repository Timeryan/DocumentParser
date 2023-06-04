using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentParser;

var inputData = new InputTable();

ReadFromFile(inputData);
WriteToFile(inputData);
Console.WriteLine("Готово");


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

void ReadFromFile(InputTable inputData1)
{
    using var doc = WordprocessingDocument.Open("Content/Input.docx", false);

    // Получаем первую таблицу из документа Word
    var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

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

        inputData1.Parameters.Add(parameter);
    }
}

void WriteToFile(InputTable inputTable)
{
    using (var doc = WordprocessingDocument.Open("Content/Template.docx", true))
    {
        // Получаем первую таблицу из документа Word
        var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

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


        for (var i = 1; i < inputTable.Parameters.Count; i++)
        {
            var newRow = new TableRow();
            if (inputTable.Parameters[i].TypeSignal == "BNR")
            {
                
                // ID параметра
               newRow.AppendChild(GetCell(""));
               
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
                newRow.AppendChild(GetCell("28"));

                // МЗР (LSB)
                newRow.AppendChild(inputTable.Parameters[i].QuantityMeaningDischarges switch
                {
                    "15" => GetCell("14"),
                    "16" => GetCell("13"),
                    _ => GetCell("")
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
                    : GetCell(""));

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
                newRow.AppendChild(GetCell(""));
                
                // Наименование параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Designation));
                
                // Расшифровка наименования параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Signal));
                
                //Label
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Address));
                
                // Тип параметра
                newRow.AppendChild(GetCell(inputTable.Parameters[i].TypeSignal));
                
                // Единица измерения
                newRow.AppendChild(GetCell(inputTable.Parameters[i].Unit));
                
                //Тип матрицы состояния (SSM)
                newRow.AppendChild(GetCell(inputTable.Parameters[i].TypeSignal));
                
                //ИИП (SDI)
                newRow.AppendChild(GetCell("00"));
                
                //СЗР (MSB)
                newRow.AppendChild(GetCell("28"));

                // МЗР (LSB)
                newRow.AppendChild(inputTable.Parameters[i].QuantityMeaningDischarges switch
                {
                    "15" => GetCell("14"),
                    "16" => GetCell("13"),
                    _ => GetCell("")
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
                    : GetCell(""));

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

TableCell GetCell(string value)
{
    var run = new Run();
    run.AppendChild(new Text(value));
    
    var paragraph = new Paragraph(run);
    var runProp = new RunProperties();
 
    var runFont = new RunFonts
    {
        Ascii = "Arial",
        HighAnsi = "Arial"
    };

    var size = new FontSize { Val = new StringValue("14") };
 
    runProp.Append(runFont);
    runProp.Append(size);
 
    run.PrependChild(runProp);

    return new TableCell(paragraph);
}