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
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("")))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].Designation)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].Signal)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].Address)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].TypeSignal)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].Unit)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].TypeSignal)))));
           
            newRow.AppendChild(inputTable.Parameters[i].TypeSignal == "BNR"
                ? new TableCell(new Paragraph(new Run(new Text("00"))))
                : new TableCell(new Paragraph(new Run(new Text("")))));
            
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("28")))));

            

            newRow.AppendChild(inputTable.Parameters[i].TypeSignal == "BNR" &&
                               inputTable.Parameters[i].QuantityMeaningDischarges == "15"
                ? new TableCell(new Paragraph(new Run(new Text("14"))))
                : new TableCell(new Paragraph(new Run(new Text("13")))));
            
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(inputTable.Parameters[i].HighDischargesPrice)))));
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("")))));


            var changeRange = inputTable.Parameters[i].ChangeRange;
            if (changeRange != null)
                newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(string.Join("...", changeRange))))));

            newRow.AppendChild(int.TryParse(inputTable.Parameters[i].FrequencyRegister, out var fr)
                ? new TableCell(new Paragraph(new Run(new Text((1000/fr).ToString()))))
                : new TableCell(new Paragraph(new Run(new Text("")))));
             
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("TBD")))));
            
            newRow.AppendChild(inputTable.Parameters[i].TypeSignal == "BNR"
                ? new TableCell(new Paragraph(new Run(new Text("N/A"))))
                : new TableCell(new Paragraph(new Run(new Text("")))));
            
            newRow.AppendChild(inputTable.Parameters[i].TypeSignal == "BNR"
                ? new TableCell(new Paragraph(new Run(new Text("N/A"))))
                : new TableCell(new Paragraph(new Run(new Text("")))));
            
            newRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("")))));

            table.AppendChild(newRow);
        }
        // Сохраняем изменения в документе Word
        doc.MainDocumentPart.Document.Save();
    }
}