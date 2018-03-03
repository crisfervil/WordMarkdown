using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordMarkdown
{
    public class Generator
    {
        public string GetJSon(string fileName)
        {
            string jsonData = null;

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, false))
            {
                var currentLevel = -1;
                string currentContainerKey = null;
                var dataStack = new Stack<object>();
                object stackBottom = null;

                var rootElement = doc.MainDocumentPart.Document;

                foreach (var documentItem in rootElement.Descendants().Where(d => (d is Paragraph || d is Table) && d.Parent == rootElement.Body))
                {
                    if (documentItem is Paragraph)
                    {
                        var currentParagraph = (Paragraph)documentItem;
                        var paragrapLevel = GetOutline(currentParagraph, doc.MainDocumentPart.StyleDefinitionsPart);
                        if (paragrapLevel > -1)
                        {
                            var asdasd = GetParagraphText(currentParagraph);
                            if (paragrapLevel > currentLevel)
                            {
                                var previousContainerKey = currentContainerKey;
                                currentContainerKey = GetParagraphText(currentParagraph);
                                var currentObjectDict = new Dictionary<string, object>();
                                currentObjectDict.Add(currentContainerKey, null);

                                if (dataStack.Count == 0) // initialize the stack
                                {
                                    stackBottom = currentObjectDict;
                                    dataStack.Push(currentObjectDict);
                                }
                                else
                                {
                                    if (dataStack.Peek() is Dictionary<string, object> topStackDic
                                        && topStackDic.ContainsKey(previousContainerKey)
                                        && topStackDic[previousContainerKey] != null)
                                    {
                                        dataStack.Push(topStackDic[previousContainerKey]);
                                    }

                                    Stack(dataStack, previousContainerKey, currentObjectDict);
                                    dataStack.Push(currentObjectDict);
                                }

                                // increase level
                                currentLevel = paragrapLevel;
                            }
                            else if (paragrapLevel == currentLevel)
                            {
                                currentContainerKey = GetParagraphText(currentParagraph);
                                Stack(dataStack, currentContainerKey, null);
                            }
                            else
                            {
                                // remove elements from the stack
                                for (; currentLevel > paragrapLevel; currentLevel--) dataStack.Pop();

                                currentContainerKey = GetParagraphText(currentParagraph);
                                Stack(dataStack, currentContainerKey, null);
                            }
                        }
                    }
                    else if (documentItem is Table tableItem)
                    {
                        if (dataStack.Count == 0) // init the stack if is empty
                        {
                            stackBottom = new List<object>();
                            dataStack.Push(stackBottom);
                        }

                        var dataItem = ProcessTable(tableItem);

                        Stack(dataStack, currentContainerKey, dataItem);
                    }
                }

                jsonData = JsonConvert.SerializeObject(stackBottom, Formatting.Indented);
                return jsonData;
            }
        }

        private int GetOutline(Paragraph paragraph, StyleDefinitionsPart styles)
        {
            // get paragraph outline
            var outlineLevel = -1;
            if (paragraph.ParagraphProperties?.OutlineLevel != null)
            {
                outlineLevel = paragraph.ParagraphProperties.OutlineLevel.Val.Value;
            }
            else if (paragraph.ParagraphProperties?.ParagraphStyleId != null)
            {
                var style = styles.RootElement.Elements<Style>().Where(s => s.StyleId.Value == paragraph.ParagraphProperties.ParagraphStyleId.Val).FirstOrDefault();
                outlineLevel = style?.StyleParagraphProperties?.OutlineLevel?.Val;
            }
            return outlineLevel;
        }

        private static string GetParagraphText(Paragraph paragraph)
        {
            var sb = new StringBuilder();
            foreach (var text in paragraph.Descendants<Text>())
            {
                sb.Append(text.InnerText);
            }
            return sb.ToString();
        }

        private static void Stack(Stack<object> stack, string key, object value)
        {
            if (stack.Peek() is Dictionary<string, object> currentContainerDictionary)
            {
                currentContainerDictionary[key] = value;
            }

            if (stack.Peek() is List<object> currentContainerList)
            {
                currentContainerList.Add(value);
            }
        }

        private static List<object> ProcessTables(IEnumerable<Table> tables)
        {
            var data = new List<object>();
            foreach (var table in tables)
            {
                var dataItem = ProcessTable(table);
                data.Add(dataItem);
            }
            return data;
        }

        private static object ProcessTable(Table table)
        {
            object data = null;

            var tableRows = table.Descendants<TableRow>().ToList();
            var firstRow = tableRows[0];
            var firstRowColumns = firstRow.Descendants<TableCell>().ToList();

            // There are three table types depending on the number of columns
            if (firstRowColumns.Count == 1)
            {
                // If the table contains a single column, the data in it will be considered a single items array

            }
            else if (firstRowColumns.Count == 2)
            {
                // If contains two rows, it will be considered an object, where the first column is the property name and the second is the property value
                data = GetTwoColumnsTableData(table);

            }
            else if (firstRowColumns.Count > 2)
            {
                // This table type is an array of objects, where every property is a column
                data = GetMultipleColumnsTableData(table);
            }

            return data;
        }

        private static object GetTwoColumnsTableData(Table table)
        {
            var data = new Dictionary<string, object>();

            foreach (var row in table.Descendants<TableRow>())
            {
                // ignore rows that are not from this table
                if (row.Parent != table) continue;

                var rowColumns = row.Descendants<TableCell>().ToList();
                var firstCol = rowColumns[0];
                var secondCol = rowColumns[1];
                var propName = GetCellText(firstCol);
                var propValue = GetCellValue(secondCol);

                data.Add(propName, propValue);
            }

            return data;
        }

        private static object GetCellValue(TableCell tableCell)
        {
            object cellValue = null;
            var nestedTables = tableCell.Descendants<Table>();

            // if this table contains other nested tables, then an object will be returned
            if (nestedTables.Any())
            {
                cellValue = ProcessTables(nestedTables.Where(x => x.Parent == tableCell));
            }
            else
            {
                // otherwise, just the text inside
                cellValue = GetCellText(tableCell);
            }

            return cellValue;
        }

        private static string GetCellText(TableCell tableCell)
        {
            var sb = new StringBuilder();
            foreach (var paragraph in tableCell.Descendants<Paragraph>())
            {
                var paragraphText = GetParagraphText(paragraph);
                sb.Append(paragraphText);
            }
            return sb.ToString();
        }

        private static object GetMultipleColumnsTableData(Table table)
        {
            var data = new List<object>();
            var tableRows = table.Descendants<TableRow>().ToList();

            // the first row contains the name of the attributes
            var attributeNames = new List<string>();
            var firstRow = tableRows[0];
            var firstRowCells = firstRow.Descendants<TableCell>().Where(c => c.Parent == firstRow);

            foreach (var firstRowCell in firstRowCells)
            {
                var attributeName = GetCellText(firstRowCell);
                attributeNames.Add(attributeName);
            }

            for (var i = 1; i < tableRows.Count; i++)
            {
                var tableRow = tableRows[i];
                var rowObject = new Dictionary<string, object>();

                // get the attribute values
                var rowCells = tableRow.Descendants<TableCell>().Where(c => c.Parent == tableRow);
                var attrIndex = 0;
                foreach (var rowCell in rowCells)
                {
                    var attributeName = attributeNames[attrIndex];

                    if (!string.IsNullOrWhiteSpace(attributeName))
                    {
                        var attributeValue = GetCellValue(rowCell);
                        rowObject.Add(attributeName, attributeValue);
                    }
                    attrIndex++;
                }

                data.Add(rowObject);
            }

            return data;
        }

    }
}
