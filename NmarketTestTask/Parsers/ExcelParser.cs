using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NmarketTestTask.Models;

namespace NmarketTestTask.Parsers
{
    public class ExcelParser : IParser
    {
        public IList<House> GetHouses(string path)
        {
            var houses = new List<House>();

            var workbook = new XLWorkbook(path);
            var sheets = workbook.Worksheets;

            foreach (var sheet in sheets)
            {
                House house = new House();

                var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
                var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                for (int row = 1; row <= lastRow; row++)
                {
                    for (int col = 1; col <= lastColumn; col++)
                    {
                        var cell = sheet.Cell(row, col);
                        string value = cell.GetValue < string>().Trim();

                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        if (value.Contains("Дом"))
                        {
                            house = houses.FirstOrDefault(h => h.Name == value);
                            if (house == null)
                            {
                                house = new House { Name = value };
                                houses.Add(house);
                            }
                        }
                        else if (value.StartsWith("№"))
                        {
                            string flatNum = value.Replace("№", "").Trim();
                            string price = sheet.Cell(row + 1, col).GetValue<string>().Trim(); //

                            house.Flats.Add(new Flat { Number = flatNum, Price = price });
                        }
                    }
                }

            }

            return houses;
        }
    }
}