using ClosedXML.Excel;
using NmarketTestTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NmarketTestTask.Parsers
{
    public class ExcelParser : IParser
    {
        public IList<House> GetHouses(string path)
        {
            var houses = new List<House>(); // для реальных данных я бы взяла Dictionary<string, House>, тк поиск FirstOrDefault медленнее

            var workbook = new XLWorkbook(path);
            var sheets = workbook.Worksheets;

            foreach (var sheet in sheets)
            {
                House house = null;

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

                        if (value.StartsWith("Дом", StringComparison.OrdinalIgnoreCase))
                        {
                            house = houses.FirstOrDefault(h => h.Name.Equals(value, StringComparison.OrdinalIgnoreCase));
                            if (house == null)
                            {
                                house = new House { Name = value };
                                houses.Add(house);
                            }
                        }
                        else if (value.StartsWith("№") && house != null)
                        {
                            if (row >= lastRow) 
                            { 
                                continue; // если нашелся номер в последней на листе строки без цены
                            }

                            string flatNum = value.Replace("№", "").Trim();
                            string price = sheet.Cell(row + 1, col).GetValue<string>().Trim();

                            house.Flats.Add(new Flat { Number = flatNum, Price = price });
                        }
                    }
                }

            }

            return houses;
        }
    }
}