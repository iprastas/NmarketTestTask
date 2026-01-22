using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using NmarketTestTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NmarketTestTask.Parsers
{
    public class HtmlParser : IParser
    {
        public IList<House> GetHouses(string path)
        {
            var houses = new List<House>();  // для реальных данных я бы взяла Dictionary<string, House>, тк поиск FirstOrDefault медленнее

            var doc = new HtmlDocument();
            doc.Load(path);
            
            var tables = doc.DocumentNode.SelectNodes("//table");
            if (tables == null) 
            {  
                return houses; 
            }

            foreach (var table in tables)
            {
                var nodes = table.SelectNodes(".//tbody/tr");
                if (nodes == null) 
                { 
                    continue; 
                }

                foreach (var node in nodes) 
                {
                    var houseNode = node.SelectSingleNode(".//td[contains(@class,'house')]");
                    var flatNumberNode = node.SelectSingleNode(".//td[contains(@class,'number')]");
                    var priceNode = node.SelectSingleNode(".//td[contains(@class,'price')]");
                    if (houseNode == null || flatNumberNode == null || priceNode == null)
                    {
                        continue;
                    }

                    var houseName = houseNode.InnerText.Trim();
                    var flat = new Flat
                    {
                        Number = flatNumberNode.InnerText.Trim(),
                        Price = priceNode.InnerText.Trim()
                    };

                    var house = houses.FirstOrDefault(h => h.Name.Equals(houseName, StringComparison.OrdinalIgnoreCase));
                    if (house == null)
                    {
                        house = new House { Name = houseName};
                        houses.Add(house);
                    }
                    house.Flats.Add(flat);
                }
            }
            
            return houses;
        }
    }
}