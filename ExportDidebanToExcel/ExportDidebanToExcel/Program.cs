using CsvHelper;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.XPath;

namespace ExportDidebanToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the MHTML file as a string
            string mhtmlContent = File.ReadAllText("./__TSETMC_. __ دیده بان بازار پیشرفته.html");

            // From File
            var doc = new HtmlDocument();
            doc.LoadHtml(mhtmlContent);

            var stocks = doc.DocumentNode.SelectSingleNode("//*[@id='main']");

            var dideban = new List<Dideban>();
            foreach (var stock in stocks.ChildNodes)
            {
                decimal eps;
                decimal pe;
                decimal.TryParse(stock.ChildNodes[15].InnerText, out eps);
                decimal.TryParse(stock.ChildNodes[16].InnerText, out pe);



                dideban.Add(new Dideban
                {
                    Stock = stock.ChildNodes[0].InnerText,
                    Name = stock.ChildNodes[1].InnerText,
                    Count = long.Parse(stock.ChildNodes[2].InnerText.Replace(",", "")),
                    Volume = long.Parse(stock.ChildNodes[3].InnerText.Replace(",", "")),
                    Cost = long.Parse(stock.ChildNodes[4].InnerText.Replace(",", "")),
                    Yesterday = int.Parse(stock.ChildNodes[5].InnerText.Replace(",", "")),
                    First = int.Parse(stock.ChildNodes[6].InnerText.Replace(",", "")),
                    LastTradePrice = int.Parse(stock.ChildNodes[7].InnerText.Replace(",", "")),
                    LastTradeChange = decimal.Parse(stock.ChildNodes[8].InnerText.Replace(",", "")),
                    LastTradePercent = float.Parse(stock.ChildNodes[9].InnerText.Replace(",", "")),
                    AvgPrice = int.Parse(stock.ChildNodes[10].InnerText.Replace(",", "")),
                    AvgChange = decimal.Parse(stock.ChildNodes[11].InnerText.Replace(",", "")),
                    AvgPercent = float.Parse(stock.ChildNodes[12].InnerText.Replace(",", "")),
                    LowestPrice = int.Parse(stock.ChildNodes[13].InnerText.Replace(",", "")),
                    HighestPrice = int.Parse(stock.ChildNodes[14].InnerText.Replace(",", "")),
                    EPS = eps,
                    PE = pe,
                    BuyCount = int.Parse(stock.ChildNodes[17].InnerText.Replace(",", "")),
                    BuyVolume = long.Parse(stock.ChildNodes[18].InnerText.Replace(",", "")),
                    BuyPrice = int.Parse(stock.ChildNodes[19].InnerText.Replace(",", "")),
                    SellCount = int.Parse(stock.ChildNodes[20].InnerText.Replace(",", "")),
                    SellVolume = long.Parse(stock.ChildNodes[21].InnerText.Replace(",", "")),
                    SellPrice = int.Parse(stock.ChildNodes[22].InnerText.Replace(",", ""))

                });
            }

            var fileName = DateTime.Now.ToString("yy-mm-dd-HH-mm-ss") + ".csv";
            using (var writer = new StreamWriter($"./Export/{fileName}", false, encoding: Encoding.UTF8))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(dideban);
            }
        }
    }
}
