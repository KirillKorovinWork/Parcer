using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace Parсer
{
    class BalanceEntry
    {
        public DateTime Date { get; set; }
        public string Country { get; set; }
        public string BalanceType { get; set; }
    }

    class BalancePeriod
    {
        public string start_date { get; set; }
        public string end_date { get; set; }
        public string balance { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var baseDir = @"C:\Users\Korov\Documents\ChillBase";
            var excelFile = Path.Combine(baseDir, "Tech GD Test Parcer.xlsx");
            var outputJson = Path.Combine(baseDir, "balances_config.json");

            var workbook = new XLWorkbook(excelFile);
            var sheet = workbook.Worksheet("Balance");

            var entries = new List<BalanceEntry>();
            foreach (var r in sheet.RowsUsed().Skip(1))
            {
                var cell = r.Cell(1);
                DateTime date;
                if (cell.DataType == XLDataType.DateTime)
                    date = cell.GetDateTime();
                else if (cell.DataType == XLDataType.Number)
                    date = DateTime.FromOADate(cell.GetDouble());
                else
                    date = DateTime.Parse(cell.GetString().Trim(), CultureInfo.InvariantCulture);

                var mapping = new[] {
                    (Col: 2, Name: "CIS"),
                    (Col: 3, Name: "EU"),
                    (Col: 4, Name: "UK")
                };

                foreach (var m in mapping)
                {
                    var raw = r.Cell(m.Col).GetString().Trim();
                    if (string.IsNullOrEmpty(raw))
                        continue;

                    var types = raw
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim())
                        .Where(t => !string.IsNullOrEmpty(t))
                        .Distinct();

                    foreach (var t in types)
                    {
                        entries.Add(new BalanceEntry
                        {
                            Date = date,
                            Country = m.Name,
                            BalanceType = t
                        });
                    }
                }
            }

            var result = new Dictionary<string, List<BalancePeriod>>();
            foreach (var country in new[] { "CIS", "EU", "UK" })
            {
                var key = country + "_Balance";
                var periods = new List<BalancePeriod>();

                var byType = entries
                    .Where(e => e.Country == country)
                    .GroupBy(e => e.BalanceType);

                foreach (var grp in byType)
                {
                    var dates = grp.Select(e => e.Date)
                                   .OrderBy(d => d)
                                   .ToList();
                    if (dates.Count == 0)
                        continue;

                    var start = dates[0];
                    var end = dates[0];

                    for (int i = 1; i < dates.Count; i++)
                    {
                        var d = dates[i];
                        if ((d - end).Days == 1)
                        {
                            end = d;
                        }
                        else
                        {
                            periods.Add(new BalancePeriod
                            {
                                start_date = start.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
                                end_date = end.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
                                balance = grp.Key
                            });
                            start = end = d;
                        }
                    }
                    periods.Add(new BalancePeriod
                    {
                        start_date = start.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
                        end_date = end.ToString("M/d/yyyy", CultureInfo.InvariantCulture),
                        balance = grp.Key
                    });
                }

                periods = periods
                             .OrderBy(p => DateTime.Parse(p.start_date, CultureInfo.InvariantCulture))
                             .ToList();

                result[key] = periods;
            }

            var json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(outputJson, json);

            Console.WriteLine("balances_config.json сгенерирован: " + outputJson);
        }
    }
}
