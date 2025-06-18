using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace Parcer
{
    internal class Program
    {

        static void Main(string[] args)
        {
            var baseDir = @"C:\Users\Korov\Documents\ChillBase";
            var excelFile = Path.Combine(baseDir, "Tech GD Test Parcer.xlsx");
            var outputJson = Path.Combine(baseDir, "balances_config.json");

            var workbook = new XLWorkbook(excelFile);
            var sheet = workbook.Worksheet("Balance");

        }
    }
}
