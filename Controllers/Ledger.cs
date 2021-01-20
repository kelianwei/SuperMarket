using System;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860
namespace SuperMarket.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class Ledger : ControllerBase
    {
        // GET: api/<Ledger>
        private decimal[] getRandomArray(decimal total, int pieceNumer, decimal leastPercet)
        {
            //为了数字精度考虑以下数字都放大100倍
            decimal perAtLeast = Math.Round(total * leastPercet / pieceNumer, 0);
            decimal Remain = total * 100 - perAtLeast * pieceNumer;
            Random random = new Random(new Guid().GetHashCode());
            decimal[] randomArray = new decimal[pieceNumer - 1].Select(r => r = random.Next((int)Remain)).ToArray();
            Array.Sort(randomArray);
            return new decimal[pieceNumer].Select((x, index) =>
            {
                var pre = index == 0 ? 0 : randomArray[index - 1];
                var current = index == pieceNumer - 1 ? Remain : randomArray[index];
                return (current - pre + perAtLeast) / 100;
            }).ToArray();
        }
        [HttpPost("file")]
        public string GET(decimal cashTotal, decimal unionPayTotal, int leastPercet, int year, int month)
        {
            int dayNumber = DateTime.DaysInMonth(year, month);
            decimal[] cashArray = getRandomArray(cashTotal, dayNumber * 6, leastPercet);
            decimal[] unionPayArray = getRandomArray(unionPayTotal, dayNumber * 6, leastPercet);
            FileStream file = new FileStream("Files/SuperMarketPerDayTemplate.xlsx", FileMode.Open, FileAccess.Read);
            NPOI.SS.UserModel.ISheet tempSheet = new XSSFWorkbook(file).GetSheet("data");
            XSSFWorkbook book = new XSSFWorkbook(); // 新建xls工作
            for (int i = 0; i < dayNumber; i++)
            {
                string sheetName = string.Format("{0}月{1}日", month, i + 1);
                NPOIHelper.CrossCloneSheet(tempSheet, book, sheetName);
                NPOI.SS.UserModel.ISheet currentSheet = book.GetSheet(sheetName); // 获取工作表sheet1的索引
                decimal sum = 0;
                for (int j = 0; j < 6; j++)
                {
                    sum += cashArray[i * 6 + j];
                    sum += unionPayArray[i * 6 + j];
                    currentSheet.GetRow(4 * j + 3).GetCell(3).SetCellValue(cashArray[i * 6 + j].ToString());
                    currentSheet.GetRow(4 * j + 5).GetCell(3).SetCellValue(unionPayArray[i * 6 + j].ToString());
                }
                currentSheet.GetRow(26).GetCell(3).SetCellValue(sum.ToString());
            }
            using (FileStream fileStream = new FileStream(string.Format("OutPut/家联超市{0}年{1}月日报.xlsx", year, month), FileMode.Create))
            {
                book.Write(fileStream); // 写入到本地
            };
            return JsonConvert.SerializeObject(new { cash = cashArray, unionPay = unionPayArray });
        }
    }
}
