using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Linq;
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
                decimal pre = index == 0 ? 0 : randomArray[index - 1];
                decimal current = index == pieceNumer - 1 ? Remain : randomArray[index];
                return (current - pre + perAtLeast) / 100;
            }).ToArray();
        }
        [HttpGet]
        public FileStreamResult GET(decimal cashTotal, decimal unionPayTotal, int leastPercet, int year, int month)
        {
            int dayNumber = DateTime.DaysInMonth(year, month);
            decimal[] cashArray = getRandomArray(cashTotal, dayNumber * 6, leastPercet);
            decimal[] unionPayArray = getRandomArray(unionPayTotal, dayNumber * 6, leastPercet);
            FileStream file = new FileStream("Files/SuperMarketPerDayTemplate.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook tempBook = new XSSFWorkbook(file);
            NPOI.SS.UserModel.ISheet tempSheet = tempBook.GetSheet("data");
            XSSFWorkbook book = new XSSFWorkbook(); // 新建xls工作
            for (int i = 0; i < dayNumber; i++)
            {
                string sheetName = string.Format("{0}月{1}日", month, i + 1);
                NPOIHelper.CrossCloneSheet(tempSheet, book, sheetName);
                NPOI.SS.UserModel.ISheet currentSheet = book.GetSheet(sheetName); // 获取工作表 
                decimal sum = 0;
                for (int j = 0; j < 6; j++)
                {
                    sum += cashArray[i * 6 + j];
                    sum += unionPayArray[i * 6 + j];
                    currentSheet.GetRow(4 * j + 3).GetCell(3).SetCellValue((double)cashArray[i * 6 + j]);
                    currentSheet.GetRow(4 * j + 5).GetCell(3).SetCellValue((double)unionPayArray[i * 6 + j]);
                }
                currentSheet.GetRow(1).GetCell(0).SetCellValue(string.Format("{0}年{1}月{2}日", year, month, i + 1));
                currentSheet.GetRow(26).GetCell(3).SetCellValue(sum.ToString());
            }
            string path = string.Format("{0}\\OutPut\\", Directory.GetCurrentDirectory());
            string fileName = string.Format("家联超市{0}年{1}月日报.xlsx", year, month);
            using (FileStream fileStream = new FileStream(path + fileName, FileMode.Create))
            {
                book.Write(fileStream);
                tempBook.Close();
                book.Close();
            }
            return File(new PhysicalFileProvider(path).GetFileInfo(fileName).CreateReadStream(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
