using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Linq;

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
            Random random = new Random(new Guid().GetHashCode());
            decimal[] randomArray = new decimal[pieceNumer].Select(x => x = random.Next()).ToArray();
            decimal perAtLeast = total * (leastPercet / 100) / pieceNumer;
            decimal avgRemain = total * ((100 - leastPercet) / 100);
            return new decimal[pieceNumer].Select((x, index) => x = avgRemain * (randomArray[index] / randomArray.Sum()) + perAtLeast).ToArray();
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

            decimal tt = 0;
            decimal ts = 0;
            for (int i = 0; i < dayNumber; i++)
            {
                string sheetName = string.Format("{0}月{1}日", month, i + 1);


                NPOIHelper.CrossCloneSheet(tempSheet, book, sheetName);
                NPOI.SS.UserModel.ISheet currentSheet = book.GetSheet(sheetName); // 获取工作表sheet1的索引

                for (int j = 0; j < 6; j++)
                {
                    currentSheet.GetRow(4 * (j + 1)).GetCell(4).SetCellValue(cashArray[i * 6 + j].ToString("C"));
                 
                    currentSheet.GetRow(4 * (j + 1) + 2).GetCell(4).SetCellValue(unionPayArray[i * 6 + j].ToString("C"));
                   
                }
            }

            using (FileStream fileStream = new FileStream(string.Format("/OutPut/家联超市{0}年{1}月日报.xlsx", year, month), FileMode.Create)) {
                book.Write(fileStream); // 写入到本地
            } ;
    
           
            
            return JsonConvert.SerializeObject(new { cash=cashArray,unionPay=unionPayArray});

        }

    }
}
