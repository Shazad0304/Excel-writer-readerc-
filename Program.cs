using System;  
using System.Collections.Generic;  
using System.Linq;  
using System.Threading.Tasks;  
using System.IO;  
using OfficeOpenXml;  
using System.Text;  

namespace excelReader
{
    class Program
    {
        static void Main(string[] args)
        {
		            string rootFolder = Directory.GetCurrentDirectory()+"\\nik";  
            string fileName = @"ExportCustomers.xlsx";  
  
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
     			using (ExcelPackage package = new ExcelPackage(file))  
            {  
  
                
  
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Customer");  
                int totalRows = 10;  
  
                worksheet.Cells[1, 1].Value = "Customer ID";  
                worksheet.Cells[1, 2].Value = "Customer Name";  
                worksheet.Cells[1, 3].Value = "Customer Email";  
                worksheet.Cells[1, 4].Value = "customer Country";  
                int i = 0;  
                for (int row = 2; row <= totalRows + 1; row++)  
                {  
                    worksheet.Cells[row, 1].Value = "1";  
                    worksheet.Cells[row, 2].Value = "2";  
                    worksheet.Cells[row, 3].Value = "33";  
                    worksheet.Cells[row, 4].Value = "33";  
                    i++;  
                }  
  
                package.Save();   
  
            }  
  
           Console.WriteLine(" Customer list has been exported successfully");
	}
     }
}
