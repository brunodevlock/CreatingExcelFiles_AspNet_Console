using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Demos\YouTubeDemo.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);

        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            //Create a Excel file
            using var package = new ExcelPackage(file);

            //Add a WorkSheet to the file (aba do Excel)
            var ws = package.Workbook.Worksheets.Add("MainReport");

            //Add load the list of people range begin in the A1 cell
            var range = ws.Cells["A1"].LoadFromCollection(people, true);

            //Preencher as colunas com os dados
            range.AutoFitColumns();

            await package.SaveAsync();

        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() {Id = 1, FirstName = "Tim", LasttName = "Corey"},
                new() {Id = 1, FirstName = "Sue", LasttName = "Storm"},
                new() {Id = 1, FirstName = "Jane", LasttName = "Smith"}
            };

            return output;
        }
    }

   
}
