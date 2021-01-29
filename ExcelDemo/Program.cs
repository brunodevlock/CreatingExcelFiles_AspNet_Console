using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.Id}{p.FirstName} {p.LasttName}");
            }

        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];


            int row = 3;
            int col = 1;

            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                PersonModel p = new();
                p.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                p.FirstName = ws.Cells[row, col + 1].Value.ToString();
                p.LasttName = ws.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;


            }

            return output;


        }
    

    private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            //Create a Excel file
            using var package = new ExcelPackage(file);

            //Add a WorkSheet to the file (aba do Excel)
            var ws = package.Workbook.Worksheets.Add("MainReport");

            //Add load the list of people range begin in the A2 cell
            var range = ws.Cells["A2"].LoadFromCollection(people, true);

            //Preencher as colunas com os dados
            range.AutoFitColumns();


            //Formats the header
            ws.Cells["A1"].Value = "Our Cool Report";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20;

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
                new() {Id = 2, FirstName = "Sue", LasttName = "Storm"},
                new() {Id = 3, FirstName = "Jane", LasttName = "Smith"}
            };

            return output;
        }
    }

   
}
