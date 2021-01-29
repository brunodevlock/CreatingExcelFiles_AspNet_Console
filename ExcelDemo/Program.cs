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
