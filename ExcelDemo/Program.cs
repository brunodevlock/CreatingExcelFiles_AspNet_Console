using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Metadata;
using OfficeOpenXml;

namespace ExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Demos\YouTubeDemo.xlsx");

        }

        static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { Id = 1, FirstName = "Tim", LasttName = "Corey"},
                new() { Id = 1, FirstName = "Sue", LasttName = "Storm" },
                new() { Id = 1, FirstName = "Jane", LasttName = "Smith" },
            }
        }
    }

   
}
