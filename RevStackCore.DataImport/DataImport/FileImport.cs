using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using CsvHelper;
using OfficeOpenXml;

namespace RevStackCore.DataImport
{
    public class FileImport : IDataImport
    {

        public IEnumerable<T> ImportCvs<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
        {
            using (var reader = new StreamReader(filePath))
                return importCvsData<T>(reader, hasHeader, matchCase);
        }

        public Task<IEnumerable<T>> ImportCvsAsync<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
        {
            return Task.FromResult(ImportCvs<T>(filePath, hasHeader, matchCase));
        }

        public IEnumerable<T> ImportCvs<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
        {
            using (var reader=new StreamReader(file))
                return importCvsData<T>(reader, hasHeader, matchCase);
        }

        public Task<IEnumerable<T>> ImportCvsAsync<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
        {
            return Task.FromResult(ImportCvs<T>(file, hasHeader, matchCase));
        }

        public IEnumerable<T> ImportExcel<T>(string filePath, bool ignoreHeader=false, bool matchCase=false) where T : class
        {
            bool hasHeader = !(ignoreHeader);
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
                var byteArray = workSheet.ToCsvByteArray(ignoreHeader,matchCase);
                using (var stream = new MemoryStream(byteArray))
                using (var reader = new StreamReader(stream))
                {
                    return importCvsData<T>(reader, hasHeader, false);
                }
            }
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase=false) where T : class
        {
            return Task.FromResult(ImportExcel<T>(filePath, ignoreHeader));
        }

        public IEnumerable<T> ImportExcel<T>(Stream file, bool ignoreHeader = false, bool matchCase=false) where T : class
        {
            bool hasHeader = !(ignoreHeader);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
                var byteArray = workSheet.ToCsvByteArray(ignoreHeader,matchCase);
                using (var stream = new MemoryStream(byteArray))
                using (var reader = new StreamReader(stream))
                {
                    return importCvsData<T>(reader, hasHeader, false);
                }
            }
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase=false) where T : class
        {
            return Task.FromResult(ImportExcel<T>(file, ignoreHeader));
        }

       
        private IEnumerable<T> importCvsData<T>(StreamReader reader, bool hasHeader = true, bool matchCase = false) where T : class
        {
            using (var csv = new CsvReader(reader))
            {
                if (!hasHeader)
                {
                    csv.Configuration.HasHeaderRecord = false;
                }
                else if (matchCase)
                {
                    csv.Configuration.PrepareHeaderForMatch = (string header, int index) => header.ToLower();
                }
                var records = csv.GetRecords<T>();
                records = records.ToList();
                return records;
            }
        }

    }
}
