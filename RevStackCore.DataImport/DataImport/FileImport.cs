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

        public IEnumerable<T> ImportCsv<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
        {
            using (var reader = new StreamReader(filePath))
                return importCvsData<T>(reader, hasHeader, matchCase);
        }

        public Task<IEnumerable<T>> ImportCsvAsync<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class
        {
            return Task.FromResult(ImportCsv<T>(filePath, hasHeader, matchCase));
        }

        public IEnumerable<T> ImportCsv<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
        {
            using (var reader=new StreamReader(file))
                return importCvsData<T>(reader, hasHeader, matchCase);
        }

        public Task<IEnumerable<T>> ImportCsvAsync<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class
        {
            return Task.FromResult(ImportCsv<T>(file, hasHeader, matchCase));
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

        public IEnumerable<T> ImportExcel<T>(string filePath, bool ignoreHeader = false, bool matchCase = false, int worksheetIndex=1) where T : class
        {
            bool hasHeader = !(ignoreHeader);
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.ElementAtOrDefault(worksheetIndex);
                var byteArray = workSheet.ToCsvByteArray(ignoreHeader, matchCase);
                using (var stream = new MemoryStream(byteArray))
                using (var reader = new StreamReader(stream))
                {
                    return importCvsData<T>(reader, hasHeader, false);
                }
            }
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

        public IEnumerable<T> ImportExcel<T>(Stream file, bool ignoreHeader = false, bool matchCase = false, int worksheetIndex=1) where T : class
        {
            bool hasHeader = !(ignoreHeader);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.ElementAtOrDefault(worksheetIndex);
                var byteArray = workSheet.ToCsvByteArray(ignoreHeader, matchCase);
                using (var stream = new MemoryStream(byteArray))
                using (var reader = new StreamReader(stream))
                {
                    return importCvsData<T>(reader, hasHeader, false);
                }
            }
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase = false) where T : class
        {
            return Task.FromResult(ImportExcel<T>(filePath, ignoreHeader, matchCase));
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase = false, int worksheetIndex = 1) where T : class
        {
            return Task.FromResult(ImportExcel<T>(filePath, ignoreHeader, matchCase, worksheetIndex));
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase=false) where T : class
        {
            return Task.FromResult(ImportExcel<T>(file, ignoreHeader, matchCase));
        }

        public Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase = false, int worksheetIndex=1) where T : class
        {
            return Task.FromResult(ImportExcel<T>(file, ignoreHeader, matchCase, worksheetIndex));
        }

        public void ExportCsv<T>(IEnumerable<T> items, string filePath, bool useQuotes=true) where T : class
        {
            using (var writer = new StreamWriter(filePath))
            using (var csv = new CsvWriter(writer))
            {
                if(useQuotes)
                {
                    csv.Configuration.ShouldQuote = (field, context) => true;
                }
                csv.WriteRecords(items);
                writer.Flush();
            }
        }

        public Task ExportCsvAsync<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class
        {
            return Task.Run(() => ExportCsv<T>(items,filePath, useQuotes));
        }

        public Stream ExportCsvStream<T>(IEnumerable<T> items, bool useQuotes = true) where T : class
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var streamWriter = new StreamWriter(memoryStream))
                using (var csv = new CsvWriter(streamWriter))
                {
                    if (useQuotes)
                    {
                        csv.Configuration.ShouldQuote = (field, context) => true;
                    }
                    csv.WriteRecords<T>(items);
                }

                return memoryStream;
            }
        }

        public Task<Stream> ExportCsvStreamAsync<T>(IEnumerable<T> items, bool useQuotes = true) where T : class
        {
            return Task.FromResult(ExportCsvStream<T>(items, useQuotes));
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
