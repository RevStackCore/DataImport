using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace RevStackCore.DataImport
{
    public interface IDataImport
    {
        IEnumerable<T> ImportCvs<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class;
        Task<IEnumerable<T>> ImportCvsAsync<T>(string filePath, bool hasHeader = true, bool matchCase = false) where T : class;
        IEnumerable<T> ImportCvs<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class;
        Task<IEnumerable<T>> ImportCvsAsync<T>(Stream file, bool hasHeader = true, bool matchCase = false) where T : class;
        IEnumerable<T> ImportExcel<T>(string filePath, bool ignoreHeader = false, bool matchCase=false) where T : class;
        Task<IEnumerable<T>> ImportExcelAsync<T>(string filePath, bool ignoreHeader = false, bool matchCase = false) where T : class;
        IEnumerable<T> ImportExcel<T>(Stream file, bool ignoreHeader = false, bool matchCase = false) where T : class;
        Task<IEnumerable<T>> ImportExcelAsync<T>(Stream file, bool ignoreHeader = false, bool matchCase = false) where T : class;
        void ExportCsv<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class;
        Task ExportCsvAsync<T>(IEnumerable<T> items, string filePath, bool useQuotes = true) where T : class;
        Stream ExportCsvStream<T>(IEnumerable<T> items, bool useQuotes = true) where T : class;
        Task<Stream> ExportCsvStreamAsync<T>(IEnumerable<T> items, bool useQuotes = true) where T : class;

    }
}
