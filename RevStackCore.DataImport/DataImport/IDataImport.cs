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
    }
}
