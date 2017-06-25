using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RepositoriesExcel;

namespace BusinessExcel
{
    public class ExcelBusiness
    {
        ExcelRepository excelRepository;

        public ExcelBusiness(string _pathArchivo)
        {
            excelRepository = new ExcelRepository(_pathArchivo);
        }

        public object RetornarRowExcel(string _hoja)
        {
            return excelRepository.CargarArchivo(_hoja);
            //return excelRepository.CargarArchivoFuncionaOK();
        }
    }
}
