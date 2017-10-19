using System.Linq;
using System.Collections.Generic;
using LinqToExcel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using LinqToExcel.Logging;
using System.Drawing;

namespace Excel
{
    /// <summary>
    /// Classe responsavel por gerenciar todas as operações feitas usando Excel
    /// </summary>
    public class ExcelFactory  : ExcelQueryFactory
    {
        public ExcelFactory()
        {

        }

        public ExcelFactory(string fileName) : this(fileName, null)
        {
            DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace;
        }

        public ExcelFactory(string fileName, ILogManagerFactory logManagerFactory) : base(fileName, logManagerFactory)
        {
        }

        /// <summary>
        /// Metodo que gera o excel e efetua o download da planilha 
        /// </summary>
        /// <typeparam name="T">Tipo do objeto</typeparam>
        /// <param name="objectCollection">Coleção de objetos</param>
        /// <param name="columns">Define as colunas que serão usadas exemplo(A1:C1), nesse caso 3 colunas serão usadas</param>
        /// <param name="worksheetname">Nome da planilha</param>
        /// <returns></returns>
        public byte[] GenerateExcel<T>(IEnumerable<T> objectCollection, string columns, string worksheetname)
        {
                ExcelPackage pck = new ExcelPackage();
                //Cria o worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(worksheetname);

                // Carrega a coleção de objetos na folha, começando pela celula A1. Imprima os nomes das colunas na linha 1
                ws.Cells["A1"].LoadFromCollection(objectCollection, true);

                //Formate o cabeçalho da coluna 
                ExcelRange rng = ws.Cells[columns];
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Definir padrão para o plano de fundo sólido
                rng.Style.Fill.BackgroundColor.SetColor(Color.Azure);  //Defina a cor para azul escuro
                rng.Style.Font.Color.SetColor(Color.White);


                ExcelRange col = ws.Cells[2, 1, 2 + objectCollection.Count(), 1];
                
                col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                return pck.GetAsByteArray();
        }
    }
}