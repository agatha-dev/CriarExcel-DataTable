using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace CriarExcel_DataTable
{
    class Program
    {
        private static DataTable                     criar;
        
        static void Main(string[] args)
        {
            string caminhoPlanilha = @"C:\Nova pasta\Vendas.xlsx";
            // define o nome do arquivo .xlsx 
            Console.WriteLine("Criação de planilha");
            CriarPlanilha(criar, caminhoPlanilha);
        
        }

        private static void CriarPlanilha(DataTable criar, string caminhoPlanilha)
        {
            // define a licença
            // cria instância do ExcelPackage 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

            Console.WriteLine("Criando Planilha..");
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                Console.WriteLine("Criando DataTable");

                DataTable dataTable = new DataTable("Planilha Excel");
                dataTable.Columns.Add("Nomes", typeof(string));
                dataTable.Columns.Add("Idades", typeof(int));
                dataTable.Columns.Add("Endereço", typeof(string));

                Console.WriteLine("Inserindo linhas");

                dataTable.Rows.Add( "Macoratti", 21, "Meier");
                dataTable.Rows.Add( "Jefferson", 20, "Rj");
                dataTable.Rows.Add( "Janice", 20, "NI");
                dataTable.Rows.Add( "Jessica", 25, "Engenho");
                dataTable.Rows.Add( "Miriam", 48, "Nova Iguaçu");

                //create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                //add all the content from the DataTable, starting at cell A1
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                // Create excel file on physical disk 
                FileStream objFileStrm = File.Create(caminhoPlanilha);
                objFileStrm.Close();

                Console.WriteLine("Escrevendo arquivo excel");

                // Write content to excel file 
                File.WriteAllBytes(caminhoPlanilha, excelPackage.GetAsByteArray());
                excelPackage.SaveAs(new FileInfo(caminhoPlanilha));
            }
            
            Console.WriteLine($"Planilha criada com sucesso");
        }

    }
}
