

using Npgsql;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Globalization;

namespace DST_AON_INFORMES
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var mc = new Program();
            await mc.selectMacro();

            Console.Read();
        }

        private async Task selectMacro()
        {
            


            Console.WriteLine("Selecciona la macro (1, 3 o 4):");
            String? macro = Console.ReadLine();

            switch (macro)
            {
                case "1":
                    DailyCalls();
                    break;
                case "2":
                    await test2();
                    break;
                
                default:
                    Console.WriteLine("Opción " + macro + " no disponible.");
                    await selectMacro();
                    break;
            }
        }
        private async Task Connection() 
        {
            //connection string
            string connString =
                "Host=10.255.169.250;Port=9628;Username=cd10c546-53ab-4116-8ce0-80b39fb57242;Password=c7b349376d1294c52878d77144d3196e;Database=cd10c546-53ab-4116-8ce0-80b39fb57242";
            //create instance of database connection
            await using var conn = new NpgsqlConnection(connString);
            

        }
        private void DailyCalls()
        {
            string today = GetToday();
            try
            {
                /*
            //CONNEXION
                string connString =
                "Host=10.255.169.250;Port=9628;Username=cd10c546-53ab-4116-8ce0-80b39fb57242;Password=c7b349376d1294c52878d77144d3196e;Database=cd10c546-53ab-4116-8ce0-80b39fb57242";
               
                await using var conn = new NpgsqlConnection(connString);
                await conn.OpenAsync();
            //PARAMETROS
                Console.WriteLine("Fecha dd/mm/yyyy (enter para fecha de hoy):");
                String? fecha = Console.ReadLine();
                fecha = CheckDate(fecha);
                var fechaFormat = Convert.ToDateTime(fecha).ToString("MM/dd/yyyy");
            //SELECT

                string sql;
                sql =
                        "select \"Call Outcome name\", count(*)  from public.\"V_1274_ALL_CALLS\" where \"Call start\"::date = '"
                        + fechaFormat
                        + "' group by \"Call Outcome name\"";

                await using var command = new NpgsqlCommand(sql, conn);
                await using var dataReader = await command.ExecuteReaderAsync();
                */
                //EXCEL

                //BUSCADOR DE COLUMNAS

                string fileName = "DAILY_CALL.xls";


                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);


                IWorkbook wb = new HSSFWorkbook(fs);



                ISheet ws = wb.GetSheet("Call Disposition");


                IRow row = ws.GetRow(12);
                foreach (ICell cell in row)
                {
                    //{7/5/19}
                    if (cell.NumericCellValue.ToString() == today)
                    {

                    }



                }

                /*
                Workbook workbook = new Workbook();

                workbook.Open(fileName);


                Worksheet ws = wb.Sheet("Call Disposition");

                Row r = ws.Row(13);

                foreach (IXLCell cell in r.Cells()) 
                {

                    if (cell.Value.ToString() == today) 
                    {

                        string letraColumna = cell.Address.ColumnLetter;
                    }

                }
                
                */


                //wb.Save();








                //conn.Close();
            }
            catch (Exception ex)
            {

                throw;
            }



        }
        private async Task test2()
        {
            Console.WriteLine("2 ok");
        }
        private static string CheckDate(string? fecha)
        {
            bool chValidity = false;
            while (!chValidity)
            {
                DateTime d;
                if (string.IsNullOrEmpty(fecha))
                {
                    fecha = DateTime.Now.ToString("dd/MM/yyyy");
                }
                else
                {
                    //CHECK FECHA FORMAT
                    chValidity = DateTime.TryParseExact(
                        fecha,
                        "dd/MM/yyyy",
                        null,
                        DateTimeStyles.None,
                        out d
                    );
                    if (!chValidity)
                    {
                        Console.WriteLine(
                            "El formato de fecha introducido " + fecha + " no es correcto."
                        );
                        Console.WriteLine("Fecha dd/mm/yyyy (enter para fecha de hoy):");
                        fecha = Console.ReadLine();
                    }
                }
            }
            return fecha;
        }
        private static string GetToday()
        {
          
              string fecha = DateTime.Now.ToString("dd/MM/yyyy");
               
            return fecha;
        }
    }
}