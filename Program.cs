

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
                    await DailyCalls();
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
        private async Task DailyCalls()
        {
            string todayDB = GetToday(3);
            string todaymin = GetToday(1);
            string sql = "";
            try
            {
                
            //CONNEXION
                string connString ="Host=10.255.169.250;Port=9628;Username=cd10c546-53ab-4116-8ce0-80b39fb57242;Password=c7b349376d1294c52878d77144d3196e;Database=cd10c546-53ab-4116-8ce0-80b39fb57242";
               
                await using var conn = new NpgsqlConnection(connString);
                await conn.OpenAsync();
                

               sql =" select count(*) from \"V_1614_TODAS_LAS_LLAMADAS\" where  \"Call start\" ::date = '" + todayDB + "' and \"Campaign Name\" = ''";

                await using var command = new NpgsqlCommand(sql, conn);
                await using var dataReader = await command.ExecuteReaderAsync();

                Console.WriteLine("Realizando consultas...");





                //EXCEL

                //BUSCADOR DE COLUMNAS

                string fileName = "DAILY_CALL.xls";

                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);


                



                    IWorkbook wb = new HSSFWorkbook(fs);

                ISheet ws = wb.GetSheet("Call Disposition");
                IRow row = ws.GetRow(12);//fechas
                int columna = 0;
                bool encontrada = false;
                


                foreach (ICell cell in row)
                {
                    
                    if (cell.ToString()==todaymin)
                    {
                         columna = cell.ColumnIndex;//dia
                        encontrada = true;
                    }



                }
                //Total Leads Loaded 14
                //Total Completes 15
                //Total Contacts 16
                //Total Elegible Contacts(see Call disposition) 17
                //Call Back 18
                if (encontrada)
                {
                    while (await dataReader.ReadAsync()) 
                    {
                        IRow rowEdit = ws.GetRow(14);
                        ICell cellEdit = rowEdit.GetCell(columna);
                        cellEdit.SetBlank();
                        //cellEdit.SetCellValue(dataReader.GetValue(0).ToString());
                        cellEdit.SetCellValue("0000000000");
                    }
                }
                wb.Write(fs);
                //cannot acces a closed file npoi !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                wb.Close();
               
                

                    conn.Close();
                //conn.Close();
            }
            catch (Exception ex)
            {

                throw ex;
            }



        }
        private async Task test2()
        {
            Console.WriteLine("2 ok");
        }
       
        /// <summary>
        /// Coge la fecha de hoy con los diferents formatos
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static string GetToday( int type)
        {
            string fecha = "";
            switch (type)
            {
                case 1:
                    fecha = DateTime.Now.ToString("d/M/yy");
                    break;
                case 2: DateTime.Now.ToString("dd/MM/yyyy");
                    break;

                case 3:
                    fecha = DateTime.Now.ToString("MM/dd/yyyy");
                    break;
                default: fecha = DateTime.Now.ToString();

                break;

            }

            return fecha;
        }
    }
}