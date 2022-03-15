

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

                //EXCEL1

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
                        break;
                    }

                }
                // Busy    14
                //NA No Answer   15
                //ST Sit Tones   16
                //CB Call Back   17
                //AB Abandome(predictive dialer)    18
                //AM Answering Machine   19
                //BS Cb 2 level  20
                //SA Successful Sales    24
                //QA QA Canceled Sale    25
                //BR Buyers Remorse  26
                //AE Already Enrolled    28
                //NI No interested   29
                //DA Do not GDPR(Contact)   30
                //LR Literature Requested    31
                //BC Business Partner Complaint  32
                //CC Customer Cancelled  34
                //IA Invalid Age 35
                //BP Business Phone  36
                //DC Do not Call 37
                //DS Do not Solicit at All   38
                //IC Ineligible Contract 39
                //PC Not able to contact in calling window(PC)  46
                //DD Deceased    47
                //HU Hang up 48
                //OL Other Languages 49
                //WN / CN   Wrong number    50
                //AS / TS   Agency suppressed/ Robinson 51
                //UA Closed by Kill Fields   52
                //DR Do not GDPR 53
                //CS Client suppressed   54



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