using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TradeAwayAuto
{
    class Program
    {
        static string SOURCE_PATH = "";
        static string DEST_PATH = "";
       

        static void Main(string[] args)
        {
            if (args.Length >= 0)
            {
                Console.WriteLine("Please enter the path to the source file folder with date:");
                SOURCE_PATH = Console.ReadLine();
                DEST_PATH = @"C:\Users\" + Environment.UserName + @"\Desktop\TradeAway\";
            }
            else
            {
                return;
            }

            Console.WriteLine("Source and destination set but not valid");
            if (Directory.Exists(SOURCE_PATH))
            {
                if (!(Directory.Exists(DEST_PATH)))
                    Directory.CreateDirectory(DEST_PATH);
                Console.WriteLine("Source is a valid directory");
            }
            //put the list of data into excel
            XLWorkbook oWb = new XLWorkbook();
            var oWS = oWb.Worksheets.Add("NBCN Executable(Bulk)");

            List<String> aSheets = new List<string>();

            foreach (String sSheet in aSheets)
            {
                if (sSheet == "NBCN Executable(Bulk")
                {
                    oWS = oWb.Worksheet(sSheet);
                    oWS.Cell(1, 1).Value = "NBCN Trade Advice - ** Please save in CSV format prior to uploading your completed spreadsheet via the NBCN Portal ** ";
                    oWS.Cell(2, 1).Value = "NBCN Client Custody Account Number";
                    oWS.Cell(2, 2).Value = "Commission/ DAP Fee";
                    oWS.Cell(2, 3).Value = "NBCN Broker Settlement Account Number";
                    oWS.Cell(2, 4).Value = "MKT - TU for Cad Equities/ MU Cad Options / NU for American";
                    oWS.Cell(2, 5).Value = "BUY/ SELL";
                    oWS.Cell(2, 6).Value = "Client Side FX";
                    oWS.Cell(2, 7).Value = "DO NOT USE";
                    oWS.Cell(2, 8).Value = "Trade Date - DD-MMM-YY";
                    oWS.Cell(2, 9).Value = "Settlement Date - DD-MMM-YY";
                    oWS.Cell(2, 10).Value = "Coded Trailers #1";
                    oWS.Cell(2, 11).Value = "Coded Trailers #2";
                    oWS.Cell(2, 12).Value = "Coded Trailers #3";
                    oWS.Cell(2, 13).Value = "Open/Close or Tax/OP";
                    oWS.Cell(2, 14).Value = "Client Free-Form Trailers #2";
                    oWS.Cell(2, 15).Value = "Client Free-Form Trailers #3";
                    oWS.Cell(2, 16).Value = "Security Name";
                    oWS.Cell(2, 17).Value = "Ticker, ISM or CUSIP";
                    oWS.Cell(2, 18).Value = "Total Shares";
                    oWS.Cell(2, 19).Value = "Price";
                    oWS.Cell(2, 20).Value = "Fixed Income Accrued Interest - See comments for Trade Away vs Traded with NBCN";
                    oWS.Cell(2, 21).Value = "DO NOT USE";
                    oWS.Cell(2, 22).Value = "Total Settlement Amount - Exclude accured interest for fixed income";
                    oWS.Cell(2, 23).Value = "Offset Free-Form Trailers #1";
                    oWS.Cell(2, 24).Value = "Offset Free-Form Trailers #2";
                    oWS.Cell(2, 25).Value = "Offset Free-Form Trailers #3";
                    oWS.Cell(2, 26).Value = "DO NOT USE";
                    oWS.Cell(2, 27).Value = "Settlement Currency CAD or USD";
                    oWS.Cell(2, 28).Value = "PSET - Provide CDS, DTC, FED, Euroclear";
                }
            }






            //This method is called recursively, handle with care
            /*
                if (File.Exists("\\\\cardinal.dom\\resources\\FIX trade allocations" + DateTime.Now.ToString("yy-MM-dd") + ".xls"))
                {                

                }



                var display = new BrokersCode();
                  var p = new List<BrokersCode>();

                 p = getBrokers();
                  foreach (var item in p)
                  {
                      Console.WriteLine(string.Format("{0} {1}, {2}", item.sName, item.sDtc, item.sCuid));
                  }*/


        }



            public List<BrokersCode> getBrokers()
            {
                List<BrokersCode> brokers = new List<BrokersCode>();

                brokers.Add(new BrokersCode { sName = "TD Securities Inc.", sDtc = 5036, sCuid = "GIST" });
                brokers.Add(new BrokersCode { sName = "CIBC World Mkts", sDtc = 5030, sCuid = "WGDB" });
                brokers.Add(new BrokersCode { sName = "Scotia Capital", sDtc = 5011, sCuid = "SCOT" });
                brokers.Add(new BrokersCode { sName = "Raymond James", sDtc = 5076, sCuid = "MSLT" });
                brokers.Add(new BrokersCode { sName = "Canaccord", sDtc = 5046, sCuid = "CCAM" });
                brokers.Add(new BrokersCode { sName = "BMO Nesbitt Burns", sDtc = 5043, sCuid = "NTDT" });
                brokers.Add(new BrokersCode { sName = "RBC Capital Markets", sDtc = 5002, sCuid = "DOMA" });
                brokers.Add(new BrokersCode { sName = "JP Morgan", sDtc = 352, sCuid = "" });

                return brokers;
            }        
    }
}


