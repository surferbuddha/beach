using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.Odbc;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;

namespace _12MosByDist
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    
    public partial class MainWindow : System.Windows.Window
    {
        OdbcConnection mySQLcon = new OdbcConnection("DSN=DSN;UID=sa;PWD=password;");
        ///string StartDate;
        DateTime? EndDate;


        public MainWindow()
        {
            InitializeComponent();
        }

        private void getDates(object sender, RoutedEventArgs e)
        {



            EndDate = dp1.SelectedDate.Value.Date;
            string foo1 = EndDate.Value.Date.ToString("MM");
            //System.Windows.Forms.MessageBox.Show(foo1, "foo");

            //Rolling12MosNoDistDollars(foo1);
            //getList();
            //Rolling12MosNoDistDollars();
            //Rolling12MosNoDistUnits();
            //Rolling12_by_Dis_Masters();
            //System.Windows.Forms.MessageBox.Show("foo", date.ToString());
            createViews(foo1);
        }

        public void createViews(string foo)
        {
            int endMonth = Convert.ToInt32(foo);
            int startMonth = 0;
            //string startYear = "2018";
            //string endYear = "2019";
            if (endMonth < 12)
            {
                startMonth = endMonth + 1;
            }
            else
            {
                startMonth = 1;
            }
            string StartMonth = DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(startMonth);
            string EndMonth = DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(endMonth);
            string currYear = DateTime.Now.ToString("yyyy");
            string lastYear = DateTime.Now.AddYears(-1).ToString("yyyy");
            string lastTwoYear = DateTime.Now.AddYears(-2).ToString("yyyy");
            System.Windows.Forms.MessageBox.Show("Start is:" + StartMonth.ToString() + " " + lastYear + "and End Month is" + EndMonth.ToString() + " " + currYear, "StartMonth");
            int StartMonthDay = DateTime.DaysInMonth(Convert.ToInt32(lastTwoYear), startMonth);
            int EndMonthDay = DateTime.DaysInMonth(Convert.ToInt32(lastYear), endMonth);
            
            endMonth.ToString("00");
            ////////////////////////////////////////////////////////////////////////
            //UNITS
            //////////////////////////////////////////////////////////////////////
            string myQueryDrop2 = "USE ESTrading " +
   "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Units' and type='v') " +
   "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Units  ";

            string myQuery2 =
               "CREATE VIEW " + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Units AS" +
               "(" +
               "SELECT tblARCustomers.CustCodeSysPro as custcode, ProdCode as sc , inv.Description as descr, sum(tblOEInvoiceFooterHistory.Quantity) as value " +
               "FROM tblARCustomers " +
               "INNER JOIN tblOEInvoiceHistory  ON tblARCustomers.CustCode = tblOEInvoiceHistory.CustCode " +
               "INNER JOIN tblOEInvoiceFooterHistory  ON tblOEInvoiceFooterHistory.Invoice = tblOEInvoiceHistory.Invoice " +
               "INNER JOIN SYSPROFastTrack_E.dbo.InvMaster inv ON tblOEInvoiceFooterHistory.ProdCode = inv.StockCode  COLLATE Latin1_General_100_BIN " +
               "where (tblOEInvoiceHistory.InvDate >= " +
               "'" + lastTwoYear + "-" + startMonth.ToString("00") + "-01' " +
               "AND tblOEInvoiceHistory.InvDate <= '2017-12-31') and tblOEInvoiceFooterHistory.ProdCode <> 'TJX' " +
               "AND ProdCode IN(SELECT StockCode COLLATE Latin1_General_100_BIN FROM SYSPROFastTrack_E.dbo.InvMaster) " +
                "GROUP BY tblARCustomers.CustCodeSysPro, tblOEInvoiceFooterHistory.ProdCode, inv.Description" +
               ") ";
            string myQueryDropThisYear2 = "USE ESTrading " +
              "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastYear + "" + "To" + EndMonth + "" + currYear + "_Units' and type='v') " +
              "DROP VIEW " + StartMonth + "" + lastYear + "" + "To" + EndMonth + "" + currYear + "_Units  ";

            string myQueryThisYear2 = "CREATE VIEW " + StartMonth + "" + lastYear + "" + "To" + EndMonth + "" + currYear + "_Units AS( " +
                "SELECT " +
                "SorMasterRep.Customer as custcode, " +
                "SorDetailRep.StockCode as sc " +
                ", InvMaster.Description as descr, " +
                "sum(SorDetailRep.OrderQty) as value " +
                "FROM(SYSPROFastTrack_E.dbo.SorDetailRep SorDetailRep " +
                "INNER JOIN SYSPROFastTrack_E.dbo.InvWarehouse InvWarehouse " +
                "ON(SorDetailRep.StockCode = InvWarehouse.StockCode) " +
                "AND(SorDetailRep.Warehouse = InvWarehouse.Warehouse) " +
                ") " +
                "INNER JOIN SYSPROFastTrack_E.dbo.SorMasterRep SorMasterRep " +
                "ON(SorDetailRep.Invoice = SorMasterRep.InvoiceNumber) " +
                "AND(SorDetailRep.SalesOrder = SorMasterRep.SalesOrder) " +
                "AND ((TrnMonth >=" + startMonth + " AND TrnYear = " + lastYear + ") OR (TrnMonth <= " + endMonth + " AND TrnYear =" + currYear + "))" +
                "INNER JOIN SYSPROFastTrack_E.dbo.InvMaster " +
                "ON  SYSPROFastTrack_E.dbo.InvMaster.StockCode = InvWarehouse.StockCode " +
                "WHERE SorDetailRep.StockCode <> 'ZPALLET' " +
                "AND Invoice IN (SELECT Invoice FROM SYSPROFastTrack_E.dbo.ArInvoice Where InvoiceDate between '" + lastYear + "-" + startMonth.ToString("00") + "-01' " +
                "AND '" + currYear + "-" + endMonth.ToString("00") + "-" + EndMonthDay + "')" +
                "AND SorDetailRep.StockCode NOT LIKE '%E' " +
                "AND(SorMasterRep.Customer <> ' ' AND SorMasterRep.Customer NOT LIKE '%15%') " +
                "GROUP BY SorMasterRep.Customer, SorDetailRep.StockCode " +
                ", InvMaster.Description)";


            string myQueryDropLastYear2 = "USE ESTrading " +
              "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Units' and type='v') " +
              "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Units  ";


            string myQueryLastYear2 = "CREATE VIEW " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Units  AS " +
                "(" +
                "SELECT custcode, sc, descr, sum(value) as value " +
                "FROM  " +
                "(" +
                " select custcode COLLATE Latin1_General_100_BIN as custcode, sc COLLATE Latin1_General_100_BIN as sc, " +
                "descr COLLATE Latin1_General_100_BIN as descr, value  " +
                " FROM  " +
                " Jan2018ToJan2018_Units  " +
                " UNION ALL  " +
                " select custcode, sc, descr, value  " +
                " FROM  " +
                StartMonth + "" + lastTwoYear + "" + "ToDec2017_Units  " +
                "  ) as foo " +
                " GROUP BY custcode, sc, descr " +
                ") ";


            string myQueryDropBothYears2 = "USE ESTrading " +
             "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Units' and type='v') " +
             "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Units  ";


            string myQueryBothYears2 = "CREATE VIEW " + StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + currYear + "_Units AS ( " +
                "SELECT foo2.custcode, foo2.sc, foo2.descr, coalesce(nu.value,0) as rep2017, coalesce(nu2.value,0) as rep2018 " +
                "FROM " +
                "(" +
                "SELECT custcode, sc, descr " +
                "FROM " +
                "( " +
                "Select custcode COLLATE Latin1_General_100_BIN as custcode, sc COLLATE Latin1_General_100_BIN as sc, " +
                " descr COLLATE Latin1_General_100_BIN as descr " +
                " , value " +
                "FROM " +
                StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + lastYear + "_Units" +
                " UNION  " +
                " select * " +
                " FROM  " +
                StartMonth + "" + lastYear + "To" + EndMonth + "" + currYear + "_Units" +
                " ) " +
                " as foo  " +
                " GROUP BY custcode, sc, descr " +
                ") " +
                "as foo2  " +
                "LEFT OUTER JOIN " + StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + lastYear + "_Units nu ON(foo2.custcode = nu.custcode and foo2.sc = nu.sc) " +
                "LEFT OUTER JOIN " + StartMonth + "" + lastYear + "To" + EndMonth + "" + currYear + "_Units nu2 ON(foo2.custcode = nu2.custcode and foo2.sc = nu2.sc) " +
                ")";

            string myQueryFullData2 = "USE ESTrading SELECT * FROM " + StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + currYear + "_Units ORDER BY custcode, sc";



            mySQLcon.Open();
            OdbcCommand myComDrop2 = new OdbcCommand(myQueryDrop2, mySQLcon);
            OdbcDataReader myReaderDrop2 = myComDrop2.ExecuteReader();
            OdbcCommand myCom2 = new OdbcCommand(myQuery2, mySQLcon);
            OdbcDataReader myReader2 = myCom2.ExecuteReader();
            OdbcCommand myComCurrDrop2 = new OdbcCommand(myQueryDropThisYear2, mySQLcon);
            OdbcDataReader myReaderCurrDrop2 = myComCurrDrop2.ExecuteReader();
            OdbcCommand myComCurr2 = new OdbcCommand(myQueryThisYear2, mySQLcon);
            OdbcDataReader myReaderCurr2 = myComCurr2.ExecuteReader();
            OdbcCommand myComDropLastYear2 = new OdbcCommand(myQueryDropLastYear2, mySQLcon);
            OdbcDataReader myReaderDropLastYear2 = myComDropLastYear2.ExecuteReader();
            OdbcCommand myComLastYear2 = new OdbcCommand(myQueryLastYear2, mySQLcon);
            OdbcDataReader myReaderLastYear2 = myComLastYear2.ExecuteReader();
            OdbcCommand myComDropBothYears2 = new OdbcCommand(myQueryDropBothYears2, mySQLcon);
            OdbcDataReader myReaderDropBothYears2 = myComDropBothYears2.ExecuteReader();
            OdbcCommand myComBothYears2 = new OdbcCommand(myQueryBothYears2, mySQLcon);
            OdbcDataReader myReaderBothYears2 = myComBothYears2.ExecuteReader();
            OdbcCommand myComFullData2 = new OdbcCommand(myQueryFullData2, mySQLcon);
            OdbcDataReader myReaderFullData2 = myComFullData2.ExecuteReader();
            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Load(myReaderFullData2);


            ////////////////////////////////////////////////////////////////////////////////
            ///END UNITS
            /////////////////////////////////////////////////////////////////////////////////
            ///
            /// 
            ///////////////////////////////////////////////////////////////////////////////
            //DOLLARS
            ///////////////////////////////////////////////////////////////////////////////////
            string myQueryDrop = "USE ESTrading " +
               "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Dollars' and type='v') " +
               "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Dollars  ";

            string myQuery =
               "CREATE VIEW " + StartMonth + "" + lastTwoYear + "" + "ToDec2017_Dollars AS" +
               "(" +
               "SELECT tblARCustomers.CustCodeSysPro as custcode, ProdCode as sc , inv.Description as descr, sum(tblOEInvoiceFooterHistory.Extension) as value " +
               "FROM tblARCustomers " +
               "INNER JOIN tblOEInvoiceHistory  ON tblARCustomers.CustCode = tblOEInvoiceHistory.CustCode "+
               "INNER JOIN tblOEInvoiceFooterHistory  ON tblOEInvoiceFooterHistory.Invoice = tblOEInvoiceHistory.Invoice " +
               "INNER JOIN SYSPROFastTrack_E.dbo.InvMaster inv ON tblOEInvoiceFooterHistory.ProdCode = inv.StockCode  COLLATE Latin1_General_100_BIN "+
               "where (tblOEInvoiceHistory.InvDate >= "+
               "'"+lastTwoYear+"-"+ startMonth.ToString("00")+"-01' "+
               "AND tblOEInvoiceHistory.InvDate <= '2017-12-31') and tblOEInvoiceFooterHistory.ProdCode <> 'TJX' "+
               "AND ProdCode IN(SELECT StockCode COLLATE Latin1_General_100_BIN FROM SYSPROFastTrack_E.dbo.InvMaster) "+
                "GROUP BY tblARCustomers.CustCodeSysPro, tblOEInvoiceFooterHistory.ProdCode, inv.Description" +
               ") ";

            string myQueryDropThisYear = "USE ESTrading " +
              "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastYear + "" + "To" + EndMonth + "" + currYear + "_Dollars' and type='v') " +
              "DROP VIEW " + StartMonth + "" + lastYear + "" + "To" + EndMonth + "" + currYear + "_Dollars  ";

            string myQueryThisYear = "CREATE VIEW " + StartMonth + "" + lastYear + "" + "To"+EndMonth+""+currYear+"_Dollars AS( " +
                "SELECT "+ 
                "SorMasterRep.Customer as custcode, "+
                "SorDetailRep.StockCode as sc "+
                ", InvMaster.Description as descr, "+
                "sum((SorDetailRep.Price * SorDetailRep.OrderQty)) - sum((SorDetailRep.Price * SorDetailRep.OrderQty) * (SorDetailRep.DiscPct1 / 100))" +
                " as value " +
                "FROM(SYSPROFastTrack_E.dbo.SorDetailRep SorDetailRep "+
                "INNER JOIN SYSPROFastTrack_E.dbo.InvWarehouse InvWarehouse "+
                "ON(SorDetailRep.StockCode = InvWarehouse.StockCode) "+
                "AND(SorDetailRep.Warehouse = InvWarehouse.Warehouse) "+
                ") "+
                "INNER JOIN SYSPROFastTrack_E.dbo.SorMasterRep SorMasterRep "+
                "ON(SorDetailRep.Invoice = SorMasterRep.InvoiceNumber) "+
                "AND(SorDetailRep.SalesOrder = SorMasterRep.SalesOrder) "+
                "AND ((TrnMonth >="+startMonth+" AND TrnYear = "+lastYear+") OR (TrnMonth <= "+endMonth+" AND TrnYear ="+currYear+"))"+
                "INNER JOIN SYSPROFastTrack_E.dbo.InvMaster "+
                "ON  SYSPROFastTrack_E.dbo.InvMaster.StockCode = InvWarehouse.StockCode "+
                "WHERE SorDetailRep.StockCode <> 'ZPALLET' "+
                "AND Invoice IN (SELECT Invoice FROM SYSPROFastTrack_E.dbo.ArInvoice Where InvoiceDate between '"+lastYear+"-"+ startMonth.ToString("00")+"-01' "+
                "AND '"+ currYear + "-" + endMonth.ToString("00") + "-" + EndMonthDay +"')"+
                "AND SorDetailRep.StockCode NOT LIKE '%E' "+
                "AND(SorMasterRep.Customer <> ' ' AND SorMasterRep.Customer NOT LIKE '%15%') "+
                "GROUP BY SorMasterRep.Customer, SorDetailRep.StockCode "+
                ", InvMaster.Description)";


            string myQueryDropLastYear = "USE ESTrading " +
              "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Dollars' and type='v') " +
              "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Dollars  ";


            string myQueryLastYear = "CREATE VIEW "+ StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + lastYear + "_Dollars  AS " +
                "("+
                "SELECT custcode, sc, descr, sum(value) as value "+
                "FROM  "+
                "("+
                " select custcode COLLATE Latin1_General_100_BIN as custcode, sc COLLATE Latin1_General_100_BIN as sc, "+
                "descr COLLATE Latin1_General_100_BIN as descr, value  "+
                " FROM  "+
                " Jan2018ToJan2018_Dollars  "+
                " UNION ALL  "+
                " select custcode, sc, descr, value  "+
                " FROM  "+
                StartMonth + "" + lastTwoYear + "" + "ToDec2017_Dollars  "+
                "  ) as foo "+
                " GROUP BY custcode, sc, descr "+
                ") ";


            string myQueryDropBothYears = "USE ESTrading " +
             "if exists(select 1 from sys.views where name='" + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Dollars' and type='v') " +
             "DROP VIEW " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Dollars  ";


            string myQueryBothYears = "CREATE VIEW " + StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + currYear + "_Dollars AS ( " +
                "SELECT foo2.custcode, foo2.sc, foo2.descr, coalesce(nu.value,0) as rep2017, coalesce(nu2.value,0) as rep2018 " +
                "FROM " +
                "(" +
                "SELECT custcode, sc, descr " +
                "FROM " +
                "( " +
                "Select custcode COLLATE Latin1_General_100_BIN as custcode, sc COLLATE Latin1_General_100_BIN as sc, " +
                " descr COLLATE Latin1_General_100_BIN as descr " +
                " , value " +
                "FROM " +
                StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + lastYear + "_Dollars" +
                " UNION  " +
                " select * " +
                " FROM  " +
                StartMonth + "" + lastYear + "To" + EndMonth + "" + currYear + "_Dollars" +
                " ) " +
                " as foo  " +
                " GROUP BY custcode, sc, descr " +
                ") " +
                "as foo2  " +
                "LEFT OUTER JOIN " + StartMonth + "" + lastTwoYear + "To" + EndMonth + "" + lastYear + "_Dollars nu ON(foo2.custcode = nu.custcode and foo2.sc = nu.sc) " +
                "LEFT OUTER JOIN " + StartMonth + "" + lastYear + "To" + EndMonth + "" + currYear + "_Dollars nu2 ON(foo2.custcode = nu2.custcode and foo2.sc = nu2.sc) " +
                ")";

                string myQueryFullData = "USE ESTrading SELECT rep2017, rep2018 FROM "+ StartMonth+""+lastTwoYear+ "To"+EndMonth+""+currYear+"_Dollars ORDER BY custcode, sc";
            


                //mySQLcon.Open();
                OdbcCommand myComDrop = new OdbcCommand(myQueryDrop, mySQLcon);
                OdbcDataReader myReaderDrop = myComDrop.ExecuteReader();
                OdbcCommand myCom = new OdbcCommand(myQuery, mySQLcon);
                OdbcDataReader myReader = myCom.ExecuteReader();
                OdbcCommand myComCurrDrop = new OdbcCommand(myQueryDropThisYear, mySQLcon);
                OdbcDataReader myReaderCurrDrop = myComCurrDrop.ExecuteReader();
                OdbcCommand myComCurr = new OdbcCommand(myQueryThisYear, mySQLcon);
                OdbcDataReader myReaderCurr = myComCurr.ExecuteReader();
                OdbcCommand myComDropLastYear = new OdbcCommand(myQueryDropLastYear, mySQLcon);
                OdbcDataReader myReaderDropLastYear = myComDropLastYear.ExecuteReader();
                OdbcCommand myComLastYear = new OdbcCommand(myQueryLastYear, mySQLcon);
                OdbcDataReader myReaderLastYear = myComLastYear.ExecuteReader();
                OdbcCommand myComDropBothYears = new OdbcCommand(myQueryDropBothYears, mySQLcon);
                OdbcDataReader myReaderDropBothYears = myComDropBothYears.ExecuteReader();
                OdbcCommand myComBothYears = new OdbcCommand(myQueryBothYears, mySQLcon);
                OdbcDataReader myReaderBothYears = myComBothYears.ExecuteReader();
                OdbcCommand myComFullData = new OdbcCommand(myQueryFullData, mySQLcon);
                OdbcDataReader myReaderFullData = myComFullData.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(myReaderFullData);
            ///////////////////////////////////////////////////////////////////////////
            ///END DOLLARS
            //////////////////////////////////////////////////////////////
                mySQLcon.Close();

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                // if you want to make excel visible to user, set this property to true, false by default
                excelApp.Visible = true;

                // open an existing workbook
                string workbookPath = "Z:\\Reports\\SysPro Reports\\BASE\\YTDRollingByDist.xlsx";
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                string currentSheet1 = "12_Mos_by_Dist_auto";
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet =
                    (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet1);

                excelWorksheet.Activate();
                Microsoft.Office.Interop.Excel.Range rng = excelWorksheet.Range["A1"];
                Microsoft.Office.Interop.Excel.Range rng1 = excelWorksheet.Range["B1"];
                Microsoft.Office.Interop.Excel.Range rng4 = excelWorksheet.Range["C1"];
                Microsoft.Office.Interop.Excel.Range rng5 = excelWorksheet.Range["D1"];
                Microsoft.Office.Interop.Excel.Range rng6 = excelWorksheet.Range["E1"];
                Microsoft.Office.Interop.Excel.Range rng7 = excelWorksheet.Range["F1"];
                Microsoft.Office.Interop.Excel.Range rng8 = excelWorksheet.Range["G1"];
                Microsoft.Office.Interop.Excel.Range rng11 = excelWorksheet.Range["H1"];
                Microsoft.Office.Interop.Excel.Range rng12 = excelWorksheet.Range["I1"];
                Microsoft.Office.Interop.Excel.Range rng13 = excelWorksheet.Range["J1"];
                Microsoft.Office.Interop.Excel.Range rng14 = excelWorksheet.Range["K1"];



                rng.Value = "Customer";
                rng1.Value = "Stock Code";
                rng4.Value = "Description";
                rng5.Value = StartMonth+""+lastTwoYear+"-"+EndMonth+""+lastYear+"Units";
                rng6.Value = StartMonth + "" + lastYear + "-" + EndMonth + "" + currYear + "Units";
                rng7.Value = "Diff1";
                rng8.Value = "Perc1";
                rng11.Value = StartMonth + "" + lastTwoYear + "-" + EndMonth + "" + lastYear + "Dollars";
                rng12.Value = StartMonth + "" + lastYear + "-" + EndMonth + "" + currYear + "Dollars";
                rng13.Value = "Diff2";
                rng14.Value = "Perc2";


                int tooArr = dt2.Rows.Count + 2;
                object[,] arr = new object[dt2.Rows.Count, dt2.Columns.Count];
                for (int r = 0; r < dt2.Rows.Count; r++)
                {
                    DataRow dr = dt2.Rows[r];
                    for (int c = 0; c < dt2.Columns.Count; c++)
                    {
                        arr[r, c] = dr[c];
                    }
                }
                Microsoft.Office.Interop.Excel.Range rng2 = excelWorksheet.Range["A3:E" + tooArr];
                Microsoft.Office.Interop.Excel.Range rng15 = excelWorksheet.Range["H3:I" + tooArr];

                object[,] arr4 = new object[dt.Rows.Count, dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    DataRow dr = dt.Rows[r];
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        arr4[r, c] = dr[c];
                    }
                }


                int tooArr2 = tooArr + 2;
                rng2.Value = arr;
                rng15.Value = arr4;
            //SUM VALUES

            Microsoft.Office.Interop.Excel.Range rng9 = excelWorksheet.Range["F3:F" + tooArr];
            Microsoft.Office.Interop.Excel.Range rng3 = excelWorksheet.Range["A1:K" + tooArr];

            object[,] arr2 = new object[tooArr, 1];

            int itr = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr2[r, 0] = "=SUM(E" + itr + "-D" + itr + ")";

                itr++;

            }

            object[,] arr5 = new object[tooArr, 1];

            int itr3 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr5[r, 0] = "=SUM(I" + itr3 + "-H" + itr3 + ")";

                itr3++;

            }
            Microsoft.Office.Interop.Excel.Range rng10 = excelWorksheet.Range["G3:G" + tooArr];
            Microsoft.Office.Interop.Excel.Range rng16 = excelWorksheet.Range["J3:J" + tooArr];
            rng16.Value = arr5;

            object[,] arr3 = new object[tooArr, 1];

            int itr2 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr3[r, 0] = "=IFERROR(IF(D" + itr2 + "=0,F" + itr2 + "/E" + itr2 + ",F" + itr2 + "/D" + itr2 + "), \"No Sales\")";

                //arr3[r, 0] = "foo";

                itr2++;

            }

            Microsoft.Office.Interop.Excel.Range rng17 = excelWorksheet.Range["K3:K" + tooArr];
            object[,] arr6 = new object[tooArr, 1];

            int itr4 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr6[r, 0] = "=IFERROR(IF(H" + itr4 + "=0,J" + itr4 + "/I" + itr4 + ",J" + itr4 + "/H" + itr4 + "), \"No Sales\")";

                //arr3[r, 0] = "foo";

                itr4++;

            }
            //= IFERROR(IF(D3 = 0, F3 / E3, F3 / D3), 'No Sales')
            //= IFERROR(IF(D2 = 0, F2 / E2, F2 / D2), "No Sales")

            rng9.Value = arr2;
            rng10.NumberFormat = "###,##%";
            rng10.Value = arr3;
            rng17.Value = arr6;
            rng17.NumberFormat = "###,##%";
            rng15.NumberFormat = "$#,##0.00";
            rng16.NumberFormat = "$#,##0.00";
            int[] fields = new int[] { 4, 5, 6, 8, 9, 10 };

            rng3.Subtotal(1, Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum,
                fields, Microsoft.Office.Interop.Excel.XlSummaryRow.xlSummaryBelow);

            excelWorksheet.Columns.AutoFit();
            /*
            int RowNumber = 1;
            int iRow = 1;
            int iLastRow = 6000;
            Microsoft.Office.Interop.Excel.Range RowFarben = null;
            // Wir schreiben nur Nummern rein - kein Gebrauch für Formular:
            for (int iRows = iRow; iRows <= iLastRow; iRows++) // Alle Spalten
            {
                excelWorksheet.Cells[iRows, 1] = RowNumber.ToString(); ; // Start at 1 in later row.
                RowNumber++;
                // Farben lassen:
                if ((RowNumber % 2) == 0)
                {
                    RowFarben = excelWorksheet.Range[excelWorksheet.Cells[iRows, 1], excelWorksheet.Cells[iRows, 12]];
                    RowFarben.Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightBlue;
                }
            }

            */
            var currEndDate = EndDate.Value.Date.ToString("MMMM");
            excelWorkbook.SaveAs("Z:\\Reports\\SysPro Reports\\2019-" + currEndDate + "\\YTDRollingByDist.xlsx");
            excelApp.Quit();
            System.Threading.Thread.Sleep(5000);

            mySQLcon.Close();

            createViewsMasters(foo);
        }
        /// <summary>
        /// ///////////////////////////////////////////////
        /// </summary> MASTERS
        /// ///////////////////////////////////////////////////////////////////
        /// <param name="foo"></param>
        public void createViewsMasters(string foo)
        {
            int endMonth = Convert.ToInt32(foo);
            int startMonth = 0;
            //string startYear = "2018";
            //string endYear = "2019";
            if (endMonth < 12)
            {
                startMonth = endMonth + 1;
            }
            else
            {
                startMonth = 1;
            }
            string StartMonth = DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(startMonth);
            string EndMonth = DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(endMonth);
            string currYear = DateTime.Now.ToString("yyyy");
            string lastYear = DateTime.Now.AddYears(-1).ToString("yyyy");
            string lastTwoYear = DateTime.Now.AddYears(-2).ToString("yyyy");
            System.Windows.Forms.MessageBox.Show("Start is:" + StartMonth.ToString() + " " + lastYear + "and End Month is" + EndMonth.ToString() + " " + currYear, "StartMonth");
            int StartMonthDay = DateTime.DaysInMonth(Convert.ToInt32(lastTwoYear), startMonth);
            int EndMonthDay = DateTime.DaysInMonth(Convert.ToInt32(lastYear), endMonth);

            endMonth.ToString("00");
            ////////////////////////////////////////////////////////////////////////
            //UNITS
            //////////////////////////////////////////////////////////////////////

            string myQuery2 = "USE Estrading SELECT ccm.CustMaster, sc, descr, " +
            "sum(rep2017), sum(rep2018) " +
            "FROM " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Units fm " +
            "INNER JOIN CustCodeMasters ccm " +
            "ON fm.custcode = ccm.CustCode " +
            "GROUP BY ccm.CustMaster, sc, descr ";

            mySQLcon.Open();
            OdbcCommand myCom2 = new OdbcCommand(myQuery2, mySQLcon);
            OdbcDataReader myReader2 = myCom2.ExecuteReader();

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Load(myReader2);


            ////////////////////////////////////////////////////////////////////////////////
            ///END UNITS
            /////////////////////////////////////////////////////////////////////////////////
            ///
            /// 
            ///////////////////////////////////////////////////////////////////////////////
            //DOLLARS
            ///////////////////////////////////////////////////////////////////////////////////
            string myQuery = "USE Estrading SELECT  " +
            "sum(rep2017), sum(rep2018) " +
            "FROM " + StartMonth + "" + lastTwoYear + "" + "To" + EndMonth + "" + currYear + "_Dollars fm " +
            "INNER JOIN CustCodeMasters ccm " +
            "ON fm.custcode = ccm.CustCode " +
            "GROUP BY ccm.CustMaster, sc, descr ";
            //mySQLcon.Open();
            OdbcCommand myCom = new OdbcCommand(myQuery, mySQLcon);
            OdbcDataReader myReader = myCom.ExecuteReader();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(myReader);
            ///////////////////////////////////////////////////////////////////////////
            ///END DOLLARS
            //////////////////////////////////////////////////////////////
            mySQLcon.Close();

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // if you want to make excel visible to user, set this property to true, false by default
            excelApp.Visible = true;

            // open an existing workbook
            var currEndDate = EndDate.Value.Date.ToString("MMMM");
            string workbookPath = "Z:\\Reports\\SysPro Reports\\2019-"+currEndDate+"\\YTDRollingByDist.xlsx";
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);

            Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;

            string currentSheet1 = "12_Mos_by_Dist_auto_Masters";
            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet =
                (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet1);

            excelWorksheet.Activate();
            Microsoft.Office.Interop.Excel.Range rng = excelWorksheet.Range["A1"];
            Microsoft.Office.Interop.Excel.Range rng1 = excelWorksheet.Range["B1"];
            Microsoft.Office.Interop.Excel.Range rng4 = excelWorksheet.Range["C1"];
            Microsoft.Office.Interop.Excel.Range rng5 = excelWorksheet.Range["D1"];
            Microsoft.Office.Interop.Excel.Range rng6 = excelWorksheet.Range["E1"];
            Microsoft.Office.Interop.Excel.Range rng7 = excelWorksheet.Range["F1"];
            Microsoft.Office.Interop.Excel.Range rng8 = excelWorksheet.Range["G1"];
            Microsoft.Office.Interop.Excel.Range rng11 = excelWorksheet.Range["H1"];
            Microsoft.Office.Interop.Excel.Range rng12 = excelWorksheet.Range["I1"];
            Microsoft.Office.Interop.Excel.Range rng13 = excelWorksheet.Range["J1"];
            Microsoft.Office.Interop.Excel.Range rng14 = excelWorksheet.Range["K1"];



            rng.Value = "Customer";
            rng1.Value = "Stock Code";
            rng4.Value = "Description";
            rng5.Value = StartMonth + "" + lastTwoYear + "-" + EndMonth + "" + lastYear + "Units";
            rng6.Value = StartMonth + "" + lastYear + "-" + EndMonth + "" + currYear + "Units";
            rng7.Value = "Diff1";
            rng8.Value = "Perc1";
            rng11.Value = StartMonth + "" + lastTwoYear + "-" + EndMonth + "" + lastYear + "Dollars";
            rng12.Value = StartMonth + "" + lastYear + "-" + EndMonth + "" + currYear + "Dollars";
            rng13.Value = "Diff2";
            rng14.Value = "Perc2";


            int tooArr = dt2.Rows.Count + 2;
            object[,] arr = new object[dt2.Rows.Count, dt2.Columns.Count];
            for (int r = 0; r < dt2.Rows.Count; r++)
            {
                DataRow dr = dt2.Rows[r];
                for (int c = 0; c < dt2.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }
            Microsoft.Office.Interop.Excel.Range rng2 = excelWorksheet.Range["A3:E" + tooArr];
            Microsoft.Office.Interop.Excel.Range rng15 = excelWorksheet.Range["H3:I" + tooArr];

            object[,] arr4 = new object[dt.Rows.Count, dt.Columns.Count];
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                DataRow dr = dt.Rows[r];
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    arr4[r, c] = dr[c];
                }
            }


            int tooArr2 = tooArr + 2;
            rng2.Value = arr;
            rng15.Value = arr4;
            //SUM VALUES

            Microsoft.Office.Interop.Excel.Range rng9 = excelWorksheet.Range["F3:F" + tooArr];
            Microsoft.Office.Interop.Excel.Range rng3 = excelWorksheet.Range["A1:K" + tooArr];

            object[,] arr2 = new object[tooArr, 1];

            int itr = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr2[r, 0] = "=SUM(E" + itr + "-D" + itr + ")";

                itr++;

            }

            object[,] arr5 = new object[tooArr, 1];

            int itr3 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr5[r, 0] = "=SUM(I" + itr3 + "-H" + itr3 + ")";

                itr3++;

            }
            Microsoft.Office.Interop.Excel.Range rng10 = excelWorksheet.Range["G3:G" + tooArr];
            Microsoft.Office.Interop.Excel.Range rng16 = excelWorksheet.Range["J3:J" + tooArr];
            rng16.Value = arr5;

            object[,] arr3 = new object[tooArr, 1];

            int itr2 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr3[r, 0] = "=IFERROR(IF(D" + itr2 + "=0,F" + itr2 + "/E" + itr2 + ",F" + itr2 + "/D" + itr2 + "), \"No Sales\")";

                //arr3[r, 0] = "foo";

                itr2++;

            }

            Microsoft.Office.Interop.Excel.Range rng17 = excelWorksheet.Range["K3:K" + tooArr];
            object[,] arr6 = new object[tooArr, 1];

            int itr4 = 3;
            for (int r = 0; r < tooArr; r++)
            {
                arr6[r, 0] = "=IFERROR(IF(H" + itr4 + "=0,J" + itr4 + "/I" + itr4 + ",J" + itr4 + "/H" + itr4 + "), \"No Sales\")";

                //arr3[r, 0] = "foo";

                itr4++;

            }
            //= IFERROR(IF(D3 = 0, F3 / E3, F3 / D3), 'No Sales')
            //= IFERROR(IF(D2 = 0, F2 / E2, F2 / D2), "No Sales")

            rng9.Value = arr2;
            rng10.NumberFormat = "###,##%";
            rng10.Value = arr3;
            rng17.Value = arr6;
            rng17.NumberFormat = "###,##%";
            rng15.NumberFormat = "$#,##0.00";
            rng16.NumberFormat = "$#,##0.00";
            int[] fields = new int[] { 4, 5, 6, 8, 9, 10 };

            rng3.Subtotal(1, Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum,
                fields, Microsoft.Office.Interop.Excel.XlSummaryRow.xlSummaryBelow);

            excelWorksheet.Columns.AutoFit();
        }
    }
}
