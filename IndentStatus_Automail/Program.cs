using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace IndentStatus_Automail
{
    /// <summary>
    /// Lightweight replacement for DataTable: holds rows as object[] plus column
    /// metadata. Kept API-compatible (RowCount/ColumnCount/Rows/ColumnNames) so the
    /// rest of the code reads almost identically to the original DataTable version.
    /// </summary>
    public class QueryResult
    {
        public List<object[]> Rows = new List<object[]>();
        public string[] ColumnNames;
        public Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        public int RowCount => Rows.Count;
        public int ColumnCount => ColumnNames?.Length ?? 0;
    }

    static class Program
    {
        private const string ConnectionString = "Data Source=192.168.4.150;Initial Catalog=SAP_ELM;User ID=sa;Password=Elnetsrv123*";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                string now = DateTime.Now.ToString("yyyyMMdd");
                string Hour = DateTime.Now.ToString("HH:mm");
                string Hr = Hour.Replace(":", "");
                AA();

                string Attachment = System.Windows.Forms.Application.StartupPath + "\\Pending_Orders_" + now + "_" + Hr + ".xlsx";

                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                string ToMailsList = "";
                string CcMailList = "";
                string SMTPHost = "smtp.gmail.com";
                string FromMail = "order@elmeasure.com";
                string Password = "cgju gymh cppd udgd";
                string MailSubject = "NEW Indent Status Report as On :" + DateTime.Now.ToString("dd/MM/yyyy") + "";
                message.From = new MailAddress(FromMail);

                using (SqlConnection con11 = new SqlConnection(ConnectionString))
                {
                    con11.Open();
                    ApplyDeadlockSafeSettings(con11);

                    // Kept for parity with original behavior (result currently unused,
                    // same as before — ToMailsList/CcMailList assignment from it stayed commented out).
                    using (SqlCommand cmd11 = new SqlCommand(
                        "SELECT T0.[U_Email] , T0.[U_CcMail] FROM [dbo].[@AUTO_EMAIL_RPT]  T0 WHERE cast(T0.[U_rptname]as varchar (max)) ='IndentStatusAsOnDate_AutoMail'",con11))
                    {
                        ExecuteWithRetry(cmd11, new QueryResult());
                    }

                    //Live
                    ToMailsList = "sam@elmeasure.com,s.jebasingh@elmeasure.com,babu@elmeasure.com";
                    CcMailList = "lakshmikanth@elmeasure.in,srilakshmi@elmeasure.in,bharani.ns@elmeasure.com";

                    //Test
                    //ToMailsList = "srilakshmi@elmeasure.in";
                    //CcMailList = "srilakshmi@elmeasure.in";

                    string TotalMailsList = ToMailsList;
                    string MailsList = TotalMailsList.TrimEnd(',');
                    string[] Multi = MailsList.Split(',');
                    foreach (string MultiMailID in Multi)
                    {
                        message.To.Add(new MailAddress(MultiMailID));
                    }
                    string CCEmailList = CcMailList;
                    if (CcMailList != "")
                    {
                        string CcList = CCEmailList.TrimEnd(',').TrimStart(',');
                        string[] Multi1 = CcList.Split(',');
                        foreach (string MultiMailID1 in Multi1)
                        {
                            string p = MultiMailID1.Replace("\"", "");
                            message.CC.Add(new MailAddress(p));
                        }
                    }

                    string textBody = " <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 400 + "><tr bgcolor='#4da6ff'><td><b>Column 1</b></td> <td> <b> Column 2</b> </td></tr>";
                    textBody += "</table>";
                    message.Subject = MailSubject;

                    Assembly assembly = Assembly.GetExecutingAssembly();
                    StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\" + "HTML_CLK_V2.htm");
                    string readFile = reader.ReadToEnd();
                    string myString = readFile;
                    CultureInfo hindi = new CultureInfo("hi-IN");

                    using (SqlCommand cmdProc = new SqlCommand("IndentStatusAsOnDate_AutoMail_CLK_V3", con11))
                    {
                        cmdProc.CommandType = CommandType.StoredProcedure;
                        cmdProc.CommandTimeout = 36000;
                        QueryResult dtELM = new QueryResult();
                        ExecuteWithRetry(cmdProc, dtELM);

                        string A1 = FormatCurrency(dtELM, 0, "#A1", hindi);
                        string A2 = FormatCurrency(dtELM, 0, "#A2", hindi);
                        string A3 = FormatCurrency(dtELM, 0, "#A3", hindi);
                        string A11 = FormatCurrency(dtELM, 0, "#A11", hindi);
                        string YTD = FormatCurrency(dtELM, 0, "#YTD", hindi);

                        string G1 = FormatCurrency(dtELM, 0, "#G1", hindi);
                        string G2 = FormatCurrency(dtELM, 0, "#G2", hindi);
                        string H1 = FormatCurrency(dtELM, 0, "#H1", hindi);
                        string H2 = FormatCurrency(dtELM, 0, "#H2", hindi);
                        string I1 = FormatCurrency(dtELM, 0, "#I1", hindi);
                        string I2 = FormatCurrency(dtELM, 0, "#I2", hindi);
                        string J1 = FormatCurrency(dtELM, 0, "#J1", hindi);
                        string J2 = FormatCurrency(dtELM, 0, "#J2", hindi);
                        string GFDP = FormatCurrency(dtELM, 0, "#GFDP", hindi);
                        string GATS = FormatCurrency(dtELM, 0, "#GATS", hindi);
                        string HFDP = FormatCurrency(dtELM, 0, "#HFDP", hindi);
                        string HATS = FormatCurrency(dtELM, 0, "#HATS", hindi);
                        string IFDP = FormatCurrency(dtELM, 0, "#IFDP", hindi);
                        string IATS = FormatCurrency(dtELM, 0, "#IATS", hindi);
                        string JFDP = FormatCurrency(dtELM, 0, "#JFDP", hindi);
                        string JATS = FormatCurrency(dtELM, 0, "#JATS", hindi);
                        string ORCSW = FormatCurrency(dtELM, 0, "#ORCSW", hindi);
                        string ORCFDP = FormatCurrency(dtELM, 0, "#ORCFDP", hindi);
                        string ORCATS = FormatCurrency(dtELM, 0, "#ORCATS", hindi);
                        string ORCOEM = FormatCurrency(dtELM, 0, "#ORCOEM", hindi);
                        string ORCPM = FormatCurrency(dtELM, 0, "#ORCPM", hindi);
                        string ORCTOTAL = FormatCurrency(dtELM, 0, "#ORCTOTAL", hindi);

                        string ELYB = FormatCurrency(dtELM, 0, "#ELYB", hindi);
                        string CBEYB = FormatCurrency(dtELM, 0, "#CBEYB", hindi);
                        string AHDYB = FormatCurrency(dtELM, 0, "#AHDYB", hindi);
                        string SWYB = FormatCurrency(dtELM, 0, "#SWYB", hindi);
                        string ADGYB = FormatCurrency(dtELM, 0, "#ADGYB", hindi);
                        string TotYB = FormatCurrency(dtELM, 0, "#TotYB", hindi);
                        string GTotYB = FormatCurrency(dtELM, 0, "#GTotYB", hindi);

                        string ELB = FormatCurrency(dtELM, 0, "#ELB", hindi);
                        string CBEB = FormatCurrency(dtELM, 0, "#CBEB", hindi);
                        string AHDB = FormatCurrency(dtELM, 0, "#AHDB", hindi);
                        string SWB = FormatCurrency(dtELM, 0, "#SWB", hindi);
                        string ADGB = FormatCurrency(dtELM, 0, "#ADGB", hindi);
                        string TotB = FormatCurrency(dtELM, 0, "#TotB", hindi);
                        string GTotB = FormatCurrency(dtELM, 0, "#GTotB", hindi);

                        string ELOH = FormatCurrency(dtELM, 0, "#ELOH", hindi);
                        string SWOH = FormatCurrency(dtELM, 0, "#SWOH", hindi);
                        string CBEOH = FormatCurrency(dtELM, 0, "#CBEOH", hindi);
                        string AHDOH = FormatCurrency(dtELM, 0, "#AHDOH", hindi);
                        string ADGOH = FormatCurrency(dtELM, 0, "#ADGOH", hindi);
                        string TotOH = FormatCurrency(dtELM, 0, "#TotOH", hindi);
                        string GTotOH = FormatCurrency(dtELM, 0, "#GTotOH", hindi);

                        string ELFG = FormatCurrency(dtELM, 0, "#ELFG", hindi);
                        string SWFG = FormatCurrency(dtELM, 0, "#SWFG", hindi);
                        string CBEFG = FormatCurrency(dtELM, 0, "#CBEFG", hindi);
                        string AHDFG = FormatCurrency(dtELM, 0, "#AHDFG", hindi);
                        string ADGFG = FormatCurrency(dtELM, 0, "#ADGFG", hindi);
                        string TotFG = FormatCurrency(dtELM, 0, "#TotFG", hindi);
                        string GTotFG = FormatCurrency(dtELM, 0, "#GTotFG", hindi);

                        string ELHBA = FormatCurrency(dtELM, 0, "#ELHBA", hindi);
                        string SWHBA = FormatCurrency(dtELM, 0, "#SWHBA", hindi);
                        string CBEHBA = FormatCurrency(dtELM, 0, "#CBEHBA", hindi);
                        string AHDHBA = FormatCurrency(dtELM, 0, "#AHDHBA", hindi);
                        string ADGHBA = FormatCurrency(dtELM, 0, "#ADGHBA", hindi);
                        string TotHBA = FormatCurrency(dtELM, 0, "#TotHBA", hindi);
                        string GTotHBA = FormatCurrency(dtELM, 0, "#GTotHBA", hindi);

                        string ELHBS = FormatCurrency(dtELM, 0, "#ELHBS", hindi);
                        string SWHBS = FormatCurrency(dtELM, 0, "#SWHBS", hindi);
                        string CBEHBS = FormatCurrency(dtELM, 0, "#CBEHBS", hindi);
                        string AHDHBS = FormatCurrency(dtELM, 0, "#AHDHBS", hindi);
                        string ADGHBS = FormatCurrency(dtELM, 0, "#ADGHBS", hindi);
                        string TotHBS = FormatCurrency(dtELM, 0, "#TotHBS", hindi);
                        string GTotHBS = FormatCurrency(dtELM, 0, "#GTotHBS", hindi);

                        string ELHBR = FormatCurrency(dtELM, 0, "#ELHBR", hindi);
                        string SWHBR = FormatCurrency(dtELM, 0, "#SWHBR", hindi);
                        string CBEHBR = FormatCurrency(dtELM, 0, "#CBEHBR", hindi);
                        string AHDHBR = FormatCurrency(dtELM, 0, "#AHDHBR", hindi);
                        string ADGHBR = FormatCurrency(dtELM, 0, "#ADGHBR", hindi);
                        string TotHBR = FormatCurrency(dtELM, 0, "#TotHBR", hindi);
                        string GTotHBR = FormatCurrency(dtELM, 0, "#GTotHBR", hindi);
                        string FDPHBR = FormatCurrency(dtELM, 0, "#FDPHBR", hindi);
                        string ATSHBR = FormatCurrency(dtELM, 0, "#ATSHBR", hindi);
                        string OEMHBR = FormatCurrency(dtELM, 0, "#OEMHBR", hindi);
                        string FDPFG = FormatCurrency(dtELM, 0, "#FDPFG", hindi);
                        string ATSFG = FormatCurrency(dtELM, 0, "#ATSFG", hindi);
                        string OEMFG = FormatCurrency(dtELM, 0, "#OEMFG", hindi);
                        string FDPHBA = FormatCurrency(dtELM, 0, "#FDPHBA", hindi);
                        string ATSHBA = FormatCurrency(dtELM, 0, "#ATSHBA", hindi);
                        string OEMHBA = FormatCurrency(dtELM, 0, "#OEMHBA", hindi);
                        string FDPHBS = FormatCurrency(dtELM, 0, "#FDPHBS", hindi);
                        string ATSHBS = FormatCurrency(dtELM, 0, "#ATSHBS", hindi);
                        string OEMHBS = FormatCurrency(dtELM, 0, "#OEMHBS", hindi);
                        string FDPYB = FormatCurrency(dtELM, 0, "#FDPYB", hindi);
                        string ATSYB = FormatCurrency(dtELM, 0, "#ATSYB", hindi);
                        string OEMYB = FormatCurrency(dtELM, 0, "#OEMYB", hindi);
                        string FDPB = FormatCurrency(dtELM, 0, "#FDPB", hindi);
                        string ATSB = FormatCurrency(dtELM, 0, "#ATSB", hindi);
                        string OEMB = FormatCurrency(dtELM, 0, "#OEMB", hindi);
                        string FDPOH = FormatCurrency(dtELM, 0, "#FDPOH", hindi);
                        string ATSOH = FormatCurrency(dtELM, 0, "#ATSOH", hindi);
                        string OEMOH = FormatCurrency(dtELM, 0, "#OEMOH", hindi);
                        string FDPTBOG = FormatCurrency(dtELM, 0, "#FDPTBOG", hindi);
                        string ATSTBOG = FormatCurrency(dtELM, 0, "#ATSTBOG", hindi);
                        string OEMTBOG = FormatCurrency(dtELM, 0, "#OEMTBOG", hindi);
                        string FDPYTOB = FormatCurrency(dtELM, 0, "#FDPYTOB", hindi);
                        string ATSYTOB = FormatCurrency(dtELM, 0, "#ATSYTOB", hindi);
                        string OEMYTOB = FormatCurrency(dtELM, 0, "#OEMYTOB", hindi);
                        //Customer Required this Month
                        string ELCRM = FormatCurrency(dtELM, 0, "#ELCRM", hindi);
                        string CBECRM = FormatCurrency(dtELM, 0, "#CBECRM", hindi);
                        string AHDCRM = FormatCurrency(dtELM, 0, "#AHDCRM", hindi);
                        string SWCRM = FormatCurrency(dtELM, 0, "#SWCRM", hindi);
                        string FDPCRM = FormatCurrency(dtELM, 0, "#FDPCRM", hindi);
                        string ATSCRM = FormatCurrency(dtELM, 0, "#ATSCRM", hindi);
                        string ADGCRM = FormatCurrency(dtELM, 0, "#ADGCRM", hindi);
                        string OEMCRM = FormatCurrency(dtELM, 0, "#OEMCRM", hindi);
                        string TotCRM = FormatCurrency(dtELM, 0, "#TotCRM", hindi);
                        string GTotCRM = FormatCurrency(dtELM, 0, "#GTotCRM", hindi);

                        //Customer Required future Month
                        string ELCRF = FormatCurrency(dtELM, 0, "#ELCRF", hindi);
                        string CBECRF = FormatCurrency(dtELM, 0, "#CBECRF", hindi);
                        string AHDCRF = FormatCurrency(dtELM, 0, "#AHDCRF", hindi);
                        string SWCRF = FormatCurrency(dtELM, 0, "#SWCRF", hindi);
                        string FDPCRF = FormatCurrency(dtELM, 0, "#FDPCRF", hindi);
                        string ATSCRF = FormatCurrency(dtELM, 0, "#ATSCRF", hindi);
                        string ADGCRF = FormatCurrency(dtELM, 0, "#ADGCRF", hindi);
                        string OEMCRF = FormatCurrency(dtELM, 0, "#OEMCRF", hindi);
                        string TotCRF = FormatCurrency(dtELM, 0, "#TotCRF", hindi);
                        string GTotCRF = FormatCurrency(dtELM, 0, "#GTotCRF", hindi);

                        string ELTBOG = FormatCurrency(dtELM, 0, "#ELTBOG", hindi);
                        string SWTBOG = FormatCurrency(dtELM, 0, "#SWTBOG", hindi);
                        string CBTBOG = FormatCurrency(dtELM, 0, "#CBTBOG", hindi);
                        string AHTBOG = FormatCurrency(dtELM, 0, "#AHTBOG", hindi);
                        string ADGTBOG = FormatCurrency(dtELM, 0, "#ADGTBOG", hindi);
                        string TotTBOG = FormatCurrency(dtELM, 0, "#TotTBOG", hindi);
                        string GTtBOG = FormatCurrency(dtELM, 0, "#GTtBOG", hindi);

                        string ELYTOB = FormatCurrency(dtELM, 0, "#ELYTOB", hindi);
                        string SWYTOB = FormatCurrency(dtELM, 0, "#SWYTOB", hindi);
                        string CBYTOB = FormatCurrency(dtELM, 0, "#CBYTOB", hindi);
                        string AHYTOB = FormatCurrency(dtELM, 0, "#AHYTOB", hindi);
                        string ADGYTOB = FormatCurrency(dtELM, 0, "#ADGYTOB", hindi);
                        string TotYTOB = FormatCurrency(dtELM, 0, "#TotYTOB", hindi);
                        string GTotTOB = FormatCurrency(dtELM, 0, "#GTotTOB", hindi);

                        string GOEM = FormatCurrency(dtELM, 0, "#GOEM", hindi);
                        string HOEM = FormatCurrency(dtELM, 0, "#HOEM", hindi);
                        string IOEM = FormatCurrency(dtELM, 0, "#IOEM", hindi);
                        string JOEM = FormatCurrency(dtELM, 0, "#JOEM", hindi);



                        myString = myString.Replace("#DATE", DateTime.Now.ToString("dd/MM/yyyy"));
                        myString = myString.Replace("#A1", A1);
                        myString = myString.Replace("#A2", A2);
                        myString = myString.Replace("#A3", A3);
                        myString = myString.Replace("#ASC", A11);
                        myString = myString.Replace("#YTD", YTD);

                        myString = myString.Replace("#G1", G1);
                        myString = myString.Replace("#G2", G2);
                        myString = myString.Replace("#H1", H1);
                        myString = myString.Replace("#H2", H2);
                        myString = myString.Replace("#I1", I1);
                        myString = myString.Replace("#I2", I2);
                        myString = myString.Replace("#J1", J1);
                        myString = myString.Replace("#J2", J2);
                        myString = myString.Replace("#GFDP", GFDP);
                        myString = myString.Replace("#GATS", GATS);
                        myString = myString.Replace("#HFDP", HFDP);
                        myString = myString.Replace("#HATS", HATS);
                        myString = myString.Replace("#IFDP", IFDP);
                        myString = myString.Replace("#IATS", IATS);
                        myString = myString.Replace("#JFDP", JFDP);
                        myString = myString.Replace("#JATS", JATS);
                        myString = myString.Replace("#ORCSW", ORCSW);

                        myString = myString.Replace("#GOEM", GOEM);
                        myString = myString.Replace("#HOEM", HOEM);
                        myString = myString.Replace("#IOEM", IOEM);
                        myString = myString.Replace("#JOEM", JOEM);


                        myString = myString.Replace("#ORCFDP", ORCFDP);
                        myString = myString.Replace("#ORCATS", ORCATS);

                        myString = myString.Replace("#ORCOEM", ORCOEM);

                        myString = myString.Replace("#ORCPM", ORCPM);

                        myString = myString.Replace("#ORCTOTAL", ORCTOTAL);

                        myString = myString.Replace("#ELYB", ELYB);
                        myString = myString.Replace("#CBEYB", CBEYB);
                        myString = myString.Replace("#AHDYB", AHDYB);
                        myString = myString.Replace("#SWYB", SWYB);
                        myString = myString.Replace("#ADGYB", ADGYB);
                        myString = myString.Replace("#TotYB", TotYB);
                        myString = myString.Replace("#GTotYB", GTotYB);

                        myString = myString.Replace("#ELB", ELB);
                        myString = myString.Replace("#CBEB", CBEB);
                        myString = myString.Replace("#AHDB", AHDB);
                        myString = myString.Replace("#SWB", SWB);
                        myString = myString.Replace("#ADGB", ADGB);
                        myString = myString.Replace("#TotB", TotB);
                        myString = myString.Replace("#GTotB", GTotB);

                        myString = myString.Replace("#ELOH", ELOH);
                        myString = myString.Replace("#CBEOH", CBEOH);
                        myString = myString.Replace("#AHDOH", AHDOH);
                        myString = myString.Replace("#SWOH", SWOH);
                        myString = myString.Replace("#ADGOH", ADGOH);
                        myString = myString.Replace("#TotOH", TotOH);
                        myString = myString.Replace("#GTotFG", GTotFG);

                        myString = myString.Replace("#ELFG", ELFG);
                        myString = myString.Replace("#CBEFG", CBEFG);
                        myString = myString.Replace("#AHDFG", AHDFG);
                        myString = myString.Replace("#SWFG", SWFG);
                        myString = myString.Replace("#ADGFG", ADGFG);
                        myString = myString.Replace("#TotFG", TotFG);
                        myString = myString.Replace("#GTotOH", GTotOH);

                        myString = myString.Replace("#ELHBA", ELHBA);
                        myString = myString.Replace("#CBEHBA", CBEHBA);
                        myString = myString.Replace("#AHDHBA", AHDHBA);
                        myString = myString.Replace("#SWHBA", SWHBA);
                        myString = myString.Replace("#ADGHBA", ADGHBA);
                        myString = myString.Replace("#TotHBA", TotHBA);
                        myString = myString.Replace("#GTotHBA", GTotHBA);

                        myString = myString.Replace("#ELHBS", ELHBS);
                        myString = myString.Replace("#CBEHBS", CBEHBS);
                        myString = myString.Replace("#AHDHBS", AHDHBS);
                        myString = myString.Replace("#SWHBS", SWHBS);
                        myString = myString.Replace("#ADGHBS", ADGHBS);
                        myString = myString.Replace("#TotHBS", TotHBS);
                        myString = myString.Replace("#GTotHBS", GTotHBS);

                        myString = myString.Replace("#ELHBR", ELHBR);
                        myString = myString.Replace("#CBEHBR", CBEHBR);
                        myString = myString.Replace("#AHDHBR", AHDHBR);
                        myString = myString.Replace("#SWHBR", SWHBR);
                        myString = myString.Replace("#ADGHBR", ADGHBR);
                        myString = myString.Replace("#TotHBR", TotHBR);
                        myString = myString.Replace("#GTotHBR", GTotHBR);
                        myString = myString.Replace("#FDPHBR", FDPHBR);
                        myString = myString.Replace("#ATSHBR", ATSHBR);
                        myString = myString.Replace("#OEMHBR", OEMHBR);
                        myString = myString.Replace("#FDPFG", FDPFG);
                        myString = myString.Replace("#ATSFG", ATSFG);
                        myString = myString.Replace("#OEMFG", OEMFG);
                        myString = myString.Replace("#FDPHBA", FDPHBA);
                        myString = myString.Replace("#ATSHBA", ATSHBA);
                        myString = myString.Replace("#OEMHBA", OEMHBA);
                        myString = myString.Replace("#FDPHBS", FDPHBS);
                        myString = myString.Replace("#ATSHBS", ATSHBS);
                        myString = myString.Replace("#OEMHBS", OEMHBS);
                        myString = myString.Replace("#FDPYB", FDPYB);
                        myString = myString.Replace("#ATSYB", ATSYB);
                        myString = myString.Replace("#OEMYB", OEMYB);
                        myString = myString.Replace("#FDPB", FDPB);
                        myString = myString.Replace("#ATSB", ATSB);
                        myString = myString.Replace("#OEMB", OEMB);
                        myString = myString.Replace("#FDPOH", FDPOH);
                        myString = myString.Replace("#ATSOH", ATSOH);
                        myString = myString.Replace("#OEMOH", OEMOH);
                        myString = myString.Replace("#FDPTBOG", FDPTBOG);
                        myString = myString.Replace("#ATSTBOG", ATSTBOG);
                        myString = myString.Replace("#OEMTBOG", OEMTBOG);
                        myString = myString.Replace("#FDPYTOB", FDPYTOB);
                        myString = myString.Replace("#ATSYTOB", ATSYTOB);
                        myString = myString.Replace("#OEMYTOB", OEMYTOB);

                        //Customer Required this Month
                        myString = myString.Replace("#ELCRM", ELCRM);
                        myString = myString.Replace("#CBECRM", CBECRM);
                        myString = myString.Replace("#AHDCRM", AHDCRM);
                        myString = myString.Replace("#SWCRM", SWCRM);
                        myString = myString.Replace("#FDPCRM", FDPCRM);
                        myString = myString.Replace("#ATSCRM", ATSCRM);
                        myString = myString.Replace("#ADGCRM", ADGCRM);
                        myString = myString.Replace("#OEMCRM", OEMCRM);
                        myString = myString.Replace("#TotCRM", TotCRM);
                        myString = myString.Replace("#GTotCRM", GTotCRM);

                        //Customer Required future Month
                        myString = myString.Replace("#ELCRF", ELCRF);
                        myString = myString.Replace("#CBECRF", CBECRF);
                        myString = myString.Replace("#AHDCRF", AHDCRF);
                        myString = myString.Replace("#SWCRF", SWCRF);
                        myString = myString.Replace("#FDPCRF", FDPCRF);
                        myString = myString.Replace("#ATSCRF", ATSCRF);
                        myString = myString.Replace("#ADGCRF", ADGCRF);
                        myString = myString.Replace("#OEMCRF", OEMCRF);
                        myString = myString.Replace("#TotCRF", TotCRF);
                        myString = myString.Replace("#GTotCRF", GTotCRF);

                        myString = myString.Replace("#ELTBOG", ELTBOG);
                        myString = myString.Replace("#CBTBOG", CBTBOG);
                        myString = myString.Replace("#AHTBOG", AHTBOG);
                        myString = myString.Replace("#SWTBOG", SWTBOG);
                        myString = myString.Replace("#ADGTBOG", ADGTBOG);
                        myString = myString.Replace("#TotTBOG", TotTBOG);
                        myString = myString.Replace("#GTtBOG", GTtBOG);

                        myString = myString.Replace("#ELYTOB", ELYTOB);
                        myString = myString.Replace("#CBYTOB", CBYTOB);
                        myString = myString.Replace("#AHYTOB", AHYTOB);
                        myString = myString.Replace("#SWYTOB", SWYTOB);
                        myString = myString.Replace("#ADGYTOB", ADGYTOB);
                        myString = myString.Replace("#TotYTOB", TotYTOB);
                        myString = myString.Replace("#GTotTOB", GTotTOB);
                    }

                    message.Body = myString;

                    System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(Attachment);
                    message.Attachments.Add(attachment);

                    message.IsBodyHtml = true;
                    message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.Host = SMTPHost;
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    smtp.Timeout = 100000;
                    smtp.Credentials = new NetworkCredential(FromMail, Password);
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object s,
                        System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };
                    message.Priority = MailPriority.High;
                    smtp.Send(message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException != null ? ex.InnerException.ToString() : ex.ToString());
            }
            System.Windows.Forms.Application.Exit();
        }

        /// <summary>
        /// Runs a command with a simple deadlock-victim retry (SQL error 1205),
        /// populating a QueryResult via SqlDataReader instead of DataTable.Fill().
        /// </summary>
        private static void ExecuteWithRetry(SqlCommand cmd, QueryResult result, int maxRetries = 3)
        {
            int attempt = 0;
            while (true)
            {
                try
                {
                    result.Rows.Clear();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        result.ColumnNames = new string[dr.FieldCount];
                        result.ColumnIndex.Clear();
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            result.ColumnNames[i] = dr.GetName(i);
                            result.ColumnIndex[result.ColumnNames[i]] = i;
                        }
                        while (dr.Read())
                        {
                            object[] row = new object[dr.FieldCount];
                            dr.GetValues(row);
                            result.Rows.Add(row);
                        }
                    }
                    return;
                }
                catch (SqlException ex) when ((ex.Number == 1205 || ex.Number == 601) && attempt < maxRetries)
                {
                    // 1205 = deadlock victim.
                    // 601  = "Could not continue scan with NOLOCK due to data movement" —
                    //        a known transient side effect of READ UNCOMMITTED when a page
                    //        the scan is reading gets physically moved (index rebuild, page
                    //        split, concurrent update) mid-scan. Retrying almost always
                    //        succeeds once the movement settles.
                    attempt++;
                    System.Threading.Thread.Sleep(2000 * attempt);
                }
            }
        }

        /// <summary>
        /// Read-uncommitted + low deadlock priority so this reporting job never blocks,
        /// and is always the victim rather than a live SAP B1 transaction, if a deadlock occurs.
        /// </summary>
        private static void ApplyDeadlockSafeSettings(SqlConnection con)
        {
            using (SqlCommand isoCmd = new SqlCommand(
                "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET DEADLOCK_PRIORITY LOW;", con))
            {
                isoCmd.ExecuteNonQuery();
            }
        }

        //Old - With decimal values
        //private static string FormatCurrency(QueryResult qr, int rowIdx, string col, CultureInfo culture)
        //{
        //    string val = qr.Rows[rowIdx][qr.ColumnIndex[col]]?.ToString().Trim();
        //    if (string.IsNullOrEmpty(val)) val = "0";
        //    return string.Format(culture, "{0:c}", decimal.Parse(val, CultureInfo.InvariantCulture)).Replace("₹", "");
        //}

        //Rounded-off values
        private static string FormatCurrency(QueryResult qr, int rowIdx, string col, CultureInfo culture)
        {
            string val = qr.Rows[rowIdx][qr.ColumnIndex[col]]?.ToString().Trim();
            if (string.IsNullOrEmpty(val)) val = "0";

            decimal amount = decimal.Parse(val, CultureInfo.InvariantCulture);
            decimal rounded = Math.Round(amount, 0, MidpointRounding.AwayFromZero);

            return string.Format(culture, "{0:C0}", rounded).Replace("₹", "");
        }

        private static QueryResult RunProc(SqlConnection con, string procName, int timeout = 3600)
        {
            QueryResult result = new QueryResult();
            using (SqlCommand cmd = new SqlCommand(procName, con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = timeout;
                ExecuteWithRetry(cmd, result);
            }
            return result;
        }

        /// <summary>
        /// Writes headers + all row values in a single COM call via Range.Value2
        /// instead of looping Cells[i,j] = value one cell at a time. Same resulting
        /// values as the original cell-by-cell writes, much faster for large sheets.
        /// </summary>
        private static void WriteSheetData(Worksheet sheet, QueryResult qr)
        {
            for (int j = 0; j < qr.ColumnCount; j++)
                sheet.Cells[1, j + 1] = qr.ColumnNames[j];

            if (qr.RowCount == 0) return;

            object[,] arr = new object[qr.RowCount, qr.ColumnCount];
            for (int i = 0; i < qr.RowCount; i++)
                for (int j = 0; j < qr.ColumnCount; j++)
                    arr[i, j] = qr.Rows[i][j] ?? "";

            Range startCell = (Range)sheet.Cells[2, 1];
            Range endCell = (Range)sheet.Cells[qr.RowCount + 1, qr.ColumnCount];
            Range writeRange = sheet.Range[startCell, endCell];
            writeRange.Value2 = arr;
        }
        public static void AA()
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel.Workbook oWB = null;
            try
            {
                string now = DateTime.Now.ToString("yyyyMMdd");
                string Hour = DateTime.Now.ToString("HH:mm");
                string Hr = Hour.Replace(":", "");

                object missing = Type.Missing;

                // Applies a real date format to known date columns (looked up by NAME,
                // not a hardcoded column number) and forces every other column to
                // General. This replaces the old hardcoded-column-index approach,
                // which silently goes stale whenever a procedure's column layout
                // changes and was the cause of dates/quantities showing in the
                // wrong columns.


                using (SqlConnection con = new SqlConnection(ConnectionString))
                {
                    con.Open();
                    ApplyDeadlockSafeSettings(con);

                    QueryResult dt0 = RunProc(con, "ARRInventoryTransferPending");
                    QueryResult dt = RunProc(con, "Indent_Delivery_Performance");
                    QueryResult dt1 = RunProc(con, "Indent_Accessories_CL");
                    QueryResult dt2 = RunProc(con, "Indent_NEW_ADIGA");
                    QueryResult dt3 = RunProc(con, "Indent_NEW_2023_CL");
                    QueryResult dt4 = RunProc(con, "Today_Billable_OAs_CL");
                    QueryResult dt5 = RunProc(con, "Indent_NEW_2023_SG_CL");
                    QueryResult dt6 = RunProc(con, "Demo_Request_Pending");

                    if (dt0 == null || dt0.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    // Suppress "overwrite?"/format-warning dialogs, which otherwise
                    // block headlessly-running automation and surface as
                    // "Exception from HRESULT: 0x800AC472".
                    oXL.DisplayAlerts = false;
                    oXL.Visible = false;
                    oWB = oXL.Workbooks.Add(missing);

                    // ---- New first sheet: ARRInventoryTransferPending ----
                    Microsoft.Office.Interop.Excel.Worksheet oSheet0 = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    Microsoft.Office.Interop.Excel.Range excelCellrange;

                    oSheet0.Name = "ARRInventoryTransferPending";
                    WriteSheetData(oSheet0, dt0);

                    // ---- Everything below is unchanged except that "24HoursDeliveryPerformance"
                    // is now added as a sheet (instead of being the workbook's initial sheet),
                    // so it lands right after the new first sheet. ----
                    Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt == null || dt.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet.Name = "24HoursDeliveryPerformance";
                    WriteSheetData(oSheet, dt);

                    for (var i = 0; i < dt.RowCount; i++)
                    {
                        for (var j = 0; j < dt.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt.Rows[i][j - 2].ToString());
                                    string Unblock = dt.Rows[i][j + 1].ToString();
                                    string PrefDel = dt.Rows[i][j + 2].ToString();

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "OA Status")
                                {
                                    double BillableQty = Convert.ToDouble(dt.Rows[i][j - 1].ToString());
                                    double FGQty = Convert.ToDouble(dt.Rows[i][j - 2].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt.Rows[i][j - 3].ToString());
                                    string Unblock = dt.Rows[i][j].ToString();
                                    string PrefDel = dt.Rows[i][j + 1].ToString();
                                    int NoODays = 0;

                                    if (dt.Rows[i][j + 3].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt.Rows[i][j + 3]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet1 (dt), row={i}, col={j}, ColName={dt.ColumnNames[j]}, ColumnCount={dt.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt1 == null || dt1.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet2.Name = "Accessories-ELM&CBE&AHD";
                    WriteSheetData(oSheet2, dt1);

                    for (var i = 0; i < dt1.RowCount; i++)
                    {
                        for (var j = 0; j < dt1.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt1.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt1.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt1.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt1.Rows[i][j - 5].ToString());
                                    string Unblock = dt1.Rows[i][j + 2].ToString();
                                    string PrefDel = dt1.Rows[i][j + 4].ToString();
                                    string ModlNm = dt1.Rows[i][j - 10].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt1.Rows[i][j - 2].ToString());
                                    double FGQty = Convert.ToDouble(dt1.Rows[i][j - 3].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt1.Rows[i][j - 7].ToString());
                                    string Unblock = dt1.Rows[i][j].ToString();
                                    string PrefDel = dt1.Rows[i][j + 2].ToString();
                                    string ModlNm = dt1.Rows[i][j - 12].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt1.Rows[i][j + 16].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt1.Rows[i][j + 16]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet2 (dt1), row={i}, col={j}, ColName={dt1.ColumnNames[j]}, ColumnCount={dt1.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet3 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt2 == null || dt2.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet3.Name = "PendingOrders-ADIGA";
                    WriteSheetData(oSheet3, dt2);

                    for (var i = 0; i < dt2.RowCount; i++)
                    {
                        for (var j = 0; j < dt2.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt2.ColumnNames[j].ToString();

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt2.Rows[i][j - 2].ToString());
                                    double FGQty = Convert.ToDouble(dt2.Rows[i][j - 3].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt2.Rows[i][j - 5].ToString());
                                    string Unblock = dt2.Rows[i][j].ToString();
                                    string PrefDel = dt2.Rows[i][j + 1].ToString();
                                    string ModlNm = dt2.Rows[i][j - 9].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt2.Rows[i][j + 12].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt2.Rows[i][j + 12]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt2.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt2.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt2.Rows[i][j - 3].ToString());
                                    string Unblock = dt2.Rows[i][j + 2].ToString();
                                    string PrefDel = dt2.Rows[i][j + 3].ToString();
                                    string ModlNm = dt2.Rows[i][j - 7].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet3 (dt2), row={i}, col={j}, ColName={dt2.ColumnNames[j]}, ColumnCount={dt2.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet4 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt3 == null || dt3.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet4.Name = "PendingOrders-ELM&CBE&AHD";
                    WriteSheetData(oSheet4, dt3);

                    for (var i = 0; i < dt3.RowCount; i++)
                    {
                        for (var j = 0; j < dt3.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt3.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt3.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt3.Rows[i][j - 5].ToString());
                                    string Unblock = dt3.Rows[i][j + 3].ToString();
                                    string PrefDel = dt3.Rows[i][j + 7].ToString();
                                    string ModlNm = dt3.Rows[i][j - 10].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);

                                        double BalanceQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                        double RemainigQty1 = Convert.ToDouble(dt3.Rows[i][dt3.ColumnIndex["Balance Qty to Be Billed"]].ToString());
                                        string BillableColor = dt3.Rows[i][dt3.ColumnIndex["Billable Colour"]].ToString();

                                        if (BillableColor == "Yellow" && PrefDel == "Partial Dispatch Allowed" && RemainigQty1 > 0)
                                        {
                                            excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                            FormattingExcelCells(excelCellrange, "#90EE90", System.Drawing.Color.Black, true);
                                        }
                                    }
                                }

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt3.Rows[i][j - 3].ToString());
                                    double FGQty = Convert.ToDouble(dt3.Rows[i][j - 4].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt3.Rows[i][j - 8].ToString());
                                    string Unblock = dt3.Rows[i][j].ToString();
                                    string PrefDel = dt3.Rows[i][j + 4].ToString();
                                    string ModlNm = dt3.Rows[i][j - 14].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt3.Rows[i][j + 22].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt3.Rows[i][j + 22]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "Bal Qty")
                                {
                                    double BalanceQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                    double RemainigQty1 = Convert.ToDouble(dt3.Rows[i][dt3.ColumnIndex["Balance Qty to Be Billed"]].ToString());
                                    string Unblock = dt3.Rows[i][j + 10].ToString();
                                    string PrefDel = dt3.Rows[i][j + 14].ToString();
                                    string BillableColor = dt3.Rows[i][dt3.ColumnIndex["Billable Colour"]].ToString();
                                    string ModlNm = dt3.Rows[i][j - 3].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (BillableColor == "Yellow" && PrefDel == "Partial Dispatch Allowed" && RemainigQty1 > 0)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#90EE90", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet4 (dt3), row={i}, col={j}, ColName={dt3.ColumnNames[j]}, ColumnCount={dt3.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet5 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt4 == null || dt4.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet5.Name = "TodayBillableOAs";
                    WriteSheetData(oSheet5, dt4);

                    for (var i = 0; i < dt4.RowCount; i++)
                    {
                        for (var j = 0; j < dt4.ColumnCount; j++)
                        {
                            string ColName = dt4.ColumnNames[j].ToString();

                            if (ColName == "Billable Qty")
                            {
                                double BillableQty = Convert.ToDouble(dt4.Rows[i][j].ToString());
                                string Color = dt4.Rows[i][j + 29].ToString();

                                if (Color == "Green")
                                {
                                    excelCellrange = oSheet5.Range[oSheet5.Cells[i + 2, j + 1], oSheet5.Cells[i + 2, j + 1]];
                                    FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                }
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet6 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt5 == null || dt5.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet6.Name = "SwitchGear";
                    WriteSheetData(oSheet6, dt5);

                    Microsoft.Office.Interop.Excel.Worksheet oSheet7 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt6 == null || dt6.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet7.Name = "Demo_Request_Pending";
                    WriteSheetData(oSheet7, dt6);

                    oXL.Visible = false;

                    // ---- Column auto-fit (now includes the new first sheet) ----
                    excelCellrange = oSheet0.Range[oSheet0.Cells[1, 1], oSheet0.Cells[1, dt0.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet2.Range[oSheet2.Cells[1, 1], oSheet2.Cells[1, dt1.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet3.Range[oSheet3.Cells[1, 1], oSheet3.Cells[1, dt2.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet4.Range[oSheet4.Cells[1, 1], oSheet4.Cells[1, dt3.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet5.Range[oSheet5.Cells[1, 1], oSheet5.Cells[1, dt4.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet6.Range[oSheet6.Cells[1, 1], oSheet6.Cells[1, dt5.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();

                    excelCellrange = oSheet0.Range[oSheet0.Cells[1, 1], oSheet0.Cells[1, dt0.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet2.Range[oSheet2.Cells[dt1.RowCount + 1, 1], oSheet2.Cells[dt1.RowCount + 1, dt1.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet3.Range[oSheet3.Cells[dt2.RowCount + 1, 1], oSheet3.Cells[dt2.RowCount + 1, dt2.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet4.Range[oSheet4.Cells[dt3.RowCount + 1, 1], oSheet4.Cells[dt3.RowCount + 1, dt3.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    oSheet0.Activate();

                    // ---- Date-column formatting, now driven by column NAME instead of a
                    // hardcoded column number, so it can't silently drift out of sync with
                    // the procedure's actual column layout again. Every non-date column is
                    // explicitly reset to General so a quantity column can never inherit a
                    // stray date format either. ----
                    string[] pendingOrdersDateColumns = {
                "OA Date", "Customer Delivery Date", "Delivery Date",
                "Initial Delivery Date", "Purchase date", "SO Date"
            };
                    ApplyColumnFormats(oSheet4, dt3, pendingOrdersDateColumns);
                    ApplyColumnFormats(oSheet5, dt4, pendingOrdersDateColumns);

                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;
                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;

                    excelCellrange = oSheet4.Range[oSheet4.Cells[1, 1], oSheet4.Cells[1, dt3.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet5.Range[oSheet5.Cells[1, 1], oSheet5.Cells[1, dt4.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet2.Range[oSheet2.Cells[1, 1], oSheet2.Cells[1, dt1.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet3.Range[oSheet3.Cells[1, 1], oSheet3.Cells[1, dt2.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    oSheet4.Columns["V"].ColumnWidth = 0;
                    oSheet5.Activate();
                    oSheet4.Columns["V"].ColumnWidth = 0;
                    oSheet4.Columns["AL"].ColumnWidth = 0;
                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;
                    oXL.ActiveWindow.FreezePanes = true;

                    oWB.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Pending_Orders_" + now + "_" + Hr + ".xlsx");
                }
            }
            catch (Exception ex)
            {
                // Original swallowed this silently. Logging now so failures (like
                // out-of-bounds column offsets) are visible instead of invisible.
                try
                {
                    string logPath = System.Windows.Forms.Application.StartupPath + "\\AA_Error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
                    File.WriteAllText(logPath, ex.ToString());
                }
                catch
                {
                    // If even logging fails, fall back to the original silent behavior.
                }
            }
            finally
            {
                // Always release Excel, whether AA() succeeded or threw. This is what
                // prevents orphaned EXCEL.EXE processes — which, left running from a
                // prior failed/aborted run, are a common cause of
                // "Exception from HRESULT: 0x800AC472" on the next run.
                if (oWB != null)
                {
                    try { oWB.Close(false); } catch { /* ignore */ }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                    oWB = null;
                }
                if (oXL != null)
                {
                    try { oXL.Quit(); } catch { /* ignore */ }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                    oXL = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        public static void ApplyColumnFormats(Microsoft.Office.Interop.Excel.Worksheet sheet, QueryResult data, string[] dateColumnNames)
        {
            if (sheet == null || data == null) return;
            var dateSet = new HashSet<string>(dateColumnNames, StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < data.ColumnCount; c++)
            {
                string colName = data.ColumnNames[c];
                var wholeCol = ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[1, c + 1]).EntireColumn;
                wholeCol.NumberFormat = dateSet.Contains(colName) ? "DD-MM-YYYY" : "General";
            }
        }
        public static void AA1()
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel.Workbook oWB = null;
            try
            {
                string now = DateTime.Now.ToString("yyyyMMdd");
                string Hour = DateTime.Now.ToString("HH:mm");
                string Hr = Hour.Replace(":", "");

                object missing = Type.Missing;

                using (SqlConnection con = new SqlConnection(ConnectionString))
                {
                    con.Open();
                    ApplyDeadlockSafeSettings(con);

                    QueryResult dt = RunProc(con, "Indent_Delivery_Performance");
                    QueryResult dt1 = RunProc(con, "Indent_Accessories");
                    QueryResult dt2 = RunProc(con, "Indent_NEW_ADIGA");
                    QueryResult dt3 = RunProc(con, "Indent_NEW_2023_CL1"); 
                    QueryResult dt4 = RunProc(con, "Today_Billable_OAs");
                    QueryResult dt5 = RunProc(con, "Indent_NEW_2023_SG");
                    QueryResult dt6 = RunProc(con, "Demo_Request_Pending");

                    if (dt == null || dt.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    // Suppress "overwrite?"/format-warning dialogs, which otherwise
                    // block headlessly-running automation and surface as
                    // "Exception from HRESULT: 0x800AC472".
                    oXL.DisplayAlerts = false;
                    oXL.Visible = false;
                    oWB = oXL.Workbooks.Add(missing);
                    Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    Microsoft.Office.Interop.Excel.Range excelCellrange;

                    oSheet.Name = "24HoursDeliveryPerformance";
                    WriteSheetData(oSheet, dt);

                    for (var i = 0; i < dt.RowCount; i++)
                    {
                        for (var j = 0; j < dt.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt.Rows[i][j - 2].ToString());
                                    string Unblock = dt.Rows[i][j + 1].ToString();
                                    string PrefDel = dt.Rows[i][j + 2].ToString();

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0)
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "OA Status")
                                {
                                    double BillableQty = Convert.ToDouble(dt.Rows[i][j - 1].ToString());
                                    double FGQty = Convert.ToDouble(dt.Rows[i][j - 2].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt.Rows[i][j - 3].ToString());
                                    string Unblock = dt.Rows[i][j].ToString();
                                    string PrefDel = dt.Rows[i][j + 1].ToString();
                                    int NoODays = 0;

                                    if (dt.Rows[i][j + 3].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt.Rows[i][j + 3]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked")
                                    {
                                        excelCellrange = oSheet.Range[oSheet.Cells[i + 2, j + 1], oSheet.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet1 (dt), row={i}, col={j}, ColName={dt.ColumnNames[j]}, ColumnCount={dt.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt1 == null || dt1.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet2.Name = "Accessories-ELM&CBE&AHD";
                    WriteSheetData(oSheet2, dt1);

                    for (var i = 0; i < dt1.RowCount; i++)
                    {
                        for (var j = 0; j < dt1.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt1.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt1.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt1.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt1.Rows[i][j - 5].ToString());
                                    string Unblock = dt1.Rows[i][j + 2].ToString();
                                    string PrefDel = dt1.Rows[i][j + 4].ToString();
                                    string ModlNm = dt1.Rows[i][j - 10].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt1.Rows[i][j - 2].ToString());
                                    double FGQty = Convert.ToDouble(dt1.Rows[i][j - 3].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt1.Rows[i][j - 7].ToString());
                                    string Unblock = dt1.Rows[i][j].ToString();
                                    string PrefDel = dt1.Rows[i][j + 2].ToString();
                                    string ModlNm = dt1.Rows[i][j - 12].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt1.Rows[i][j + 16].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt1.Rows[i][j + 16]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet2.Range[oSheet2.Cells[i + 2, j + 1], oSheet2.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet2 (dt1), row={i}, col={j}, ColName={dt1.ColumnNames[j]}, ColumnCount={dt1.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet3 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt2 == null || dt2.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet3.Name = "PendingOrders-ADIGA";
                    WriteSheetData(oSheet3, dt2);

                    for (var i = 0; i < dt2.RowCount; i++)
                    {
                        for (var j = 0; j < dt2.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt2.ColumnNames[j].ToString();

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt2.Rows[i][j - 2].ToString());
                                    double FGQty = Convert.ToDouble(dt2.Rows[i][j - 3].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt2.Rows[i][j - 5].ToString());
                                    string Unblock = dt2.Rows[i][j].ToString();
                                    string PrefDel = dt2.Rows[i][j + 1].ToString();
                                    string ModlNm = dt2.Rows[i][j - 9].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt2.Rows[i][j + 12].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt2.Rows[i][j + 12]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt2.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt2.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt2.Rows[i][j - 3].ToString());
                                    string Unblock = dt2.Rows[i][j + 2].ToString();
                                    string PrefDel = dt2.Rows[i][j + 3].ToString();
                                    string ModlNm = dt2.Rows[i][j - 7].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet3.Range[oSheet3.Cells[i + 2, j + 1], oSheet3.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet3 (dt2), row={i}, col={j}, ColName={dt2.ColumnNames[j]}, ColumnCount={dt2.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet4 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt3 == null || dt3.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet4.Name = "PendingOrders-ELM&CBE&AHD";
                    WriteSheetData(oSheet4, dt3);

                    for (var i = 0; i < dt3.RowCount; i++)
                    {
                        for (var j = 0; j < dt3.ColumnCount; j++)
                        {
                            try
                            {
                                string ColName = dt3.ColumnNames[j].ToString();

                                if (ColName == "Billable Qty")
                                {
                                    double BillableQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                    double FGQty = Convert.ToDouble(dt3.Rows[i][j - 1].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt3.Rows[i][j - 5].ToString());
                                    string Unblock = dt3.Rows[i][j + 3].ToString();
                                    string PrefDel = dt3.Rows[i][j + 7].ToString();
                                    string ModlNm = dt3.Rows[i][j - 10].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (FGQty >= TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                    }
                                    if (BillableQty == 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if (FGQty < TotalOAOpenQty && BillableQty != 0 && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);

                                        double BalanceQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                        double RemainigQty1 = Convert.ToDouble(dt3.Rows[i][dt3.ColumnIndex["Balance Qty to Be Billed"]].ToString());
                                        string BillableColor = dt3.Rows[i][dt3.ColumnIndex["Billable Colour"]].ToString();

                                        if (BillableColor == "Yellow" && PrefDel == "Partial Dispatch Allowed" && RemainigQty1 > 0)
                                        {
                                            excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                            FormattingExcelCells(excelCellrange, "#90EE90", System.Drawing.Color.Black, true);
                                        }
                                    }
                                }

                                if (ColName == "OA FINANCE STATUS (BLOCKED/UNBLOCKED)")
                                {
                                    double BillableQty = Convert.ToDouble(dt3.Rows[i][j - 3].ToString());
                                    double FGQty = Convert.ToDouble(dt3.Rows[i][j - 4].ToString());
                                    double TotalOAOpenQty = Convert.ToDouble(dt3.Rows[i][j - 8].ToString());
                                    string Unblock = dt3.Rows[i][j].ToString();
                                    string PrefDel = dt3.Rows[i][j + 4].ToString();
                                    string ModlNm = dt3.Rows[i][j - 14].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");
                                    int NoODays = 0;

                                    if (dt3.Rows[i][j + 22].ToString() != "")
                                    {
                                        NoODays = Convert.ToInt32(dt3.Rows[i][j + 22]);
                                    }

                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && Unblock != "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#FF0000", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && (PrefDel == "Partial Dispatch Allowed" || PrefDel == "Partial Dispatch") && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#0000FF", System.Drawing.Color.Black, true);
                                    }
                                    if ((FGQty >= TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0) || FGQty < TotalOAOpenQty && (BillableQty != 0 || BillableQty == 0)) && NoODays > 1 && Unblock == "Unblocked" && CBCT == false && uAlpha == false && EDC == false && iELR == false && ADC == false && VDC == false)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#800080", System.Drawing.Color.Black, true);
                                    }
                                }

                                if (ColName == "Bal Qty")
                                {
                                    double BalanceQty = Convert.ToDouble(dt3.Rows[i][j].ToString());
                                    double RemainigQty1 = Convert.ToDouble(dt3.Rows[i][dt3.ColumnIndex["Balance Qty to Be Billed"]].ToString());
                                    string Unblock = dt3.Rows[i][j + 10].ToString();
                                    string PrefDel = dt3.Rows[i][j + 14].ToString();
                                    string BillableColor = dt3.Rows[i][dt3.ColumnIndex["Billable Colour"]].ToString();
                                    string ModlNm = dt3.Rows[i][j - 3].ToString();
                                    bool CBCT = ModlNm.Contains("CBCT");
                                    bool uAlpha = ModlNm.Contains("µAlpha");
                                    bool EDC = ModlNm.Contains("EDC");
                                    bool iELR = ModlNm.Contains("iELR");
                                    bool ADC = ModlNm.Contains("ADC");
                                    bool VDC = ModlNm.Contains("VDC");

                                    if (BillableColor == "Yellow" && PrefDel == "Partial Dispatch Allowed" && RemainigQty1 > 0)
                                    {
                                        excelCellrange = oSheet4.Range[oSheet4.Cells[i + 2, j + 1], oSheet4.Cells[i + 2, j + 1]];
                                        FormattingExcelCells(excelCellrange, "#90EE90", System.Drawing.Color.Black, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed in sheet4 (dt3), row={i}, col={j}, ColName={dt3.ColumnNames[j]}, ColumnCount={dt3.ColumnCount}: {ex.Message}", ex);
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet5 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt4 == null || dt4.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet5.Name = "TodayBillableOAs";
                    WriteSheetData(oSheet5, dt4);

                    for (var i = 0; i < dt4.RowCount; i++)
                    {
                        for (var j = 0; j < dt4.ColumnCount; j++)
                        {
                            string ColName = dt4.ColumnNames[j].ToString();

                            if (ColName == "Billable Qty")
                            {
                                double BillableQty = Convert.ToDouble(dt4.Rows[i][j].ToString());
                                string Color = dt4.Rows[i][j + 29].ToString();

                                if (Color == "Green")
                                {
                                    excelCellrange = oSheet5.Range[oSheet5.Cells[i + 2, j + 1], oSheet5.Cells[i + 2, j + 1]];
                                    FormattingExcelCells(excelCellrange, "#008000", System.Drawing.Color.Black, true);
                                }
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet oSheet6 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt5 == null || dt5.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet6.Name = "SwitchGear";
                    WriteSheetData(oSheet6, dt5);

                    Microsoft.Office.Interop.Excel.Worksheet oSheet7 = oWB.Sheets.Add(missing, missing, 1, missing) as Microsoft.Office.Interop.Excel.Worksheet;
                    if (dt6 == null || dt6.ColumnCount == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    oSheet7.Name = "Demo_Request_Pending";
                    WriteSheetData(oSheet7, dt6);

                    oXL.Visible = false;
                    excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet2.Range[oSheet2.Cells[1, 1], oSheet2.Cells[1, dt1.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet3.Range[oSheet3.Cells[1, 1], oSheet3.Cells[1, dt2.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet4.Range[oSheet4.Cells[1, 1], oSheet4.Cells[1, dt3.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet5.Range[oSheet5.Cells[1, 1], oSheet5.Cells[1, dt4.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange = oSheet6.Range[oSheet6.Cells[1, 1], oSheet6.Cells[1, dt5.ColumnCount]];
                    excelCellrange.EntireColumn.AutoFit();

                    excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet2.Range[oSheet2.Cells[dt1.RowCount + 1, 1], oSheet2.Cells[dt1.RowCount + 1, dt1.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet3.Range[oSheet3.Cells[dt2.RowCount + 1, 1], oSheet3.Cells[dt2.RowCount + 1, dt2.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet4.Range[oSheet4.Cells[dt3.RowCount + 1, 1], oSheet4.Cells[dt3.RowCount + 1, dt3.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    oSheet.Activate();

                    Range rg = (Microsoft.Office.Interop.Excel.Range)oSheet4.Cells[1, 3];
                    rg.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg2 = (Microsoft.Office.Interop.Excel.Range)oSheet4.Cells[1, 28];
                    rg2.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg3 = (Microsoft.Office.Interop.Excel.Range)oSheet4.Cells[1, 32];
                    rg3.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg4 = (Microsoft.Office.Interop.Excel.Range)oSheet4.Cells[1, 33];
                    rg4.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg5 = (Microsoft.Office.Interop.Excel.Range)oSheet4.Cells[1, 47];
                    rg5.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg6 = (Microsoft.Office.Interop.Excel.Range)oSheet5.Cells[1, 3];
                    rg6.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg7 = (Microsoft.Office.Interop.Excel.Range)oSheet5.Cells[1, 25];
                    rg7.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg8 = (Microsoft.Office.Interop.Excel.Range)oSheet5.Cells[1, 29];
                    rg8.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg9 = (Microsoft.Office.Interop.Excel.Range)oSheet5.Cells[1, 30];
                    rg9.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    Range rg10 = (Microsoft.Office.Interop.Excel.Range)oSheet5.Cells[1, 44];
                    rg10.EntireColumn.NumberFormat = "DD-MM-YYYY";

                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;
                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;

                    excelCellrange = oSheet4.Range[oSheet4.Cells[1, 1], oSheet4.Cells[1, dt3.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet5.Range[oSheet5.Cells[1, 1], oSheet5.Cells[1, dt4.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet2.Range[oSheet2.Cells[1, 1], oSheet2.Cells[1, dt1.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    excelCellrange = oSheet3.Range[oSheet3.Cells[1, 1], oSheet3.Cells[1, dt2.ColumnCount]];
                    FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                    excelCellrange.EntireRow.Font.Bold = true;

                    oSheet4.Columns["V"].ColumnWidth = 0;
                    oSheet5.Activate();
                    oSheet4.Columns["V"].ColumnWidth = 0;
                    oSheet4.Columns["AL"].ColumnWidth = 0;
                    oXL.ActiveWindow.SplitRow = 1;
                    oXL.ActiveWindow.FreezePanes = true;
                    oXL.ActiveWindow.FreezePanes = true;

                    oWB.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Pending_Orders_" + now + "_" + Hr + ".xlsx");
                }
            }
            catch (Exception ex)
            {
                // Original swallowed this silently. Logging now so failures (like
                // out-of-bounds column offsets) are visible instead of invisible.
                try
                {
                    string logPath = System.Windows.Forms.Application.StartupPath + "\\AA_Error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
                    File.WriteAllText(logPath, ex.ToString());
                }
                catch
                {
                    // If even logging fails, fall back to the original silent behavior.
                }
            }
            finally
            {
                // Always release Excel, whether AA() succeeded or threw. This is what
                // prevents orphaned EXCEL.EXE processes — which, left running from a
                // prior failed/aborted run, are a common cause of
                // "Exception from HRESULT: 0x800AC472" on the next run.
                if (oWB != null)
                {
                    try { oWB.Close(false); } catch { /* ignore */ }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                    oWB = null;
                }
                if (oXL != null)
                {
                    try { oXL.Quit(); } catch { /* ignore */ }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                    oXL = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            // Excel Interop property sets (Interior.Color / Font.Color) occasionally
            // fail transiently with COMException 0x800AC472 ("Excel busy"/call
            // rejected), especially under rapid repeated calls or a non-interactive
            // session. Retrying briefly resolves the transient case; if it's not
            // transient (e.g. no interactive desktop session), it will keep failing
            // and surface after the retries are exhausted, same as before.
            RunComWithRetry(() => range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode));
            RunComWithRetry(() => range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor));
            if (IsFontbool == true)
            {
                RunComWithRetry(() => range.Font.Bold = IsFontbool);
            }
        }

        private static void RunComWithRetry(System.Action action, int maxRetries = 3)
        {
            int attempt = 0;
            while (true)
            {
                try
                {
                    action();
                    return;
                }
                catch (System.Runtime.InteropServices.COMException) when (attempt < maxRetries)
                {
                    attempt++;
                    System.Threading.Thread.Sleep(150 * attempt);
                }
            }
        }

        public static void DeleteRows(Workbook book, int n)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet workSheet in book.Sheets)
            {
                Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A1", "A" + n);
                Microsoft.Office.Interop.Excel.Range row = range.EntireRow;
                row.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
            }
        }

        public enum ExcelHorizontalAlignment
        {
            Left = 1
        }
        public enum ExcelVerticalAlignment
        {
            Distributed = 3
        }
    }
}