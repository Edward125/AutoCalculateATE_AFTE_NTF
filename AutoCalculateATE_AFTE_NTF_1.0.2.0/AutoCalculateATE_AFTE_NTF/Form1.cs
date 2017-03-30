using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace AutoCalculateATE_AFTE_NTF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Define Param

        string exeTitle = "Auto Calculate ATE AFTE NTF,Ver.:" + System.Windows.Forms.Application.ProductVersion + ",Author:Edward Song";
        
        Line AP2 = new Line("AP2");
        Line AP3 = new Line("AP3");
        Line AP4 = new Line("AP4");
        Line AP5 = new Line("AP5");
        Line AP6 = new Line("AP6");
        Line AP7 = new Line("AP7");
        Line AP8 = new Line("AP8");
        Line AP9 = new Line("AP9");

        string reportDate = DateTime.Now.AddDays(-1).ToString("M/d");

        string AFTE_Stage = "TD.DIP_FUNCTION_A";
        string ATE_Stage = "TA.DIP_ATE";

        string AFTE_Stage_= "TD#DIP_FUNCTION_A";
        string ATE_Stage_ = "TA#DIP_ATE";
      
        #endregion

        #region Form Code

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = exeTitle;
            //
            InitForm();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtSchedule.Text = string.Empty;
            txtNTF.Text = string.Empty;
            txtYRDN.Text = string.Empty;
            txtYRD.Text = string.Empty;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void txtSchedule_DoubleClick(object sender, EventArgs e)
        {
            OpenExcelFile(txtSchedule);
        }

        private void txtNTF_DoubleClick(object sender, EventArgs e)
        {
            OpenExcelFile(txtNTF);
        }

        private void txtYRDN_DoubleClick(object sender, EventArgs e)
        {
            OpenExcelFile(txtYRDN);
        }

        private void txtYRD_DoubleClick(object sender, EventArgs e)
        {
            OpenExcelFile(txtYRD);
        }


        private void btnCalculate_Click(object sender, EventArgs e)
        {

           // listBox1.Items.Clear();


            //OutputData2Excel();

            

            string reportDate = dateTimePicker1.Value.ToString("M/d");
            //MessageBox.Show(reportDate);
           // this.Enabled = false;
            // check file exits 
            if (!CheckFileExits(txtSchedule.Text))
                return;
            if (!CheckFileExits(txtNTF.Text))
                return;
            if (!CheckFileExits(txtYRDN.Text))
                return;
            if (!CheckFileExits(txtYRD.Text))
                return;

            this.Enabled = false;


            

            DataSet ds_Schedule = DataSetParse(txtSchedule.Text.Trim());
            dataGridView1.DataSource = ds_Schedule.Tables[0];
            this.Enabled = true;
            return;
            GetReportModel(ds_Schedule);
            for (int i = 0; i < ds_Schedule.Tables.Count; i++)
            {
                listBox1.Items.Add(ds_Schedule.Tables[i].ToString());
            }


            dataGridView1.DataSource = ds_Schedule.Tables[3];

            DataSet ds_NTF = DataSetParse(txtNTF.Text.Trim());
            GetNTFData(ds_NTF);

            //MessageBox.Show(AP2.LineTestokCountATE.ToString());

            DataSet ds_YR_D_N = DataSetParse(txtYRDN.Text.Trim());
            GetTestData(ds_YR_D_N, "DN");

            DataSet ds_YR_D = DataSetParse(txtYRD.Text.Trim());
            GetTestData(ds_YR_D, "D");
            this.Enabled = true;

            string filePath = System.Windows.Forms.Application.StartupPath + @"\" + "WCD_ATE_AFTE_NTF_" + DateTime.Now.Date.AddDays(-1).ToString("yyyyMMdd") + @".xls";

            if (!DownloadResouceFile(filePath))
            {
                this.Enabled = true;
                return;
            }

            OverwriteExcelSample(filePath);

            MessageBox.Show("Create file OK,File full path is " + @"'" + @filePath + @"'");
            this.Enabled = true;

        }


        //void displaytest(Line ap)
        //{
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Total_ATE:" + ap.LineTestCountATE);
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Day_ATE:" + ap.LineTestCountDayATE);
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Night_ATE:" + ap.LineTestCountNightATE);
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Total_AFTE:" + ap.LineTestCountAFTE);
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Day_AFTE:" + ap.LineTestCountDayAFTE);
        //    listBox1.Items.Add(ap.LineName + "_Test_Count_Night_AFTE:" + ap.LineTestCountNightAFTE);



        //    listBox1.Items.Add(ap.LineName + "_ATE_NTF:" + ap.NTF_ATE);
        //    listBox1.Items.Add(ap.LineName + "_ATE_Day_NTF:" + ap.NTF_DAY_ATE);
        //    listBox1.Items.Add(ap.LineName + "_ATE_Night_NTF:" + ap.NTF_NIGHT_ATE);

        //    listBox1.Items.Add(ap.LineName + "_AFTE_NTF:" + ap.NTF_AFTE);
        //    listBox1.Items.Add(ap.LineName + "_ATFE_Day_NTF:" + ap.NTF_DAY_AFTE);
        //    listBox1.Items.Add(ap.LineName + "_AFTE_Night_NTF:" + ap.NTF_NIGHT_AFTE);

        //}
        #endregion

        #region Init Form

        private void InitForm()
        {
            //
            txtSchedule.SetWatermark("双击此处，选择PC的schedule excel文件.");
            txtNTF .SetWatermark ("双击此处，选择SFCS下载的NTF excel文件.");
            txtYRDN.SetWatermark("双击此处，选择SFCS下载的白夜班良率数据excel文件.");
            txtYRD.SetWatermark("双击此处，选择SFCS下载的白班良率数据excel文件.");

            txtNote.Text = "注意：公司内部网络下载的Excel文档，为非标准Excel文档。请打开后另存为Excel文档后再导入小软件进行计算，否则将出错。";
            //
            //AP2.LineName = "AP2";
            //AP3.LineName = "AP3";
            //AP4.LineName = "AP4";
            //AP5.LineName = "AP5";
            //AP6.LineName = "AP6";
            //AP7.LineName = "AP7";
            //AP8.LineName = "AP8";
            //AP9.LineName = "AP9";

          //  MessageBox.Show(AP2.LineName);
            //AP2.LineTestCount = 1000;
            //AP2.LineTetestokCount = 10;
            //MessageBox.Show(AP2.NTF.ToString());
        }
        
        #endregion

        #region open file 

        private void OpenExcelFile(System.Windows.Forms. TextBox textbox)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "excel 2003(*.xls)|*.xls|excel 2010(*.xlsx)|*.xlsx|all files(*.*)|*.*";
            openfile.FilterIndex = 1;
            try
            {
                if (openfile.ShowDialog() == DialogResult.OK)
                {
                    textbox.Text = openfile.FileName;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }



        #endregion

        #region check excel file exits

        private bool CheckFileExits(string filepath)
        {
            if (string.IsNullOrEmpty(filepath))
                return false;
            if (System.IO.File.Exists(filepath.Trim()))
                return true;
            else
            {
                string filename = filepath.Trim().Substring(filepath.LastIndexOf('\\') + 1, filepath.Length - filepath.LastIndexOf('\\') - 1);
                MessageBox.Show(filename + " is not exits,retry please");
                return false;
            }
        }

        #endregion

        #region KillExcel
        private void KillExcel()
        {
            System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in excelProcess)
                p.Kill();
        }
        #endregion

        #region GetExcelDataSet

        /// <summary>     
        /// EXCEL数据转换DataSet     
        /// </summary>     
        /// <param name="filePath">文件全路径</param>     
        /// <param name="search">表名</param>     
        /// <returns></returns>         
        private DataSet GetDataSet(string fileName)
        {
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";
            OleDbConnection objConn = null;
            objConn = new OleDbConnection(strConn);
            objConn.Open();
            DataSet ds = new DataSet();
            //List<string> List = new List<string> { "收款金额", "代付关税", "垫付费用", "超期", "到账利润" };       
            List<string> List = new List<string> { };
            System.Data.DataTable dtSheetName = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            foreach (DataRow dr in dtSheetName.Rows)
            {
                if (dr["Table_Name"].ToString().Contains("$") && !dr[2].ToString().EndsWith("$"))
                {
                    continue;
                }
                string s = dr["Table_Name"].ToString();
                List.Add(s);
            }
            try
            {
                for (int i = 0; i < List.Count; i++)
                {
                    ds.Tables.Add(List[i]);
                    string SheetName = List[i];
                    string strSql = "select * from [" + SheetName + "]";
                    OleDbDataAdapter odbcCSVDataAdapter = new OleDbDataAdapter(strSql, objConn);
                    System.Data.DataTable dt = ds.Tables[i];
                    odbcCSVDataAdapter.Fill(dt);
                }
                return ds;
            }
            catch (Exception )
            {
                return null;
            }
            finally
            {
                objConn.Close();
                objConn.Dispose();
            }
        }


        #endregion

        #region ExcelToDataSet


        private DataSet ExcelToDataSet(string path)
        {

            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            if (conn.State.ToString() == "Open")
            {
                conn.Close();
            }

            try
            {
                conn.Open();
            }
            catch (Exception )
            {
                throw;
                //return null;

                //MessageBox.Show(ex.Message);
                //DataSet ds1 = null;
                //return ds1;
            }
            
            string s = conn.State.ToString();

            OleDbDataAdapter myCommand = null;
            DataSet ds = null;

            System.Data.DataTable yTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });

            string tableName = yTable.Rows[0]["Table_Name"].ToString();
            string strSel = "select * from [" + tableName + "]";//xls
            //strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strSel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            conn.Close();
            return ds;

            
        }

        #endregion

        #region Class Line

        public class Line
        {

            public readonly string LineName;
           // private double inNTF;

            //public string LineName
            //{
            //    set;
            //    get;
            //}


            private string modelname = string.Empty;
            public string ModelName
            {
                set
                {
                    modelname = value;
                }
                get
                {
                    return modelname  != string.Empty   ? modelname  : "NO Schedule";

                }
            }
            
            public Int32 LineTestCount { set; get;}

            public Int32 LineTestCountDay { set; get; }

            public Int32 LineTestCountNight { set; get; }


            public Int32 LineTestCountATE{set;get;}
            public Int32 LineTestCountDayATE{set;get;}
            public Int32 LineTestCountNightATE{set;get;}

            public Int32 LineTestCountAFTE{set;get;}
            public Int32 LineTestCountDayAFTE { set; get; }
            public Int32 LineTestCountNightAFTE { set; get; }

            public Int32 LineTesktokCount { set; get; }

            public Int32 LineTestokCountDay { set; get; }

            public Int32 LineTestokCountNight { set; get; }

            public Int32 LineTestokCountATE { set; get; }

            public Int32 LineTestokCountDayATE { set; get; }

            public Int32 LineTestokCountNightATE { set; get; }

            public Int32 LineTestokCountAFTE { set; get; }

            public Int32 LineTestokCountDayAFTE { set; get; }

            public Int32 LineTestokCountNightAFTE { set; get; }



            public string NTF
            {
                get
                {
                   return  calcntf(LineTestCount, LineTesktokCount);
                }
            }

            public string NTF_ATE
            {
                get
                {
                    return calcntf(LineTestCountATE, LineTestokCountATE);
                }
            }


            public string NTF_DAY_ATE
            {
                get
                {
                    return calcntf(LineTestCountDayATE, LineTestokCountDayATE);
                }
            }

            public string NTF_NIGHT_ATE
            {
                get
                {
                    return calcntf(LineTestCountNightATE, LineTestokCountNightATE);
                }
            }

            public string NTF_AFTE
            {
                get
                {
                    return calcntf(LineTestCountAFTE, LineTestokCountAFTE);
                }
            }


            public string NTF_DAY_AFTE
            {
                get
                {
                    return calcntf(LineTestCountDayAFTE, LineTestokCountDayAFTE);
                }
            }

            public string NTF_NIGHT_AFTE
            {
                get
                {
                    return calcntf(LineTestCountNightAFTE, LineTestokCountNightAFTE);
                }
            }
            //public string NTF_ATE
            //{
            //    get
            //    {
                  
            //    }
            //}

            public string calcntf(Int32 testcount, Int32 testokcout)
            {
                if( testcount > 0 )
                {
                    double temp = ((double)testokcout) / ((double)testcount);

                    return string.Format("{0:#.00%}", temp);
                }
                else 
                {
                    return "0";
                }
            }


            private Line()
                : this("Default Name")
            {
            }

            public Line(string newName)
            {
                LineName = newName;
            }
        }
        #endregion

        #region DataSet


        static DataSet DataSetParse(string fileName)
        {
           // string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);


            ////2003（Microsoft.Jet.Oledb.4.0）
            //string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            ////2010（Microsoft.ACE.OLEDB.12.0）
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);

             string connectionString  = string.Empty ;
            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);
            //MessageBox.Show(fi.Extension);

            if (fi.Extension == ".xls")
                connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
            if (fi.Extension == ".xlsx")
                connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", fileName);


            DataSet data = new DataSet();

            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    Console.WriteLine(sheetName);
                    var dataTable = new System.Data.DataTable(sheetName);             
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    con.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);
                    
                }
            }

            return data;
        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
          System.Data.  DataTable dt = null;
            con = new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            return excelSheetNames;
        }

        #endregion

        #region GetReportModel

        private void GetReportModel(DataSet ds)
        {

            string ColumnName = reportDate + @" Schedule";

            int tablecout = 0;
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                Console.WriteLine(ds.Tables[i].ToString());
                if (ds.Tables[i].ToString().ToUpper().Trim ().Contains ("SUMMARY"))
                {                    
                    tablecout = i;
                    break;
                }
            }

            Console.WriteLine("Table Summary=" + tablecout);
            
            //get ColumnName 
            for (int i = 0; i <ds.Tables[tablecout ].Columns .Count; i++)
            {
                MessageBox.Show(ds.Tables[0].Columns[i].ToString());
                if (ds.Tables[0].Columns[i].ToString().ToUpper ().Trim () == reportDate.ToUpper ().Trim ())
                {
                    ColumnName = ds.Tables[0].Columns[i].ToString();
                    break;
                }
            }
           // MessageBox.Show (ColumnName );
            dataGridView1.DataSource = ds.Tables[tablecout];




            for (int i = 0; i < ds.Tables[tablecout].Rows.Count; i++)
            {
                // MessageBox.Show(ds.Tables[0].Rows[i][0].ToString());

                string Line = ds.Tables[tablecout].Rows[i][0].ToString();
               

                string Model = ds.Tables[tablecout].Rows[i][ColumnName].ToString();

                switch (Line)
                {
                    case "AS2":
                        AP2.ModelName = Model;
                        break;
                    case "AS3":
                        AP3.ModelName = Model;
                        break;
                    case "AS4":
                        AP4.ModelName = Model;
                        break;
                    case "AS5":
                        AP5.ModelName = Model;
                        break;
                    case "AS6":
                        AP6.ModelName = Model;
                        break;
                    case "AS7":
                        AP7.ModelName = Model;
                        break;
                    case "AS8":
                        AP8.ModelName = Model;
                        break;
                    case "AS9":
                        AP9.ModelName = Model;
                        break;
                    default:
                        break;
                }
            }

        }

        #endregion

        #region GetNTFData

        private void GetNTFData(DataSet ds)
        {

            DateTime s_datetime = DateTime.Parse(reportDate + "/" + DateTime.Now.Year.ToString() + " 8:30:00 AM");
            DateTime e_datetime = DateTime.Parse(reportDate + "/" + DateTime.Now.Year.ToString() + " 8:30:00 PM");



            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                //Day shift
                DateTime  testtime = DateTime.Parse ( ds.Tables[0].Rows[i]["Transation Date"].ToString());

                string LineName = ds.Tables[0].Rows[i]["Line"].ToString();
                string LineStage = ds.Tables[0].Rows[i]["Retest Stage"].ToString().Trim().ToUpper();
                
                if ((testtime > s_datetime) & (testtime < e_datetime))
                {
                    //ate
                    if (LineStage  == ATE_Stage.ToUpper ())
                    {
                        switch (LineName )
                        {
                            case "AP2":
                                AP2.LineTestokCountDayATE += 1;
                                break;
                            case "AP3":
                                AP3.LineTestokCountDayATE += 1;
                                break;
                            case "AP4":
                                AP4.LineTestokCountDayATE += 1;
                                break;
                            case "AP5":
                                AP5.LineTestokCountDayATE +=1;
                                break;
                            case "AP6":
                                AP6.LineTestokCountDayATE +=1;
                                break;
                            case "AP7":
                                AP7.LineTestokCountDayATE +=1;
                                break;
                            case "AP8":
                                AP8.LineTestokCountDayATE += 1;
                                break;
                            case "AP9":
                                AP9.LineTestokCountDayATE += 1;
                                break;
                            default:
                                break;
                        }
                    }

                    //afte
                    if (LineStage   == AFTE_Stage.ToUpper ())
                    {
                        switch (LineName)
                        {
                            case "AP2":
                                AP2.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP3":
                                AP3.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP4":
                                AP4.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP5":
                                AP5.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP6":
                                AP6.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP7":
                                AP7.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP8":
                                AP8.LineTestokCountDayAFTE += 1;
                                break;
                            case "AP9":
                                AP9.LineTestokCountDayAFTE += 1;
                                break;
                            default:
                                break;
                        }
                    }

                }
                else
                {
                    //ate
                    if (LineStage  == ATE_Stage.ToUpper ())
                    {
                        switch (LineName)
                        {
                            case "AP2":
                                AP2.LineTestokCountNightATE += 1;
                                break;
                            case "AP3":
                                AP3.LineTestokCountNightATE += 1;
                                break;
                            case "AP4":
                                AP4.LineTestokCountNightATE += 1;
                                break;
                            case "AP5":
                                AP5.LineTestokCountNightATE += 1;
                                break;
                            case "AP6":
                                AP6.LineTestokCountNightATE += 1;
                                break;
                            case "AP7":
                                AP7.LineTestokCountNightATE += 1;
                                break;
                            case "AP8":
                                AP8.LineTestokCountNightATE += 1;
                                break;
                            case "AP9":
                                AP9.LineTestokCountNightATE += 1;
                                break;
                            default:
                                break;
                        }
                    }

                    //afte
                    if (LineStage  == AFTE_Stage.ToUpper())
                    {
                        switch (LineName)
                        {
                            case "AP2":
                                AP2.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP3":
                                AP3.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP4":
                                AP4.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP5":
                                AP5.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP6":
                                AP6.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP7":
                                AP7.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP8":
                                AP8.LineTestokCountNightAFTE += 1;
                                break;
                            case "AP9":
                                AP9.LineTestokCountNightAFTE += 1;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }


            UpdateLineTestOKQty(AP2);
            UpdateLineTestOKQty(AP3);
            UpdateLineTestOKQty(AP4);
            UpdateLineTestOKQty(AP5);
            UpdateLineTestOKQty(AP6);
            UpdateLineTestOKQty(AP7);
            UpdateLineTestOKQty(AP8);
            UpdateLineTestOKQty(AP9);

        }


        void UpdateLineTestOKQty(Line ap)
        {
            ap.LineTestokCountATE = ap.LineTestokCountDayATE + ap.LineTestokCountNightATE;
            ap.LineTestokCountAFTE = ap.LineTestokCountDayAFTE + ap.LineTestokCountNightAFTE;
        }
        #endregion

        #region GetTestData
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="shiftflag">DN,D</param>
        private void GetTestData(DataSet ds, string shiftflag)
        {

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string lineName = ds.Tables[0].Rows[i]["F1"].ToString();
                string ate = ds.Tables[0].Rows[i][ATE_Stage_].ToString();
                string afte = ds.Tables[0].Rows[i][AFTE_Stage_].ToString();
                switch (lineName )
                {
                    case "AP2":
                        if (shiftflag == "D")
                        {

                            AP2.LineTestCountDayATE = AP2.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP2.LineTestCountDayAFTE =AP2.LineTestCountDayAFTE+ Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP2.LineTestCountATE = AP2.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP2.LineTestCountAFTE = AP2.LineTestCountAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP3":
                        if (shiftflag == "D")
                        {
                            AP3.LineTestCountDayATE = AP3.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP3.LineTestCountDayAFTE = AP3.LineTestCountDayAFTE+ Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP3.LineTestCountATE = AP3.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP3.LineTestCountAFTE = AP3.LineTestCountAFTE+ Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP4":
                         if (shiftflag == "D")
                        {
                            AP4.LineTestCountDayATE = AP4.LineTestCountDayATE+ Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP4.LineTestCountDayAFTE = AP4.LineTestCountDayAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP4.LineTestCountATE =   AP4.LineTestCountATE+ Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP4.LineTestCountAFTE = AP4.LineTestCountAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP5":
                         if (shiftflag == "D")
                        {
                            AP5.LineTestCountDayATE = AP5.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP5.LineTestCountDayAFTE =AP5.LineTestCountDayAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP5.LineTestCountATE =  AP5.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP5.LineTestCountAFTE = AP5.LineTestCountAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP6":
                         if (shiftflag == "D")
                        {
                            AP6.LineTestCountDayATE = AP6.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP6.LineTestCountDayAFTE = AP6.LineTestCountDayAFTE +Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP6.LineTestCountATE = AP6.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP6.LineTestCountAFTE =  AP6.LineTestCountAFTE+ Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP7":
                         if (shiftflag == "D")
                        {
                            AP7.LineTestCountDayATE = AP7.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP7.LineTestCountDayAFTE =  AP7.LineTestCountDayAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP7.LineTestCountATE = AP7.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP7.LineTestCountAFTE = AP7.LineTestCountAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP8":
                         if (shiftflag == "D")
                        {
                            AP8.LineTestCountDayATE = AP8.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP8.LineTestCountDayAFTE = AP8.LineTestCountDayAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        if (shiftflag =="DN")
                        {
                            AP8.LineTestCountATE = AP8.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP8.LineTestCountAFTE = AP8.LineTestCountAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                        }
                        break;
                    case "AP9":
                         if (shiftflag == "D")
                        {
                            AP9.LineTestCountDayATE = AP9.LineTestCountDayATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP9.LineTestCountDayAFTE = AP9.LineTestCountDayAFTE + Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());  
                        }
                        if (shiftflag =="DN")
                        {
                            AP9.LineTestCountATE =  AP9.LineTestCountATE + Convert.ToInt16(ds.Tables[0].Rows[i][ATE_Stage_].ToString());
                            AP9.LineTestCountAFTE =  AP9.LineTestCountAFTE  +Convert.ToInt16(ds.Tables[0].Rows[i][AFTE_Stage_].ToString());
                            
                        }
                        break;
                    default:
                        break;
                }

            }

            AP2.LineTestCountNightATE = AP2.LineTestCountATE - AP2.LineTestCountDayATE;
            AP2.LineTestCountNightAFTE = AP2.LineTestCountAFTE - AP2.LineTestCountDayAFTE;

            AP3.LineTestCountNightATE = AP3.LineTestCountATE - AP3.LineTestCountDayATE;
            AP3.LineTestCountNightAFTE = AP3.LineTestCountAFTE - AP3.LineTestCountDayAFTE;

            AP4.LineTestCountNightATE = AP4.LineTestCountATE - AP4.LineTestCountDayATE;
            AP4.LineTestCountNightAFTE = AP4.LineTestCountAFTE - AP4.LineTestCountDayAFTE;

            AP5.LineTestCountNightATE = AP5.LineTestCountATE - AP5.LineTestCountDayATE;
            AP5.LineTestCountNightAFTE = AP5.LineTestCountAFTE - AP5.LineTestCountDayAFTE;

            AP6.LineTestCountNightATE = AP6.LineTestCountATE - AP6.LineTestCountDayATE;
            AP6.LineTestCountNightAFTE = AP6.LineTestCountAFTE - AP6.LineTestCountDayAFTE;

            AP7.LineTestCountNightATE = AP7.LineTestCountATE - AP7.LineTestCountDayATE;
            AP7.LineTestCountNightAFTE = AP7.LineTestCountAFTE - AP7.LineTestCountDayAFTE;

            AP8.LineTestCountNightATE = AP8.LineTestCountATE - AP8.LineTestCountDayATE;
            AP8.LineTestCountNightAFTE = AP8.LineTestCountAFTE - AP8.LineTestCountDayAFTE;

            AP9.LineTestCountNightATE = AP9.LineTestCountATE - AP9.LineTestCountDayATE;
            AP9.LineTestCountNightAFTE = AP9.LineTestCountAFTE - AP9.LineTestCountDayAFTE;

        }
        #endregion



        #region OutputData2Excel()

        private void OutputData2Excel()
        {
            string filePath = System.Windows.Forms.Application.StartupPath + @"\" + "WCD_ATE_AFTE_NTF_" + DateTime.Now.Date.AddDays(-1).ToString("yyyyMMdd") + @".xls";
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Workbook wBook = appExcel.Workbooks.Add(true);
            Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
            wSheet.Name = "ATE NTF";

            wSheet.Cells[1, 1] = "ATE Re-Test OK F.R. Summary";
            wSheet.Cells[2, 1] = "Line";
            wSheet.Cells[2, 2] = "Model";
            wSheet.Cells[2, 3] = "Fixture ID";
            wSheet.Cells[2, 4] = "Input Q'ty";
            wSheet.Cells[2, 5] = "ReTest OK Q'ty";
            wSheet.Cells[2, 6] = "Retest OK F.R.(Target<2%)";
            wSheet.Cells[2, 7] = "Remark";



            //appExcel.get_Range(wSheet.Cells[1, 2], wSheet.Cells[2, 6]).Merge();
            wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[2, 2]).Merge();

           // //Borders.LineStyle 单元格边框线
           // //Excel.Range excelRange = _workSheet.get_Range(_workSheet.Cells[2, 2], _workSheet.Cells[4, 6]);

           


           // Range excelRange = wSheet.get_Range(wSheet.Cells[2, 2],Type.Missing );
           // //单元格边框线类型(线型,虚线型)
           // excelRange.Borders.LineStyle = 1;
           // excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
           // //指定单元格下边框线粗细,和色彩
           // excelRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

           // excelRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).ColorIndex = 3;

           // //设置字体大小
           // excelRange.Font.Size = 15;
           // //设置字体是否有下划线
           // excelRange.Font.Underline = true;

           // //设置字体在单元格内的对其方式
           // excelRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
           // //设置单元格的宽度
           // excelRange.ColumnWidth = 15;
           // //设置单元格的背景色
           // excelRange.Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 204, 153).ToArgb();
           // // 给单元格加边框
           // excelRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThick,
           // XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());
           // //自动调整列宽
           // excelRange.EntireColumn.AutoFit();
           // // 文本水平居中方式
           // excelRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
           // //文本自动换行
           // excelRange.WrapText = true;
           // //填充颜色为淡紫色
           // excelRange.Interior.ColorIndex = 39;

           // //合并单元格
           // excelRange.Merge(excelRange.MergeCells);
           //wSheet .get_Range("A15", "B15").Merge(wSheet.get_Range("A15", "B15").MergeCells);


            //设置禁止弹出保存和覆盖的询问提示框   
            appExcel.DisplayAlerts = false;
            appExcel.AlertBeforeOverwriting = false;
            //保存工作簿   
            wBook.Save();
            //保存excel文件   
            appExcel.Save(filePath);
            appExcel.SaveWorkspace(filePath);
            appExcel.Quit();
            appExcel = null;



        }


        #region downloadresoucefile

        private bool DownloadResouceFile(string filename)
        {
            if (System.IO.File.Exists(filename)) 
                return true ;


            
            byte[] file = global::AutoCalculateATE_AFTE_NTF.Properties.Resources.sample;
            try
            {
                System.IO.FileStream fsObj = new System.IO.FileStream(filename, System.IO.FileMode.Create);
                fsObj.Write(file, 0, file.Length);
                fsObj.Close();

                //System.IO.FileInfo fi = new System.IO.FileInfo(filename);
                //fi.Attributes = System.IO.FileAttributes.Hidden;

                return true;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return false;
            }


        }
        #endregion

        #region OverwriteExcelSample

        private void OverwriteExcelSample(string filename)
        {
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Workbook wBook = appExcel.Workbooks.Open(filename);
            Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
            //

            ////////ATE
            wSheet.Cells[3, 2] = AP2.ModelName;
            wSheet.Cells[3, 4] = AP2.LineTestCountATE ;
            wSheet.Cells[3, 5] = AP2.LineTestokCountATE;
            //
            wSheet.Cells[4, 2] = AP3.ModelName;
            wSheet.Cells[4, 4] = AP3.LineTestCountATE;
            wSheet.Cells[4, 5] = AP3.LineTestokCountATE;
            //
            wSheet.Cells[5, 2] = AP4.ModelName;
            wSheet.Cells[5, 4] = AP4.LineTestCountATE;
            wSheet.Cells[5, 5] = AP4.LineTestokCountATE;
            //
            wSheet.Cells[6, 2] = AP5.ModelName;
            wSheet.Cells[6, 4] = AP5.LineTestCountATE;
            wSheet.Cells[6, 5] = AP5.LineTestokCountATE;
            //
            wSheet.Cells[7, 2] = AP6.ModelName;
            wSheet.Cells[7, 4] = AP6.LineTestCountATE;
            wSheet.Cells[7, 5] = AP6.LineTestokCountATE;
            //
            wSheet.Cells[8, 2] = AP7.ModelName;
            wSheet.Cells[8, 4] = AP7.LineTestCountATE;
            wSheet.Cells[8, 5] = AP7.LineTestokCountATE;
            //
            wSheet.Cells[9, 2] = AP8.ModelName;
            wSheet.Cells[9, 4] = AP8.LineTestCountATE;
            wSheet.Cells[9, 5] = AP8.LineTestokCountATE;
            //
            wSheet.Cells[10, 2] = AP9.ModelName;
            wSheet.Cells[10, 4] = AP9.LineTestCountATE;
            wSheet.Cells[10, 5] = AP9.LineTestokCountATE;
            ///////AFTE
            wSheet.Cells[15, 2] = AP2.ModelName;
            wSheet.Cells[15, 4] = AP2.LineTestCountAFTE;
            wSheet.Cells[15, 5] = AP2.LineTestokCountAFTE;
            //
            wSheet.Cells[16, 2] = AP3.ModelName;
            wSheet.Cells[16, 4] = AP3.LineTestCountAFTE;
            wSheet.Cells[16, 5] = AP3.LineTestokCountAFTE;
            //
            wSheet.Cells[17, 2] = AP4.ModelName;
            wSheet.Cells[17, 4] = AP4.LineTestCountAFTE;
            wSheet.Cells[17, 5] = AP4.LineTestokCountAFTE;
            //
            wSheet.Cells[18, 2] = AP5.ModelName;
            wSheet.Cells[18, 4] = AP5.LineTestCountAFTE;
            wSheet.Cells[18, 5] = AP5.LineTestokCountAFTE;
            //
            wSheet.Cells[19, 2] = AP6.ModelName;
            wSheet.Cells[19, 4] = AP6.LineTestCountAFTE;
            wSheet.Cells[19, 5] = AP6.LineTestokCountAFTE;
            //
            wSheet.Cells[20, 2] = AP7.ModelName;
            wSheet.Cells[20, 4] = AP7.LineTestCountAFTE;
            wSheet.Cells[20, 5] = AP7.LineTestokCountAFTE;
            //
            wSheet.Cells[21, 2] = AP8.ModelName;
            wSheet.Cells[21, 4] = AP8.LineTestCountAFTE;
            wSheet.Cells[21, 5] = AP8.LineTestokCountAFTE;
            //
            wSheet.Cells[22, 2] = AP9.ModelName;
            wSheet.Cells[22, 4] = AP9.LineTestCountAFTE;
            wSheet.Cells[22, 5] = AP9.LineTestokCountAFTE;
            //
            //////ATE DAY
            wSheet.Cells[27, 2] = AP2.ModelName;
            wSheet.Cells[27, 4] = AP2.LineTestCountDayATE;
            wSheet.Cells[27, 5] = AP2.LineTestokCountDayATE;
            //
            wSheet.Cells[28, 2] = AP3.ModelName;
            wSheet.Cells[28, 4] = AP3.LineTestCountDayATE;
            wSheet.Cells[28, 5] = AP3.LineTestokCountDayATE;
            //
            wSheet.Cells[29, 2] = AP4.ModelName;
            wSheet.Cells[29, 4] = AP4.LineTestCountDayATE;
            wSheet.Cells[29, 5] = AP4.LineTestokCountDayATE;
            //
            wSheet.Cells[30, 2] = AP5.ModelName;
            wSheet.Cells[30, 4] = AP5.LineTestCountDayATE;
            wSheet.Cells[30, 5] = AP5.LineTestokCountDayATE;
            //
            wSheet.Cells[31, 2] = AP6.ModelName;
            wSheet.Cells[31, 4] = AP6.LineTestCountDayATE;
            wSheet.Cells[31, 5] = AP6.LineTestokCountDayATE;
            //
            wSheet.Cells[32, 2] = AP7.ModelName;
            wSheet.Cells[32, 4] = AP7.LineTestCountDayATE;
            wSheet.Cells[32, 5] = AP7.LineTestokCountDayATE;
            //
            wSheet.Cells[33, 2] = AP8.ModelName;
            wSheet.Cells[33, 4] = AP8.LineTestCountDayATE;
            wSheet.Cells[33, 5] = AP8.LineTestokCountDayATE;
            //
            wSheet.Cells[34, 2] = AP9.ModelName;
            wSheet.Cells[34, 4] = AP9.LineTestCountDayATE;
            wSheet.Cells[34, 5] = AP9.LineTestokCountDayATE;

            ///////ATE NIGHT
            wSheet.Cells[39, 2] = AP2.ModelName;
            wSheet.Cells[39, 4] = AP2.LineTestCountNightATE;
            wSheet.Cells[39, 5] = AP2.LineTestokCountNightATE;
            //
            wSheet.Cells[40, 2] = AP3.ModelName;
            wSheet.Cells[40, 4] = AP3.LineTestCountNightATE;
            wSheet.Cells[40, 5] = AP3.LineTestokCountNightATE;
            //
            wSheet.Cells[41, 2] = AP4.ModelName;
            wSheet.Cells[41, 4] = AP4.LineTestCountNightATE;
            wSheet.Cells[41, 5] = AP4.LineTestokCountNightATE;
            //
            wSheet.Cells[42, 2] = AP5.ModelName;
            wSheet.Cells[42, 4] = AP5.LineTestCountNightATE;
            wSheet.Cells[42, 5] = AP5.LineTestokCountNightATE;
            //
            wSheet.Cells[43, 2] = AP6.ModelName;
            wSheet.Cells[43, 4] = AP6.LineTestCountNightATE;
            wSheet.Cells[43, 5] = AP6.LineTestokCountNightATE;
            //
            wSheet.Cells[44, 2] = AP7.ModelName;
            wSheet.Cells[44, 4] = AP7.LineTestCountNightATE;
            wSheet.Cells[44, 5] = AP7.LineTestokCountNightATE;
            //
            wSheet.Cells[45, 2] = AP8.ModelName;
            wSheet.Cells[45, 4] = AP8.LineTestCountNightATE;
            wSheet.Cells[45, 5] = AP8.LineTestokCountNightATE;
            //
            wSheet.Cells[46, 2] = AP9.ModelName;
            wSheet.Cells[46, 4] = AP9.LineTestCountNightATE;
            wSheet.Cells[46, 5] = AP9.LineTestokCountNightATE;

            //////AFTE DAY
            wSheet.Cells[51, 2] = AP2.ModelName;
            wSheet.Cells[51, 4] = AP2.LineTestCountDayAFTE;
            wSheet.Cells[51, 5] = AP2.LineTestokCountDayAFTE;
            //
            wSheet.Cells[52, 2] = AP3.ModelName;
            wSheet.Cells[52, 4] = AP3.LineTestCountDayAFTE;
            wSheet.Cells[52, 5] = AP3.LineTestokCountDayAFTE;
            //
            wSheet.Cells[53, 2] = AP4.ModelName;
            wSheet.Cells[53, 4] = AP4.LineTestCountDayAFTE;
            wSheet.Cells[53, 5] = AP4.LineTestokCountDayAFTE;
            //
            wSheet.Cells[54, 2] = AP5.ModelName;
            wSheet.Cells[54, 4] = AP5.LineTestCountDayAFTE;
            wSheet.Cells[54, 5] = AP5.LineTestokCountDayAFTE;
            //
            wSheet.Cells[55, 2] = AP6.ModelName;
            wSheet.Cells[55, 4] = AP6.LineTestCountDayAFTE;
            wSheet.Cells[55, 5] = AP6.LineTestokCountDayAFTE;
            //
            wSheet.Cells[56, 2] = AP7.ModelName;
            wSheet.Cells[56, 4] = AP7.LineTestCountDayAFTE;
            wSheet.Cells[56, 5] = AP7.LineTestokCountDayAFTE;
            //
            wSheet.Cells[57, 2] = AP8.ModelName;
            wSheet.Cells[57, 4] = AP8.LineTestCountDayAFTE;
            wSheet.Cells[57, 5] = AP8.LineTestokCountDayAFTE;
            //
            wSheet.Cells[58, 2] = AP9.ModelName;
            wSheet.Cells[58, 4] = AP9.LineTestCountDayAFTE;
            wSheet.Cells[58, 5] = AP9.LineTestokCountDayAFTE;

            ///////AFTE NIGHT
            wSheet.Cells[63, 2] = AP2.ModelName;
            wSheet.Cells[63, 4] = AP2.LineTestCountNightAFTE;
            wSheet.Cells[63, 5] = AP2.LineTestokCountNightAFTE;
            //
            wSheet.Cells[64, 2] = AP3.ModelName;
            wSheet.Cells[64, 4] = AP3.LineTestCountNightAFTE;
            wSheet.Cells[64, 5] = AP3.LineTestokCountNightAFTE;
            //
            wSheet.Cells[65, 2] = AP4.ModelName;
            wSheet.Cells[65, 4] = AP4.LineTestCountNightAFTE;
            wSheet.Cells[65, 5] = AP4.LineTestokCountNightAFTE;
            //
            wSheet.Cells[66, 2] = AP5.ModelName;
            wSheet.Cells[66, 4] = AP5.LineTestCountNightAFTE;
            wSheet.Cells[66, 5] = AP5.LineTestokCountNightAFTE;
            //
            wSheet.Cells[67, 2] = AP6.ModelName;
            wSheet.Cells[67, 4] = AP6.LineTestCountNightAFTE;
            wSheet.Cells[67, 5] = AP6.LineTestokCountNightAFTE;
            //
            wSheet.Cells[68, 2] = AP7.ModelName;
            wSheet.Cells[68, 4] = AP7.LineTestCountNightAFTE;
            wSheet.Cells[68, 5] = AP7.LineTestokCountNightAFTE;
            //
            wSheet.Cells[69, 2] = AP8.ModelName;
            wSheet.Cells[69, 4] = AP8.LineTestCountNightAFTE;
            wSheet.Cells[69, 5] = AP8.LineTestokCountNightAFTE;
            //
            wSheet.Cells[70, 2] = AP9.ModelName;
            wSheet.Cells[70, 4] = AP9.LineTestCountNightAFTE;
            wSheet.Cells[70, 5] = AP9.LineTestokCountNightAFTE;


            //设置禁止弹出保存和覆盖的询问提示框   
            appExcel.DisplayAlerts = false;
            appExcel.AlertBeforeOverwriting = false;
            //保存工作簿   
            wBook.Save();
            //保存excel文件   
            appExcel.Save();
            appExcel.SaveWorkspace();
            appExcel.Quit();
            appExcel = null;

        }

        #endregion

        private void txtNote_TextChanged(object sender, EventArgs e)
        {

        }


        #endregion
    }
     
}
