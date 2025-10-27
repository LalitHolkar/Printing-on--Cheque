using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Web.UI;
using System.Web;
using System.Reflection;

namespace ChequePrint
{
    public partial class Form1 : Form
    {
        string Message = "";
        public string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        public string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        int rowCnt = 0;
        //  string Acc_PAyee = "";
        public Form1()
        {
            InitializeComponent();
            //  label1.Parent =;
            //panelSingle.Visible = true;
            label1.BackColor = Color.Transparent;
            label2.BackColor = Color.Transparent;
            label3.BackColor = Color.Transparent;
            label4.BackColor = Color.Transparent;
            label5.BackColor = Color.Transparent;
            lblAmountInWord.BackColor = Color.Transparent;
            chkacpay.BackColor = Color.Transparent;
            panelSingle.BackColor = Color.Transparent;
            //radioButton2.BackColor = Color.Transparent;
            rdMultiple.BackColor = Color.Transparent;
            rdSingle.BackColor = Color.Transparent;
            panelMultiple.BackColor = Color.Transparent;
            panelSingle.Visible = true;
            panelMultiple.Visible = false;
            button4.Visible = false;
            button3.Visible = false;
            button2.Visible = true;
            button1.Visible = true;
            btnBrowse.Visible = false;

        }
        PrintDialog ObjPrintDialog = new PrintDialog();
        //  PrintDocument objPrintDocumnt = new PrintDocument();
        //private System.Windows.Forms.OpenFileDialog openFileDialog1;  
        // this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
        private void txtAmount_Leave(object sender, EventArgs e)
        {
            if (txtAmount.Text != "")
            {
                lblAmountInWord.Text = words(Convert.ToDouble(txtAmount.Text));
            }
        }

        #region convert number to words
        public string words(double? numbers, Boolean paisaconversion = false)
        {
            var pointindex = numbers.ToString().IndexOf(".");
            var paisaamt = 0;
            if (pointindex > 0)
                paisaamt = Convert.ToInt32(numbers.ToString().Substring(pointindex + 1, 2));

            int number = Convert.ToInt32(numbers);

            if (number == 0) return "Zero";
            if (number == -2147483648) return "Minus Two Hundred and Fourteen Crore Seventy Four Lakh Eighty Three Thousand Six Hundred and Forty Eight";
            int[] num = new int[4];
            int first = 0;
            int u, h, t;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (number < 0)
            {
                sb.Append("Minus ");
                number = -number;
            }
            string[] words0 = { "", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine " };
            string[] words1 = { "Ten ", "Eleven ", "Twelve ", "Thirteen ", "Fourteen ", "Fifteen ", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen " };
            string[] words2 = { "Twenty ", "Thirty ", "Forty ", "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety " };
            string[] words3 = { "Thousand ", "Lakh ", "Crore " };
            num[0] = number % 1000; // units
            num[1] = number / 1000;
            num[2] = number / 100000;
            num[1] = num[1] - 100 * num[2]; // thousands
            num[3] = number / 10000000; // crores
            num[2] = num[2] - 100 * num[3]; // lakhs
            for (int i = 3; i > 0; i--)
            {
                if (num[i] != 0)
                {
                    first = i;
                    break;
                }
            }
            for (int i = first; i >= 0; i--)
            {
                if (num[i] == 0) continue;
                u = num[i] % 10; // ones
                t = num[i] / 10;
                h = num[i] / 100; // hundreds
                t = t - 10 * h; // tens
                if (h > 0) sb.Append(words0[h] + "Hundred ");
                if (u > 0 || t > 0)
                {
                    //  if (h > 0 || i == 0) sb.Append("and ");
                    if (t == 0)
                        sb.Append(words0[u]);
                    else if (t == 1)
                        sb.Append(words1[u]);
                    else
                        sb.Append(words2[t - 2] + words0[u]);
                }
                if (i != 0) sb.Append(words3[i - 1]);
            }

            if (paisaamt == 0 && paisaconversion == false)
            {
                sb.Append(" only");
            }
            else if (paisaamt > 0)
            {
                var paisatext = words(paisaamt, true);
                sb.AppendFormat("rupees {0} paise only", paisatext);
            }
            return sb.ToString().TrimEnd();
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {


                if (txtAmount.Text == "" && txtName.Text == "")
                {
                    MessageBox.Show("Please Fill Required Details");
                }
                else
                {
                    if (printDialog1.ShowDialog() == DialogResult.OK)
                    {
                        printDocument1.DefaultPageSettings.Landscape = false;
                        printDocument1.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("PaperA4", 840, 1180);
                        printDocument1.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                        printDocument1.Print();
                    }
                    clear();
                    MessageBox.Show("Cheque Print Succeesfull");
                }
            }
            catch (Exception ex)
            {
                Message = ex.Message;
            }
        }


        private void txtAmount_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }


        void clear()
        {
            txtAmount.Clear();
            lblAmountInWord.Text = "";
            txtName.Clear();
            chkacpay.Checked = false;
            txtDate.Text = "";
        }
        private void button2_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {

            if (txtDate.Text != "" || txtDate.Text != null)
            {
                e.Graphics.DrawString(txtDate.Text, new Font("verdana", 10), new
            SolidBrush(Color.Black), new RectangleF(620, 5,
            printDocument1.DefaultPageSettings.PrintableArea.Width,
           printDocument1.DefaultPageSettings.PrintableArea.Height));
            }
            if (chkacpay.Checked == true)
            {
                e.Graphics.DrawString("A/C PAYEE ONLY", new Font("verdana", 10), new
                SolidBrush(Color.Black), new RectangleF(300, 220,
                printDocument1.DefaultPageSettings.PrintableArea.Width,
                printDocument1.DefaultPageSettings.PrintableArea.Height));
            }
            e.Graphics.DrawString(txtName.Text + "****", new Font("verdana", 10), new
            SolidBrush(Color.Black), new RectangleF(100, 60,
           printDocument1.DefaultPageSettings.PrintableArea.Width,
           printDocument1.DefaultPageSettings.PrintableArea.Height));

            e.Graphics.DrawString(lblAmountInWord.Text, new Font("verdana", 10), new
            SolidBrush(Color.Black), new RectangleF(130, 90,
            printDocument1.DefaultPageSettings.PrintableArea.Width,
            printDocument1.DefaultPageSettings.PrintableArea.Height));

            e.Graphics.DrawString("**" + txtAmount.Text + "/-", new Font("verdana", 10), new
            SolidBrush(Color.Black), new RectangleF(640, 120,
           printDocument1.DefaultPageSettings.PrintableArea.Width,
           printDocument1.DefaultPageSettings.PrintableArea.Height));

        }

        //private void rdSingle_CheckedChanged(object sender, EventArgs e)
        //{
        //    panelSingle.Visible = true;
        //    panelMultiple.Visible = false;
        //    button4.Visible = false;
        //    button3.Visible = false;
        //    button2.Visible = true;
        //    button1.Visible = true;
        //    btnBrowse.Visible = false;
        //}

        //private void rdMultiple_CheckedChanged(object sender, EventArgs e)
        //{
        //    panelMultiple.Visible = true;
        //    panelSingle.Visible = false;

        //    button4.Visible = true;
        //    button3.Visible = true;
        //    button2.Visible = false;
        //    button1.Visible = false;
        //    btnBrowse.Visible = true;
        //}

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xls files (*.xls;*.xlsx)|*.XLS;*.XLSX",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true

            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtBrowse.Text = openFileDialog1.FileName;
                //Marshal.FinalReleaseComObject(excelApp);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = txtBrowse.Text;
                string justTheFileName = Path.GetFileNameWithoutExtension(filePath);  // MyFile
                string justTheFileNameWithExtension = Path.GetFileName(filePath);     // MyFile.txt
                string extension = Path.GetExtension(filePath);                // .txt
                string justThefolderItsIn = Path.GetDirectoryName(filePath);
                //string extension = Path.GetExtension(filePath);
                //if (openFileDialog1.ShowDialog() == DialogResult.OK)
                //{
                //     filePath = openFileDialog1.FileName;
                //     extension = Path.GetExtension(filePath);
                //}

                string header = "NO";
                // string conStr, sheetName;
                DataSet ds = ConnectionData(filePath, extension, header);
                AddClass.ds = RemoveEmptyRowsFromDataTable(ds);
                if (printDialog2.ShowDialog() == DialogResult.OK)
                {
                    printDocument2.DefaultPageSettings.Landscape = false;
                    printDocument2.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("PaperA4", 840, 1180);
                    printDocument2.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                    for (int i = 0; i < AddClass.ds.Tables[0].Rows.Count; i++)
                    {
                        rowCnt = i;
                        printDocument2.Print();
                    }
                    txtBrowse.Clear();
                    MessageBox.Show("Multiple Print Done");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Enter Required Formate and values!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtBrowse.Clear();
        }
        DataSet RemoveEmptyRowsFromDataTable(DataSet dt)
        {
            dt.Tables[0].Rows[0].Delete();
            dt.AcceptChanges();
            dt.Tables[0].Columns["F1"].ColumnName = "Date";
            dt.Tables[0].Columns["F2"].ColumnName = "Name";
            dt.Tables[0].Columns["F3"].ColumnName = "Amount";
            dt.Tables[0].Columns["F4"].ColumnName = "AccPayee";


            for (int i = dt.Tables[0].Rows.Count - 1; i >= 0; i--)
            {
                if (dt.Tables[0].Rows[i][1] == DBNull.Value)
                    dt.Tables[0].Rows[i].Delete();
            }
            dt.AcceptChanges();
            return dt;
        }
        private void PrintPage2(object sender, PrintPageEventArgs e)
        {

            if (AddClass.ds.Tables[0].Rows.Count > 0)
            {
                //for (int i = 0; i < AddClass.ds.Tables[0].Rows.Count; i++)
                //{
                if (AddClass.ds.Tables[0].Rows[rowCnt]["Date"].ToString() != "")
                {
                    // string formatted= AddClass.ds.Tables[0].Rows[rowCnt]["Date"].ToString("dd-MM-yyyy");
                    DateTime dt = Convert.ToDateTime(AddClass.ds.Tables[0].Rows[rowCnt]["Date"].ToString());
                    String dateInString = dt.ToString("dd-MM-yyyy");

                    e.Graphics.DrawString(dateInString, new Font("verdana", 10), new
                    SolidBrush(Color.Black), new RectangleF(620, 5,
                    printDocument1.DefaultPageSettings.PrintableArea.Width,
                   printDocument1.DefaultPageSettings.PrintableArea.Height));
                }
                if (AddClass.ds.Tables[0].Rows[rowCnt]["AccPayee"].ToString() != "")
                {
                    e.Graphics.DrawString("A/C PAYEE ONLY", new Font("verdana", 10), new
                    SolidBrush(Color.Black), new RectangleF(300, 220,
                    printDocument1.DefaultPageSettings.PrintableArea.Width,
                    printDocument1.DefaultPageSettings.PrintableArea.Height));
                }
                e.Graphics.DrawString(AddClass.ds.Tables[0].Rows[rowCnt]["Name"].ToString() + "****", new Font("verdana", 10), new
                SolidBrush(Color.Black), new RectangleF(100, 60,
               printDocument1.DefaultPageSettings.PrintableArea.Width,
               printDocument1.DefaultPageSettings.PrintableArea.Height));

                e.Graphics.DrawString(words(Convert.ToDouble(AddClass.ds.Tables[0].Rows[rowCnt]["Amount"]), false), new Font("verdana", 10), new
                SolidBrush(Color.Black), new RectangleF(130, 90,
                printDocument1.DefaultPageSettings.PrintableArea.Width,
                printDocument1.DefaultPageSettings.PrintableArea.Height));

                e.Graphics.DrawString("**" + AddClass.ds.Tables[0].Rows[rowCnt]["Amount"].ToString() + "/-", new Font("verdana", 10), new
                SolidBrush(Color.Black), new RectangleF(640, 120,
               printDocument1.DefaultPageSettings.PrintableArea.Width,
               printDocument1.DefaultPageSettings.PrintableArea.Height));
                // }
            }
        }

        public DataSet ConnectionData(string filePath, string extension, string header)
        {
            string conStr, sheetName;
            DataSet DT = new DataSet();
            conStr = string.Empty;
            // string sheetName = "";

            // DataTable dt = new DataTable();
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {

                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(DT);
                        con.Close();
                    }
                }
            }
            return DT;
        }

        private void rdMultiple_Click(object sender, EventArgs e)
        {
            panelMultiple.Visible = true;
            panelSingle.Visible = false;

            button4.Visible = true;
            button3.Visible = true;
            button2.Visible = false;
            button1.Visible = false;
            btnBrowse.Visible = true;
        }

        private void rdSingle_Click(object sender, EventArgs e)
        {
            panelSingle.Visible = true;
            panelMultiple.Visible = false;
            button4.Visible = false;
            button3.Visible = false;
            button2.Visible = true;
            button1.Visible = true;
            btnBrowse.Visible = false;
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "xls";
            saveFileDialog.Filter = "Excel files (*.xls)|*.xls |All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string execPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var filePath = Path.Combine(execPath, "Book1.xlsx");
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks.Open(filePath);
                book.SaveAs(saveFileDialog.FileName); //Save
                book.Close();
            }

        }
    }
}
