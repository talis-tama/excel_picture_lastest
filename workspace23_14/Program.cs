using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ClosedXML.Excel;

namespace workspace23_14
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
    public class Form1 : Form
    {
        TextBox textbox1, /*textbox2,*/ textbox3;
        //Label label2;
        //ProgressBar bar;
        Button button;
        OpenFileDialog dialog;
        string dataPath, inputFileName, outputFileName, nameFile;
        string[] dataName;
        int number = 0/*, r, g, bb*/;
        public Form1()
        {
            Width = 350;
            Height = 165;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Label label1 = new Label();
            label1.Location = new Point(10, 10);
            label1.Size = new Size(300, 15);
            label1.Text = "画像ファイルのディレクトリ、ネームファイルを指定してください";
            Controls.Add(label1);
            textbox1 = new TextBox();
            textbox1.Location = new Point(10, 35);
            textbox1.Size = new Size(300, 15);
            Controls.Add(textbox1);
            textbox3 = new TextBox();
            textbox3.Location = new Point(10, 60);
            textbox3.Size = new Size(300, 15);
            textbox3.ReadOnly = true;
            textbox3.Click += new EventHandler(textbox3_Click);
            Controls.Add(textbox3);
            button = new Button();
            button.Location = new Point(10, 85);
            button.Size = new Size(300, 30);
            button.Text = "開始";
            button.Enabled = false;
            button.Click += new EventHandler(button_click);
            Controls.Add(button);
            open();
            /*bar = new ProgressBar();
            bar.Location = new Point(10, 125);
            bar.Size = new Size(300, 25);
            Controls.Add(bar);
            label2 = new Label();
            label2.Location = new Point(10, 160);
            label2.Size = new Size(300, 15);
            Controls.Add(label2);
            bar.Minimum = 0;
            bar.Maximum = 100;
            label2.Text = "0,0,0,0.00%";
            label2.Update();*/
            /*textbox2 = new TextBox();
            textbox2.Location = new Point(10, 125);
            textbox2.Size = new Size(300, 725);
            textbox2.Multiline = true;
            textbox2.ScrollBars = ScrollBars.Vertical;
            textbox2.ReadOnly = true;
            Controls.Add(textbox2);*/
        }
        void button_click(object sender, EventArgs e)
        {
            dataPath = textbox1.Text;
            /*using (var book = new XLWorkbook())
            {
                var sheet1 = book.Worksheets.Add("Sheet1");
                sheet1.Columns().Width = 0.08;
                sheet1.Rows().Height = 4.50;
                book.SaveAs(dataPath + "\\testafter.xlsx");
                book.Dispose();
            }*/
            painting();
        }
        void textbox3_Click(object sender,EventArgs e) { open(); }
        void open()
        {
            int i = 0;
            textbox3.Text = "";
            dialog = new OpenFileDialog();
            dialog.FileName = "";
            dialog.InitialDirectory = @"C:\";
            dialog.Filter = "テキストファイル(*.txt)|*.txt";
            dialog.Title = "ファイルを選択";
            dialog.RestoreDirectory = true;
            dialog.CheckFileExists = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textbox3.Text = dialog.FileName;
                nameFile = dialog.FileName;
                dataName = System.IO.File.ReadAllLines(nameFile, System.Text.Encoding.GetEncoding("shift_jis"));
                button.Enabled = true;
            }
        }
        void painting()
        {
            button.Enabled = false;
            button.Text = "initalizing...";
            button.Update();
            inputFileName = dataPath + number.ToString() + ".png";
            outputFileName = dataPath + dataName[number];
            int wid, hig, r, g, bb, i = 0;
            //string r, g, bb;
            try
            {
                Bitmap img = new Bitmap(inputFileName);
                wid = img.Width;
                hig = img.Height;
            BitmapData data = img.LockBits(new Rectangle(0, 0, wid, hig), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            byte[] buf = new byte[wid * hig * 4];
            Marshal.Copy(data.Scan0, buf, 0, buf.Length);
                using (var book = new XLWorkbook("testafter.xlsx"))
                {
                    button.Text = "processing...";
                    button.Update();
                    var sheet1 = book.Worksheet(1);
                    for (int a = 0; a < hig; a++)
                    {
                        for (int b = 0; b < wid; b++)
                        {
                            var cell = sheet1.Cell(a + 1, b + 1);
                            bb = buf[i++];
                            g = buf[i++];
                            r = buf[i++];
                            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, bb);
                            i++;
                            //textbox2.AppendText(b.ToString() + "," + a.ToString() + "," + number.ToString() + "," + progress.ToString() + "%" + Environment.NewLine);
                            //bar.Value = (int)prog;
                            //label2.Text = b.ToString() + "," + a.ToString() + "," + number.ToString() + "," + prog.ToString() + "%";
                            //label2.Update();
                        }
                    }
                    sheet1.Dispose();
                    book.Dispose();
                    img.Dispose();
                    button.Text = "finalizing...";
                    button.Update();
                    book.SaveAs(outputFileName);
                }
                number++;
                if (File.Exists(dataPath + number.ToString() + ".png")) { painting(); }
            }
            catch (System.ArgumentException) { MessageBox.Show("ファイルパスが異常です", "", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                button.Text = "開始";
                button.Enabled = true;
                button.Update();
            }
        }
    }
}
