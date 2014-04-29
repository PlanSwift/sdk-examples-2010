using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using pswift = PlanSwift9;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public void itmtoword(pswift.IItem itm, Word.Table tbl, int rowidx, string tp)
        {
            if (rowidx > 20)
            {
                tbl.Rows.Add(System.Reflection.Missing.Value);
            }
            if (tp == "Digitizer")
            {
                tbl.Cell(rowidx, 1).Range.Font.Bold = 1;
                tbl.Cell(rowidx, 1).Range.Font.Italic = 0;
            }
            else
            {
                tbl.Cell(rowidx, 1).Range.Font.Bold = 0;
                tbl.Cell(rowidx, 1).Range.Font.Italic = 1;
            }
            tbl.Cell(rowidx, 1).Range.Text = itm.Name;
            tbl.Cell(rowidx, 2).Range.Text = itm.GetPropertyResultAsString("Qty","");
            tbl.Cell(rowidx, 3).Range.Text = itm.GetProperty("Qty").Units;
            tbl.Cell(rowidx, 4).Range.Text = itm.GetPropertyResultAsString("Price Each","");
            tbl.Cell(rowidx, 5).Range.Text = itm.GetPropertyResultAsString("Price Total","");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form2 RForm = new Form2();
            RForm.ShowDialog();
            if (RForm.DialogResult != DialogResult.OK){
                RForm.Dispose();
     
                return;
            }
            Object template = Directory.GetCurrentDirectory() + "\\Includes\\Word_Template.dotx";
            Console.WriteLine(template.ToString());
            Object oMissing = System.Reflection.Missing.Value;
            Object isvisible = true;
            Word._Application oWord = new Word.Application();
            try
            {
                Word._Document oDoc = oWord.Documents.Add(template, oMissing, oMissing, isvisible);
                Word.Table otable = oDoc.Tables[2];
                String ReportType = RForm.ReportCBX.Text;
                if (ReportType == "Digitizer Items Only")
                {
                    int rowidx = 1;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        Boolean isDigitizer = false;
                        if (itm.GetPropertyResultAsBoolean("isArea", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isLinear", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isSegment", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isCount", false) == true)
                        {
                            isDigitizer = true;
                        }
                        
                        if (isDigitizer) {
                            rowidx++;
                            itmtoword(itm, otable, rowidx, "Digitizer");


                        }
                        this.progressBar1.Value = idx;      
                    }
                }
                if (ReportType == "Parts Only")
                {
                    int rowidx = 1;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == true) 
                        {
                            rowidx++;
                             itmtoword(itm,otable,rowidx,"Part");
                        }
                        this.progressBar1.Value = idx;
                    }
                }
                if (ReportType == "Digitizer Items w/Parts")
                {
                    int rowidx = 1;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                      
                            rowidx++;
                            if (itm.GetPropertyResultAsBoolean("isPart", false) == true)
                            {
                                itmtoword(itm, otable, rowidx, "Part");
                            }
                            else
                            {
                                itmtoword(itm, otable, rowidx, "Digitizer");
                            }

                            this.progressBar1.Value = idx;
                    }
                }
                
            }
            finally
            {
                this.progressBar1.Value = 0;
                oWord.Visible = true;
                oWord = null;
                RForm.Dispose();

            }


        }

        public void loadTakeoffItems(pswift.IItem itm, List<string> lst)
        {
            for (int idx = 0; idx <= itm.ChildCount() - 1; idx++)
            {
                pswift.IItem citm = itm[idx];
                if (citm.GetProperty("Type").ResultAsString() == "Folder")
                {
                    loadTakeoffItems(citm,lst);
                }
                else
                {
                    Boolean isItem = citm.GetPropertyResultAsBoolean("IsItem", false);
                    if (isItem)
                    {
                        lst.Add(citm.GUID());
                    }
                    if (citm.ChildCount() != 0)
                    {
                        loadTakeoffItems(citm, lst);
                    }
                    
                }

            } 
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            TGlobal.ps = new pswift.PlanSwift();
            string tkpath = TGlobal.ps.Root().FullPath() + "\\Job\\TakeOff";
            pswift.IItem takeoff = TGlobal.ps.GetItem(tkpath);
            TGlobal.tlst = new List<string>();
           
            loadTakeoffItems(takeoff,TGlobal.tlst);
            this.progressBar1.Minimum = 0;
            this.progressBar1.Maximum = TGlobal.tlst.Count - 1;
            
        }
      

        class TGlobal
        {
            public static pswift.IPlanSwift ps;
            public static List<string> tlst;
        }
        public void itmtoexcel(pswift.IItem itm, int rowidx, Excel._Worksheet osheet)
        {
            if (rowidx > 36)
            {
                osheet.Cells[rowidx, 1].EntireRow.Insert(System.Reflection.Missing.Value);
            }
            osheet.Cells[rowidx, 1].value = itm.GetPropertyResultAsString("Qty", "");
            osheet.Cells[rowidx, 2].value = itm.GetProperty("Qty").Units;
            osheet.Cells[rowidx, 3].value = itm.GetPropertyResultAsString("Item #","");
            osheet.Cells[rowidx, 4].value = itm.Name;
            osheet.Cells[rowidx,5].value = itm.GetPropertyResultAsString("Price Each","");
            osheet.Cells[rowidx, 6].value = itm.GetPropertyResultAsString("Price Total", "");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 RForm = new Form2();
            RForm.ShowDialog();
            if (RForm.DialogResult != DialogResult.OK)
            {
                RForm.Dispose();

                return;
            }
            Object template = Directory.GetCurrentDirectory() + "\\Includes\\Excel_Template.XLT";
            Console.WriteLine(template.ToString());
            Object oMissing = System.Reflection.Missing.Value;
            Object isvisible = true;
            Excel._Application oExcel = new Excel.Application();
            try
            {
                Excel._Workbook oBook = oExcel.Workbooks.Add(template);
                Excel._Worksheet oSheet = oExcel.Worksheets[1];
                String ReportType = RForm.ReportCBX.Text;
                if (ReportType == "Digitizer Items Only")
                {
                    int rowidx = 17;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        Boolean isDigitizer = false;
                        if (itm.GetPropertyResultAsBoolean("isArea", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isLinear", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isSegment", false) == true)
                        {
                            isDigitizer = true;
                        }
                        if (itm.GetPropertyResultAsBoolean("isCount", false) == true)
                        {
                            isDigitizer = true;
                        }

                        if (isDigitizer)
                        {
                            rowidx++;
                            itmtoexcel(itm, rowidx, oSheet);


                        }
                        this.progressBar1.Value = idx;
                    }
                }
                if (ReportType == "Parts Only")
                {
                    int rowidx = 17;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == true)
                        {
                            rowidx++;
                            itmtoexcel(itm, rowidx,oSheet);
                        }
                        this.progressBar1.Value = idx;
                    }
                }
                if (ReportType == "Digitizer Items w/Parts")
                {
                    int rowidx = 17;
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);

                        rowidx++;
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == true)
                        {
                            itmtoexcel(itm, rowidx, oSheet);
                        }
                        else
                        {
                            itmtoexcel(itm, rowidx, oSheet);
                        }

                        this.progressBar1.Value = idx;
                    }
                }

            }
            finally
            {
                this.progressBar1.Value = 0;
                oExcel.Visible = true;
                oExcel = null;
                RForm.Dispose();

            }
        }
        public string createMailBody(string RT)
        {
            char quo = (char)34;
             string body = "<body>";
                body = body + "<h1>Report Type: " + RT + "</h1>" + "\r\n";
                body = body + "<table style="+ quo +"width:auto; font-weight:Bold; font-family:Tahoma; font-size:14px;"+ quo +">" + "\r\n";
                body = body + "<tr>" + "\r\n";
                body = body + "<td style="+ quo +"width:200px; font-weight:Bold; font-family:Tahoma; font-size:14px;"+ quo +">";
                body = body + "Name";
                body = body + "</td>" + "\r\n";
                body = body + "<td style="+ quo +"width:100px; font-weight:Bold; font-family:Tahoma; font-size:14px;"+ quo +">";
                body = body + "Qty";
                body = body + "</td>" + "\r\n";
                body = body + "<td style="+ quo +"width:100px; font-weight:Bold; font-family:Tahoma; font-size:14px;"+ quo +">";
                body = body + "Price Each";
                body = body + "</td>" + "\r\n";
                body = body + "<td style="+ quo +"width:100px; font-weight:Bold; font-family:Tahoma;font-size:14px;"+ quo +">";
                body = body + "Price Total";
                body = body + "</td>" + "\r\n";
                body = body + "</tr>" + "\r\n";
                if (RT == "Digitizer Items Only")
                {
                for (int idx = 0; idx <=  TGlobal.tlst.Count -1; idx++){
                    pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                    if (itm.GetPropertyResultAsBoolean("isPart",false) == false) {
                        body = body + "<tr><td style="+ quo +"font-size:12px; font-weight:normal;"+ quo +">" + itm.Name + "</td>";
                        body = body + "<td style="+ quo +"font-size:12px; font-weight:normal;"+ quo +">" + itm.GetPropertyResultAsString("Qty","") + "</td>";
                        body = body + "<td style="+ quo +"font-size:12px; font-weight:normal;"+ quo +">" + itm.GetPropertyResultAsString("Price Each","") + "</td>";
                        body = body + "<td style=" + quo + "font-size:12px; font-weight:normal;" + quo + ">" + itm.GetPropertyResultAsString("Price Total","") + "</td>";
                        body = body + "</tr>";
                    }

                }
                

                }
                if (RT == "Parts Only")
                {
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == true)
                        {
                            body = body + "<tr><td style=" + quo + "font-size:12px; font-weight:normal;" + quo + ">" + itm.Name + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal;" + quo + ">" + itm.GetPropertyResultAsString("Qty", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal;" + quo + ">" + itm.GetPropertyResultAsString("Price Each", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal;" + quo + ">" + itm.GetPropertyResultAsString("Price Total", "") + "</td>";
                            body = body + "</tr>";
                        }

                    }


                }
                if (RT == "Digitizer Items w/Parts")
                {
                    for (int idx = 0; idx <= TGlobal.tlst.Count - 1; idx++)
                    {
                        pswift.IItem itm = TGlobal.ps.GetItem(TGlobal.tlst[idx]);
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == false)
                        {
                            body = body + "<tr><td style=" + quo + "font-size:12px; font-weight:normal; color:#007dc3;" + quo + ">" + itm.Name + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#007dc3;" + quo + ">" + itm.GetPropertyResultAsString("Qty", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#007dc3;" + quo + ">" + itm.GetPropertyResultAsString("Price Each", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#007dc3;" + quo + ">" + itm.GetPropertyResultAsString("Price Total", "") + "</td>";
                            body = body + "</tr>";
                        }
                        if (itm.GetPropertyResultAsBoolean("isPart", false) == true)
                        {
                            body = body + "<tr><td style=" + quo + "font-size:12px; font-weight:normal; color:#0000FF;" + quo + ">" + itm.Name + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#0000FF;" + quo + ">" + itm.GetPropertyResultAsString("Qty", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#0000FF;" + quo + ">" + itm.GetPropertyResultAsString("Price Each", "") + "</td>";
                            body = body + "<td style=" + quo + "font-size:12px; font-weight:normal; color:#0000FF;" + quo + ">" + itm.GetPropertyResultAsString("Price Total", "") + "</td>";
                            body = body + "</tr>";
                        }

                    }


                }
            return body;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Form2 RForm = new Form2();
            RForm.ShowDialog();
            if (RForm.DialogResult != DialogResult.OK)
            {
                RForm.Dispose();

                return;
            }
            try
            {
                Object display = true;
                Outlook._Application olook = new Outlook.Application();
                Outlook._MailItem omail = olook.CreateItem(0);
                omail.HTMLBody = createMailBody(RForm.ReportCBX.Text);
                omail.Display(display);
            }
            finally
            {
                
                RForm.Dispose();
            }
        }
    }
   
}
