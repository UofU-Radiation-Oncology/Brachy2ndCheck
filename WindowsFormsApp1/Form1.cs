using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp;

namespace brachy2ndcheck
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }



        private void Form1_Load(object sender, EventArgs e)
        {


            //Prompt user to select the folder containing the DCM files which need their iso adjusted
            openfile.ShowDialog();
            byte[] filein = File.ReadAllBytes(openfile.FileName);

            ActiveForm.Show();
            textBoxTop.AppendText("Date & Time: " + DateTime.Now);
            textBox1.AppendText("DICOM Information" + Environment.NewLine + Environment.NewLine);
            textBox2.AppendText("Comparison Information" + Environment.NewLine + "(User Entered or Lookup)" + Environment.NewLine);
            textBox3.AppendText("Matched?" + Environment.NewLine + Environment.NewLine);
            int pathloc = openfile.FileName.LastIndexOf("\\");
            string path = openfile.FileName.Remove(pathloc+1);
            //textBox1.AppendText("FileName: " + path);

            byte[] ptnamedcm = { 0x10, 0x00, 0x10, 0x00 };
            byte[] ptnumdcm = { 0x10, 0x00, 0x20, 0x00 };
            byte[] ctstudyiddcm = { 0x20, 0x00, 0x10, 0x00 };
            byte[] studydescdcm = { 0x08, 0x00, 0x30, 0x10 };
            byte[] seriesdescdcm = { 0x08, 0x00, 0x3E, 0x10 };
            byte[] rtplannamedcm = { 0x0A, 0x30, 0x03, 0x00 };
            byte[] rtplandescdcm = { 0x0A, 0x30, 0x04, 0x00 };
            byte[] coorddcm = { 0x0A, 0x30, 0x18, 0x00 };
            byte[] pntnamedcm = { 0x0A, 0x30, 0x16, 0x00 };
            byte[] dwelldcm = { 0x0A, 0x30, 0xD4, 0x02 };
            byte[] channeldcm = { 0x0A, 0x30, 0x82, 0x02 };
            byte[] frameofreferenceuid = { 0x08, 0x00, 0x55, 0x11 };
            byte[] patientpositiondcm = { 0x18, 0x00, 0x00, 0x51 };
            byte[] rxdosedcm = { 0x0A, 0x30, 0xA4, 0x00 };
            byte[] refdistdcm = { 0x0A, 0x30, 0x84, 0x02 };
            byte[] ContourData = { 0x06, 0x30, 0x50, 0x00 };
            byte[] sourcestrengthdcm = { 0x0A, 0x30, 0x2A, 0x02 };
            
            int channeltotal = Program.CountOccurences(filein, channeldcm);
            
            string ptname = Program.stringTag(ptnamedcm, filein);
            ptname = ptname.Replace("^^^"," ");
            ptname = ptname.Replace("^", ", ");
            string ptnum = Program.stringTag(ptnumdcm, filein);
            string ctstudyid = Program.stringTag(ctstudyiddcm, filein);
            string studydesc = Program.stringTag(studydescdcm, filein);
            string seriesdesc = Program.stringTag(seriesdescdcm, filein);
            string rtplanname = Program.stringTag(rtplannamedcm, filein);
            string rtplandesc = Program.stringTag(rtplandescdcm, filein);
            string rxdose = Program.stringTag(rxdosedcm, filein);
            string patientposition = "";
            string refdiststr1 = Program.stringTag(refdistdcm, filein, 1);
            string refdiststr2 = Program.stringTag(refdistdcm, filein, 2);
            string refdiststr3 = Program.stringTag(refdistdcm, filein, 3);

            double rxdosenum = double.Parse(rxdose);
            rxdosenum = Math.Round(rxdosenum * 100);

            double refdist1 = double.Parse(refdiststr1);
            double refdist2 = double.Parse(refdiststr2);
            double refdist3 = double.Parse(refdiststr3);
            string[] refdist = new string[3];
            refdist[0] = Math.Round(refdist1).ToString();
            refdist[1] = Math.Round(refdist2).ToString();
            refdist[2] = Math.Round(refdist3).ToString();

            double[,] applicators = Program.ApplicatorPoints(filein);
            double[,] firstdwells = Program.FirstDwellPosition(filein);


            string frameofreference = Program.stringTag(frameofreferenceuid, filein, 3);
            bool CT = File.Exists(path + "CT" + frameofreference + ".dcm");
            if (CT)
            {
                byte[] CTset = File.ReadAllBytes(path + "CT" + frameofreference + ".dcm");
                patientposition = Program.stringTag(patientpositiondcm, CTset);
            }
            if (!CT)
            {
                patientposition = "To verify patient orientation please export associated image sets with the RTPlan";
            }

            textBox1.AppendText("Patient Name: \t" + ptname + Environment.NewLine);
            textBox2.AppendText(Environment.NewLine);
            textBox3.AppendText(Environment.NewLine);

            string usermrn = Program.Prompt.ShowDialog("Enter the Patient MRN.", "2nd Check Input");

            textBox2.AppendText("MRN: " + usermrn + Environment.NewLine + Environment.NewLine);
            if (usermrn == ptnum)
            {
                textBox3.AppendText("Match" + Environment.NewLine + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine + Environment.NewLine);
            }
            textBox1.AppendText("MRN: \t \t" + ptnum + Environment.NewLine + Environment.NewLine);
            string userCTnumb = Program.Prompt.ShowDialog("Enter the CT Exam number.", "2nd Check Input");

            double userCTnum = double.Parse(userCTnumb);
            double ctstudyidnum = double.Parse(ctstudyid);

            textBox2.AppendText("CT Exam #: " + userCTnum + Environment.NewLine + Environment.NewLine);
            if (userCTnum == ctstudyidnum)
            {
                textBox3.AppendText("Match" + Environment.NewLine + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine + Environment.NewLine);
            }
            textBox1.AppendText("CT Exam #: \t" + ctstudyid + Environment.NewLine + Environment.NewLine);
            textBox1.AppendText("Patient Orientation: \t" + patientposition + Environment.NewLine + Environment.NewLine);
            textBox1.AppendText("Study Description: \t" + studydesc + Environment.NewLine);
            textBox1.AppendText("Series Description: \t" + seriesdesc + Environment.NewLine + Environment.NewLine);
            textBox1.AppendText("RTPlan Name: \t" + rtplanname + Environment.NewLine);
            textBox1.AppendText("RTPlan Description: " + rtplandesc + Environment.NewLine + Environment.NewLine);

            for (int i=0; i<8; i++)
            {
                textBox2.AppendText(Environment.NewLine);
                textBox3.AppendText(Environment.NewLine);
            }

            string userRx = Program.Prompt.ShowDialog("Please enter the prescribed dose per fraction in cGy.", "2nd Check Input");

            textBox1.AppendText("Rx Dose: " + rxdosenum.ToString() + Environment.NewLine + Environment.NewLine);
            textBox2.AppendText("Rx Dose: " + userRx + Environment.NewLine + Environment.NewLine);
            if (rxdosenum.ToString()  == userRx)
            {
                textBox3.AppendText("Match" + Environment.NewLine + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine + Environment.NewLine);
            }

            int n = applicators.Length;
            n = n / 6;
            double[] expectoffset = new double[3];
            expectoffset[0] = -6;
            expectoffset[1] = -6;
            expectoffset[2] = -5;
            string expectrefdist = "1300";
            for (int i = 0; i < n; i++)
            {
                double temp = Math.Sqrt(Math.Pow(firstdwells[i, 0] - applicators[i, 3], 2) + Math.Pow(firstdwells[i, 1] - applicators[i, 4], 2) + Math.Pow(firstdwells[i, 2] - applicators[i, 5], 2)) - Math.Sqrt(Math.Pow(applicators[i, 0] - applicators[i, 3], 2) + Math.Pow(applicators[i, 1] - applicators[i, 4], 2) + Math.Pow(applicators[i, 2] - applicators[i, 5], 2));
                int num = i + 1;
                textBox1.AppendText("Reference Distance for Applicator " + num.ToString() + ":  " + refdist[i] + Environment.NewLine);
                textBox1.AppendText("Offset for Applicator " + num.ToString() + ":  " + Math.Round(temp).ToString() + Environment.NewLine);
                textBox2.AppendText(expectrefdist + "mm" + Environment.NewLine);
                if (refdist[i] == expectrefdist)
                {
                    textBox3.AppendText("Match" + Environment.NewLine);
                } else
                {
                    textBox3.AppendText("ERROR" + Environment.NewLine);
                }
                textBox2.AppendText(expectoffset[i] + "mm" + Environment.NewLine);
                if (Math.Abs(temp - expectoffset[i]) < 0.01 )
                {
                    textBox3.AppendText("Match" + Environment.NewLine);
                }
                else
                {
                    textBox3.AppendText("ERROR" + Environment.NewLine);
                }
            }

            textBox1.AppendText(Environment.NewLine);
            textBox2.AppendText(Environment.NewLine);
            textBox3.AppendText(Environment.NewLine);



            string excelpath = "X:\\RadOnc\\Physics\\BRACHYTHERAPY\\HDR\\HDR QA\\Source exchange\\";
            excelpath = excelpath + DateTime.Now.Year.ToString();
            excelpath = excelpath + "_SourceExchangeQA";

            if (!Directory.Exists(excelpath))
            {
                excelpath = "X:\\RadOnc\\Physics\\BRACHYTHERAPY\\HDR\\HDR QA\\Source exchange\\";
                excelpath = excelpath + DateTime.Now.AddYears(-1).Year.ToString();
                excelpath = excelpath + "_SourceExchangeQA";
            }

            string[,] points = new string[4, 4];
            int pnttotal = Program.CountOccurences(filein, coorddcm);
            int dwelltotal = Program.CountOccurences(filein, dwelldcm);
            //textBox1.AppendText("count occurence: " + pnttotal + Environment.NewLine + Environment.NewLine);

            for (int j = 1; j < pnttotal+1; j++)
            {
                string pntname = Program.stringTag(pntnamedcm, filein, j);
                string coord = Program.stringTag(coorddcm, filein, j);
                //textBox1.AppendText("test: " + pntname + " " + coord + Environment.NewLine);
                int foundA = pntname.IndexOf("A");
                int founda = pntname.IndexOf("a");
                int foundB = pntname.IndexOf("B");
                int foundb = pntname.IndexOf("b");
                int foundR = pntname.IndexOf("Rt");
                int foundr = pntname.IndexOf("rt");
                int foundL = pntname.IndexOf("Lt");
                int foundl = pntname.IndexOf("lt");
                //textBox1.AppendText("A: " + foundA + "  B: " + foundB + "  r: " + foundr + "  l: " + foundl+Environment.NewLine);
                if (foundA > -1 || founda > -1) {
                    if (foundr > -1 || foundR > -1) {
                        points[1, 0] = pntname;
                        points[1, 1] = coord.Substring(0, coord.IndexOf("\\"));
                        points[1, 2] = coord.Substring(coord.IndexOf("\\") + 1, coord.IndexOf("\\", coord.IndexOf("\\") + 1) - coord.IndexOf("\\")-1);
                        points[1, 3] = coord.Substring(coord.IndexOf("\\", coord.IndexOf("\\") + 1) + 1, coord.Length - coord.IndexOf("\\", coord.IndexOf("\\") + 1)-1);
                    }
                    if (foundl > -1 || foundL > -1) {
                        points[0, 0] = pntname;
                        points[0, 1] = coord.Substring(0, coord.IndexOf("\\"));
                        points[0, 2] = coord.Substring(coord.IndexOf("\\") + 1, coord.IndexOf("\\", coord.IndexOf("\\") + 1) - coord.IndexOf("\\") - 1);
                        points[0, 3] = coord.Substring(coord.IndexOf("\\", coord.IndexOf("\\") + 1) + 1, coord.Length - coord.IndexOf("\\", coord.IndexOf("\\") + 1) - 1);
                    }
                }
                if (foundB > -1 || foundb > -1) {
                    if (foundr > -1 || foundR > -1) {
                        points[3, 0] = pntname;
                        points[3, 1] = coord.Substring(0, coord.IndexOf("\\"));
                        points[3, 2] = coord.Substring(coord.IndexOf("\\") + 1, coord.IndexOf("\\", coord.IndexOf("\\") + 1) - coord.IndexOf("\\") - 1);
                        points[3, 3] = coord.Substring(coord.IndexOf("\\", coord.IndexOf("\\") + 1) + 1, coord.Length - coord.IndexOf("\\", coord.IndexOf("\\") + 1) - 1);
                    }
                    if (foundl > -1 || foundL > -1) {
                        points[2, 0] = pntname;
                        points[2, 1] = coord.Substring(0, coord.IndexOf("\\"));
                        points[2, 2] = coord.Substring(coord.IndexOf("\\") + 1, coord.IndexOf("\\", coord.IndexOf("\\") + 1) - coord.IndexOf("\\") - 1);
                        points[2, 3] = coord.Substring(coord.IndexOf("\\", coord.IndexOf("\\") + 1) + 1, coord.Length - coord.IndexOf("\\", coord.IndexOf("\\") + 1) - 1);
                    }
                }
                

            }

            
            double[,] tandem = Program.TandemArray(filein);
            int tandemlen = tandem.GetLength(0);
            double x = 0;
            double y = 0;
            double z = 0;
            for (int i = 0; i < tandemlen; i++)
            {
                x = tandem[i, 0] + x;
                y = tandem[i, 1] + y;
                z = tandem[i, 2] + z;
            }
            x = x / tandemlen;
            y = y / tandemlen;
            z = z / tandemlen;
            for (int i = 0; i < tandemlen; i++)
            {
                tandem[i, 0] = tandem[i, 0] - x;
                tandem[i, 1] = tandem[i, 1] - y;
                tandem[i, 2] = tandem[i, 2] - z;
            }
            Matrix<double> tndm = DenseMatrix.OfArray(tandem);
            var svd = tndm.Svd(true);
            //            double[,] refframe = MathNet.Numerics.LinearAlgebra.Factorization.Svd(DenseMatrix.OfArray(tandem)) ;

            //textBox1.AppendText(svd.U.ToString());
            //textBox1.AppendText(svd.VT.ToString());
            //textBox1.AppendText(svd.W.ToString());

            string lastdwellcoord = Program.stringTag(dwelldcm, filein, dwelltotal);
            double tandemlr = double.Parse(lastdwellcoord.Substring(0, lastdwellcoord.IndexOf("\\")));
            double tandemap = double.Parse(lastdwellcoord.Substring(lastdwellcoord.IndexOf("\\") + 1, lastdwellcoord.IndexOf("\\", lastdwellcoord.IndexOf("\\") + 1) - lastdwellcoord.IndexOf("\\") - 1));
            double tandemsi = double.Parse(lastdwellcoord.Substring(lastdwellcoord.IndexOf("\\", lastdwellcoord.IndexOf("\\") + 1) + 1, lastdwellcoord.Length - lastdwellcoord.IndexOf("\\", lastdwellcoord.IndexOf("\\") + 1) - 1));

            //textBox1.AppendText(Environment.NewLine + tandemlr + Environment.NewLine + tandemap + Environment.NewLine + tandemsi);

            double Adis = Math.Round(Math.Sqrt(Math.Pow(double.Parse(points[0, 1]) - double.Parse(points[1, 1]), 2) + Math.Pow(double.Parse(points[0, 2]) - double.Parse(points[1, 2]), 2) + Math.Pow(double.Parse(points[0, 3]) - double.Parse(points[1, 3]), 2)), 0);
            double Bdis = Math.Round(Math.Sqrt(Math.Pow(double.Parse(points[2, 1]) - double.Parse(points[3, 1]), 2) + Math.Pow(double.Parse(points[2, 2]) - double.Parse(points[3, 2]), 2) + Math.Pow(double.Parse(points[2, 3]) - double.Parse(points[3, 3]), 2)), 0);
            double Al = Math.Round(double.Parse(points[0, 1]) - tandemlr, 0);
            double Bl = Math.Round(double.Parse(points[2, 1]) - tandemlr, 0);
            double Ar = Math.Round(double.Parse(points[1, 1]) - tandemlr, 0);
            double Br = Math.Round(double.Parse(points[3, 1]) - tandemlr, 0);

            Ar = -Ar;
            Br = -Br;

            textBox1.AppendText("Left Point A is " + Al.ToString() + "mm left of the tandem." + Environment.NewLine);
            textBox1.AppendText("Right Point A is " + Ar.ToString() + "mm right of the tandem." + Environment.NewLine);
            textBox1.AppendText("Left Point B is " + Bl.ToString() + "mm left of the tandem." + Environment.NewLine);
            textBox1.AppendText("Right Point B is " + Br.ToString() + "mm right of the tandem." + Environment.NewLine);
            textBox2.AppendText("20mm" + Environment.NewLine);
            textBox2.AppendText("20mm" + Environment.NewLine);
            textBox2.AppendText("50mm" + Environment.NewLine);
            textBox2.AppendText("50mm" + Environment.NewLine);
            if (Math.Abs(Al - 20) < 2)
            {
                textBox3.AppendText("Match (within 1mm)" + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine);
            }
            if (Math.Abs(Ar - 20) < 2)
            {
                textBox3.AppendText("Match (within 1mm)" + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine);
            }
            if (Math.Abs(Bl - 50) < 4)
            {
                textBox3.AppendText("Match (within 3mm)" + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine);
            }
            if (Math.Abs(Br - 50) < 4)
            {
                textBox3.AppendText("Match (within 3mm)" + Environment.NewLine);
            }
            else
            {
                textBox3.AppendText("ERROR" + Environment.NewLine);
            }

            textBox1.AppendText(Environment.NewLine);
            textBox2.AppendText(Environment.NewLine);
            textBox3.AppendText(Environment.NewLine);

            if (Directory.Exists(excelpath))
            {
                //textBox1.AppendText(Environment.NewLine + excelpath);
                string[] folders = Directory.GetDirectories(excelpath);

                if (!Directory.Exists(folders[folders.Length - 1]))
                {
                    excelpath = "X:\\RadOnc\\Physics\\BRACHYTHERAPY\\HDR\\HDR QA\\Source exchange\\";
                    excelpath = excelpath + DateTime.Now.AddYears(-1).Year.ToString();
                    excelpath = excelpath + "_SourceExchangeQA";
                    folders = Directory.GetDirectories(excelpath);
                }

                excelpath = folders[folders.Length - 1] + "\\" + folders[folders.Length - 1].Substring(folders[folders.Length - 1].Length - 8) + ".xlsx";

                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(excelpath);
                Worksheet ws = wb.Worksheets[3];
                DateTime sourcetime = ws.Cells[14, 9].Value;
                double sourcestrength = ws.Cells[14, 3].Value2;
                if (Math.Abs(sourcestrength - 12) > 3)
                {
                    textBox2.AppendText("Source Exchange Data could not be found." + Environment.NewLine);

                }
                DateTime dtcurrent = DateTime.Now;
                TimeSpan decaytime = dtcurrent - sourcetime;
                double currentstrength = sourcestrength * Math.Exp(decaytime.TotalDays * Math.Log(0.5) / 73.83);
                textBox2.AppendText("Current Source Strength: " + Math.Round(currentstrength, 3).ToString() + "mCi" + Environment.NewLine);

                string plansourcestrength = Program.stringTag(sourcestrengthdcm, filein);
                double planstrength = double.Parse(plansourcestrength)*0.000243;
                textBox1.AppendText("Plan Source Strength: " + Math.Round(planstrength,3).ToString() + "mCi");
                if (Math.Abs(planstrength - currentstrength) < 0.1*planstrength)
                {
                    textBox3.AppendText("Match (within 1% of plan Activity)" + Environment.NewLine);
                }
                else
                {
                    textBox3.AppendText("ERROR (off by >1% of plan Activity)" + Environment.NewLine);
                }

            }
            else
            {
                textBox2.AppendText("Source Exchange Data could not be found." + Environment.NewLine);
            }

            





            if (Math.Abs(Math.Abs(Adis) - 40) > 2)
            {
                textBox1.AppendText(Environment.NewLine + "CHECK YOUR POINT A DISTANCE FROM THE APPLICATOR. IT SHOULD BE 2cm. THE DISTANCE BETWEEN THE TWO POINT A POINTS IS " + Adis.ToString() + "MM" + Environment.NewLine + Environment.NewLine);
            }
            if (Al - Ar < 0 && (patientposition == "HFS" || patientposition == "FFP"))
            {
                textBox1.AppendText(Environment.NewLine + "MAKE SURE THAT YOUR POINT A POINTS ARE ON THE CORRECT SIDE OF THE PATIENT AND THAT YOUR LABELS ARE CORRECT.");
            }
            if (Ar - Al < 0 && (patientposition == "FFS" || patientposition == "HFP"))
            {
                textBox1.AppendText(Environment.NewLine + "MAKE SURE THAT YOUR POINT A POINTS ARE ON THE CORRECT SIDE OF THE PATIENT AND THAT YOUR LABELS ARE CORRECT.");
            }
            if (Math.Abs(Math.Abs(Bdis) - 100) > 2)
            {
                textBox1.AppendText(Environment.NewLine + "CHECK YOUR POINT B DISTANCE FROM THE APPLICATOR. IT SHOULD BE 5cm. THE DISTANCE BETWEEN THE TWO POINT B POINTS IS " + Bdis.ToString() + "MM" + Environment.NewLine + Environment.NewLine);
            }
            if (Bl - Br < 0 && (patientposition == "HFS" || patientposition == "FFP"))
            {
                textBox1.AppendText(Environment.NewLine + "MAKE SURE THAT YOUR POINT B POINTS ARE ON THE CORRECT SIDE OF THE PATIENT AND THAT YOUR LABELS ARE CORRECT.");
            }
            if (Br - Bl < 0 && (patientposition == "FFS" || patientposition == "HFP"))
            {
                textBox1.AppendText(Environment.NewLine + "MAKE SURE THAT YOUR POINT B POINTS ARE ON THE CORRECT SIDE OF THE PATIENT AND THAT YOUR LABELS ARE CORRECT.");
            }

            
            //textBox1.AppendText("A distance: " + Adis.ToString() + Environment.NewLine);
            //textBox1.AppendText("B distance: " + Bdis.ToString() + Environment.NewLine);

           

            

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }

}
