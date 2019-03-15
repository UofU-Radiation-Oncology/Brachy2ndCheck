﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;




namespace brachy2ndcheck
{

    public class GrowLabel : System.Windows.Forms.Label
    {
        private bool mGrowing;
        public GrowLabel()
        {
            this.AutoSize = false;
        }
        private void resizeLabel()
        {
            if (mGrowing) return;
            try
            {
                mGrowing = true;
                Size sz = new Size(this.Width, Int32.MaxValue);
                sz = TextRenderer.MeasureText(this.Text, this.Font, sz, TextFormatFlags.WordBreak);
                this.Height = sz.Height;
            }
            finally
            {
                mGrowing = false;
            }
        }
        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            resizeLabel();
        }
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            resizeLabel();
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            resizeLabel();
        }
    }


    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 

        // ByteSearch is a function which will find a small byte array in a larger byte array and return its index location
        public static int ByteSearch(byte[] searchIn, byte[] searchBytes, int start = 0)
        {
            int found = -1;
            bool matched = false;
            //only look at this if we have a populated search array and search bytes with a sensible start
            if (searchIn.Length > 0 && searchBytes.Length > 0 && start <= (searchIn.Length - searchBytes.Length) && searchIn.Length >= searchBytes.Length)
            {
                //iterate through the array to be searched
                for (int i = start; i <= searchIn.Length - searchBytes.Length; i++)
                {
                    //if the start bytes match we will start comparing all other bytes
                    if (searchIn[i] == searchBytes[0])
                    {
                        if (searchIn.Length > 1)
                        {
                            //multiple bytes to be searched we have to compare byte by byte
                            matched = true;
                            for (int y = 1; y <= searchBytes.Length - 1; y++)
                            {
                                if (searchIn[i + y] != searchBytes[y])
                                {
                                    matched = false;
                                    break;
                                }
                            }
                            //everything matched up
                            if (matched)
                            {
                                found = i;
                                break;
                            }

                        }
                        else
                        {
                            //search byte is only one bit nothing else to do
                            found = i;
                            break; //stop the loop
                        }

                    }
                }

            }
            return found;
        }

        //InsertByte is a function which inserts a small byte array into a larger byte array
        public static byte[] InsertByte(byte[] orig, byte[] insertion, int startloc, int endloc)
        {
            byte[] newb = new byte[orig.Length + startloc - endloc + 1 + insertion.Length];
            Array.Copy(orig, 0, newb, 0, startloc + 1);
            Array.Copy(insertion, 0, newb, startloc + 1, insertion.Length);
            Array.Copy(orig, endloc, newb, newb.Length - orig.Length + endloc, orig.Length - endloc);

            return newb;

        }


        public static string stringTag(byte[] tag, byte[] file, int nth = 1, int startpos = 0)
        {
            int tagloc = ByteSearch(file, tag, startpos);
            //Console.WriteLine(tagloc.ToString());
            for (int i = 1; i < nth; i++)
            {
                tagloc = ByteSearch(file, tag, tagloc + 1);
                //Console.WriteLine(tagloc.ToString());
            }
            byte[] length = new byte[2];
            Array.Copy(file, tagloc + 6, length, 0, 2);
            short len = BitConverter.ToInt16(length, 0);
            string tagvalue = System.Text.Encoding.Default.GetString(file, tagloc + 8, len);
            return tagvalue;
        }
        public static int CountOccurences(byte[] searchIn, byte[] searchFor)
        {
            //Console.WriteLine("test");
            int count = 0;
            int loc = 0;
            while (loc > -1)
            {
                loc = ByteSearch(searchIn, searchFor, loc + 1);
                if (loc > -1)
                {
                    count++;
                }
            }
            return count;
        }

        public static double[,] Catheters(byte[] searchIn)
        {
            byte[] cntrtypedcm = { 0x06, 0x30, 0x42, 0x00 };
            byte[] ctrlpointsdcm = { 0x06, 0x30, 0x50, 0x00 };
            byte[] endofcath = { 0x06, 0x30, 0x80, 0x00 };
            string result = "";
            int temp = 0;
            List<int> cathloc = new List<int>();
            string cath = "OPEN_NONPLANAR";
            temp = ByteSearch(searchIn, cntrtypedcm, temp);
            while (temp > -1) {
                result = stringTag(cntrtypedcm, searchIn, 1, temp-10);
                if (result == cath)
                {
                    cathloc.Add(temp);
                }
                temp = ByteSearch(searchIn, cntrtypedcm, temp+1);
            }

            double[,] points = new double[cathloc.Count, 6];


            for (int i = 0; i < cathloc.Count; i++)
            {
                result = stringTag(ctrlpointsdcm, searchIn, 1, cathloc[i]);
                string[] results = result.Split(new char[] { '\\' });
                for (int j=0; j<6; j++)
                {
                    points[i, j] = double.Parse(results[j]);
                }
            }

            double A = Math.Abs(points[0, 2] - points[1, 2]);
            double B = Math.Abs(points[0, 2] - points[2, 2]);
            double C = Math.Abs(points[1, 2] - points[2, 2]);
            double[,] tno = points;
            if (B < C && B < A)
            {
                for (int i = 0; i < 6; i++)
                {
                    tno[2, i] = points[1, i];
                    tno[1, i] = points[2, i];
                }
            }
            if (C < B && C < A)
            {
                for (int i = 0; i < 6; i++)
                {
                    tno[2, i] = points[0, i];
                    tno[0, i] = points[2, i];
                }
            }
            points = tno;
            if (points[0, 0] > points[1, 0])
            {
                for (int i = 0; i < 6; i++)
                {
                    tno[1, i] = points[0, i];
                    tno[0, i] = points[1, i];
                }
            }
            return points;
        }

        public static double[,] TandemArray(byte[] searchIn, int num)
        {
            byte[] cathnumdcm = { 0x0A, 0x30, 0x82, 0x02 };
            byte[] ctrlpointsdcm = { 0x0A, 0x30, 0x10, 0x01 };
            byte[] dcm3D = { 0x0A, 0x30, 0xD4, 0x02};
            int tandemloc = 0;
            string temp = "";
            for (int i = 1; i < num+1; i++)
            {
                tandemloc = ByteSearch(searchIn, cathnumdcm, tandemloc +1);
            }
            string ctrlpoints = stringTag(ctrlpointsdcm, searchIn, 1, tandemloc - 50);
            int ctrlpnts = int.Parse(ctrlpoints) / 2;
            Console.WriteLine(ctrlpnts);
            double[,] points = new double[ctrlpnts, 3];

            
            for (int i=0; i<ctrlpnts;i++)
            {
                temp = stringTag(dcm3D, searchIn, 2*i + 1, tandemloc);
                points[i, 0] = double.Parse(temp.Substring(0, temp.IndexOf("\\")));
                points[i, 1] = double.Parse(temp.Substring(temp.IndexOf("\\") + 1, temp.IndexOf("\\", temp.IndexOf("\\") + 1) - temp.IndexOf("\\") - 1));
                points[i, 2] = double.Parse(temp.Substring(temp.IndexOf("\\", temp.IndexOf("\\") + 1) + 1, temp.Length - temp.IndexOf("\\", temp.IndexOf("\\") + 1) - 1));

            }
            return points;
        }

        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 500,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };
                System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label() { Left = 50, Top = 20, Text = text };
                textLabel.AutoSize = true;
                System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Left = 50, Top = 50, Width = 400 };
                System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }

        [STAThread]
        static void Main()
        {
            //run the windows program to fix dicom, additional code in form1.cs
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());
        }
       

    }
}
