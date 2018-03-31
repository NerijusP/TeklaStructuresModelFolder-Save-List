using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab
using System.Diagnostics;


namespace CD_Last_Saved
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {

            // string direktorija = Application.StartupPath;
            string direktorija = @"C:\TeklaStructuresModels";

            // string pathString = System.IO.Path.Combine(direktorija, ".This_is_multiuser_model");
            // string[] dirs = Directory.GetFiles(direktorija, "*.This_is_multiuser_model");
            //MessageBox.Show(direktorija);


            string aryrasenas = direktorija + "\\CD_Last_Saved.txt";

            if (File.Exists(aryrasenas))
            {
                File.Delete(aryrasenas);
            }


            string dabartinis = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            System.IO.StreamWriter file3 = new StreamWriter(direktorija+ "\\CD_Last_Saved.txt", true);
            file3.Flush();

            foreach (string d in Directory.GetDirectories(direktorija))
            {
                foreach (string f in Directory.GetFiles(d, "save_history.log"))
                {
                    // lstFilesFound.Add(f);
                    // MessageBox.Show(f);

                    string modeliukas = d.Split(new[] { "C:\\TeklaStructuresModels" }, StringSplitOptions.None).Last();

                    string line;
                    System.IO.StreamReader file = new System.IO.StreamReader(f);

                    string ipadresas2 = "";
                    string portas = "";
                    double nrofDays = 0.0;

                    DateTime dabaryra = DateTime.Now;
                   // DateTime myDate;
                    while ((line = file.ReadLine()) != null)
                    {
                        //MessageBox.Show(line);

                        if (line.Contains(" Save "))
                        {
                            var ipadresas = line.Split(new[] { " Save " }, StringSplitOptions.None).Last();
                            string ipadresastarp = ipadresas.Split(new[] { "	*** " }, StringSplitOptions.None).First();
                            ////ipadresas2 = ipadresas.Substring(0, 15);
                            ipadresas2 = ipadresastarp;

                            DateTime myDate = DateTime.ParseExact(ipadresas2, "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            TimeSpan t = dabaryra - myDate;

                            nrofDays = Math.Round(t.TotalDays,0);

                            //var portalas = line.Split(new[] { "," }, StringSplitOptions.None).Last();
                            //portas = portalas.Substring(0, 4);
                            //  MessageBox.Show

                        }

                    }
                    //if (ipadresas2!="192.168.100.241")
                    //{
                        file3.Write("\r\n" + dabartinis + "\t" + modeliukas + "\t" + "\t" + "\t" + ipadresas2+ "\t"+nrofDays, Environment.NewLine);
                        file3.Flush();
                    //}

                }
                //   DirSearch(d);


            }


            MessageBox.Show("Sarasas parengtas");


        }
    }
}
