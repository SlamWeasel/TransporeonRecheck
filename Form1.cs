using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TransporeonRechnungen
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        string connectionString = @"secret";
        string path = "";
        string IDs;
        string NRs;
        bool hasFile = false;
        int lastUsedRow;

        Dictionary<string, string> translation;
        Dictionary<string, string> tourEigentümer;
        Dictionary<string, string> translationRef;
        Dictionary<string, string> tourEigentümerRef;
        Excel.Workbook file;
        Excel.Worksheet page;
        Excel.Range rang;
        string p;

        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length == 1)
                Console.WriteLine(files[0].Split('.')[files[0].Split('.').Length - 1]);

            Control c = this.label1;

            if (files[0].Split('.')[files[0].Split('.').Length - 1] == "xls" ||
                files[0].Split('.')[files[0].Split('.').Length - 1] == "xlsx" ||
                files[0].Split('.')[files[0].Split('.').Length - 1] == "xlsm")
                c.Text = "Datei wurde abgelegt";
            else return;

            path = files[0];
            hasFile = true;
        }

        private Excel.Workbook readXLS(string path)
        {
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");

                return null;
            }

            return xlApp.Workbooks.Open(path);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Text = "Exceldaten werden ausgelesen";
            #region Excel Daten Auslesen
            if (!hasFile)
                return;

            file = readXLS(path);
            page = file.Sheets[1];
            rang = page.UsedRange;
            xlApp.DisplayNoteIndicator = false;
            xlApp.DisplayInfoWindow = false;

            lastUsedRow = page.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            List<string> TransporeonIDs = new List<string>();
            List<string> Transportnummern = new List<string>();
            for (int i = 2; i <= lastUsedRow; i++)
            {

                ///
                /// Untersuchen der Referenznummern!!!
                ///

                string valID = ((Excel.Range)page.Cells[i, 4]).Value2.ToString();
                TransporeonIDs.Add("\'" + valID + "\'");

                string cell = ((Excel.Range)page.Cells[i, 3]).Value2.ToString();

                if (!cell.Contains("+"))
                {
                    string valRef = cell;

                    if (valRef.Contains("."))
                    {
                        valRef = valRef[0] == '.' ? valRef.Substring(1) : valRef;
                        valRef = valRef.Trim('.');
                        valRef = valRef[valRef.Length-2] == '.' || valRef[valRef.Length - 3] == '.' || valRef[valRef.Length - 4] == '.' ? valRef.Split('.')[0] : valRef;
                    }
                    if(valRef.Contains(" "))
                    {
                        valRef = valRef.StartsWith("CMR ") ? valRef.Substring(4) : valRef;
                        valRef = valRef.StartsWith("LS ") || valRef.StartsWith("RG ") ? valRef.Substring(3) : valRef;
                        valRef = valRef.Trim(' ');
                        valRef = valRef.Contains(" ") ? valRef.Split(' ')[0] : valRef;
                    }

                    if (valRef.Length > 4)
                    {
                        Console.WriteLine(valRef);
                        Transportnummern.Add(valRef);
                    }
                }
                else
                {
                    string valRef = cell.Split(" + ".ToCharArray())[0];
                    Transportnummern.Add(valRef);
                }
            }
            #endregion

            #region SQL Abfrage
            

            IDs = string.Join(", ", TransporeonIDs);
            NRs = "";


            foreach (string nr in Transportnummern)
                if(!NRs.Contains(nr))
                    NRs += $@" OR RefNr LIKE '%{nr}%'" + '\n';


            readSQL(1);
            readSQL(2);

            #endregion
            /*foreach (KeyValuePair<string, string> keyValuePair in translation)
                Console.WriteLine($"({keyValuePair.Key} | {keyValuePair.Value}");*/

            #region Excel Daten eintragen
            Excel.Range formatRange;
            formatRange = page.get_Range("a1");
            formatRange.EntireRow.Font.Bold = true;
            page.Cells[1, 17] = "Position";
            page.Cells[1, 18] = "Eigentümer";

            p = file.Path + "\\" + file.Name + "_modified.xls";

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);

            worker.RunWorkerAsync(worker);
            #endregion
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)e.Argument;

            int progress = 0;

            List<string> xlTransIDs = new List<string>();
            List<string> xlTransNrs = new List<string>();
            for(int i = 2; i <= lastUsedRow; i++)
            {
                xlTransIDs.Add(((Excel.Range)page.Cells[i, 4]).Value2.ToString());
                xlTransNrs.Add(((Excel.Range)page.Cells[i, 3]).Value2.ToString());
            }

            string wtf = "";

            foreach (string s in xlTransIDs)
            {
                foreach(KeyValuePair<string, string> kv in translation)
                {
                    if (kv.Key.Contains(s) || s.Contains(kv.Key) || s == kv.Key)
                    {
                        string value1 = kv.Value;
                        string value2 = tourEigentümer[value1];
                        int i = xlTransIDs.IndexOf(s) + 2;

                        page.Cells[i, 17] = value1;
                        page.Cells[i, 18] = value2;
                    }

                    if (progress != translation.Count + translationRef.Count)
                        worker.ReportProgress(progress, translation.Count + translationRef.Count);
                }
                progress++;
            }
            int lineCounter = 2;
            foreach (string s in xlTransNrs)
            {
                foreach(string key in translationRef.Keys)
                {
                    //wtf += key + " = " + s + " -> " + (key.Contains(s) || s.Contains(key) || s == key).ToString() + '\n';

                    if (key.Contains(s) || s.Contains(key) || s == key)
                    {
                        string value1 = translationRef[key];
                        string value2 = tourEigentümerRef[value1];

                        var cell = page.Cells[lineCounter, 17].Value2;
                        if (cell==null || cell=="" || cell.Contains("NotFound"))
                        {
                            page.Cells[lineCounter, 17] = value1;
                            page.Cells[lineCounter, 18] = value2;
                        }
                    }


                    if (progress != xlTransIDs.Count + xlTransNrs.Count)
                        worker.ReportProgress(progress, xlTransIDs.Count + xlTransNrs.Count);
                }
                progress++;
                lineCounter++;
            }



            /*for (int i = 2; i <= lastUsedRow; i++)
            {
                Console.WriteLine("Write " + i + "/" + lastUsedRow);

                string value1 = "";
                string value2 = "";

                if (translation.TryGetValue(((Excel.Range)page.Cells[i, 4]).Value.ToString(), out value1))
                {
                    tourEigentümer.TryGetValue(value1, out value2);

                    page.Cells[i, 17] = value1;
                    page.Cells[i, 18] = value2;
                }
                else if(getKeyLike(((Excel.Range)page.Cells[i, 3]).Value.ToString(), translationRef, out value1))
                {
                    string FERN = "";
                    translationRef.TryGetValue(value1, out FERN);
                    tourEigentümerRef.TryGetValue(FERN, out value2);

                    page.Cells[i, 17] = FERN;
                    page.Cells[i, 18] = value2;
                }
                else
                {
                    page.Cells[i, 17] = "";
                    page.Cells[i, 17].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    page.Cells[i, 18] = "";
                    page.Cells[i, 18].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                if(i != lastUsedRow)
                    worker.ReportProgress(i, lastUsedRow);
            }*/
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                this.Text = ((e.ProgressPercentage * 100) / (int)e.UserState) + "%";
            }
            catch (NullReferenceException) { }
            catch (Exception) { }
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Text = "Transporeon Rechnungen unten reinziehen";
            this.label1.Text = "Datei hier ablegen";
            file.SaveAs(p, Excel.XlFileFormat.xlWorkbookNormal, null, null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            file.Close(false, null, null);

            MessageBox.Show("Die neue Datei finden sie unter " + p);


            Marshal.ReleaseComObject(page);
            Marshal.ReleaseComObject(file);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        void readSQL(int flag)
        {
            if(flag == 1)
            {
                SqlConnection cnn = new SqlConnection(connectionString);

                cnn.Open();


                this.Text = "SQL-Server-Abfrage 1/2";
                SqlCommand TransporeonIDs_Kommissionsnummern = new SqlCommand("" +
                    $@"  SELECT KommNr, LPosNr, EigenUS 
                    FROM XXASLAuf WHERE
                    KommNr IN ({IDs})
                    AND ErstDat >= DATEADD(MONTH, DATEDIFF(MONTH, 0, DATEADD(MONTH, -6, current_timestamp)), 0)
                    GROUP BY KommNr, LPosNr, EigenUS"
                    , cnn);
                TransporeonIDs_Kommissionsnummern.CommandTimeout = 6000;

                SqlDataReader read = TransporeonIDs_Kommissionsnummern.ExecuteReader(System.Data.CommandBehavior.Default);

                translation = new Dictionary<string, string>();
                tourEigentümer = new Dictionary<string, string>();

                int ind = 1;
                while (read.Read())
                {
                    try
                    {
                        string Komm = "NotFound_" + ind.ToString(), 
                                Pos = "NotFound_" + ind.ToString(), 
                                Usr = "NotFound_" + ind.ToString();

                        if(!read.IsDBNull(0))
                            Komm = read.GetString(0);
                        if (!read.IsDBNull(1))
                            Pos = read.GetString(1);
                        if (!read.IsDBNull(2))
                            Usr = read.GetString(2);

                        if (Komm.Contains("/"))
                            foreach (string partRaw in Komm.Split('/'))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translation.ContainsKey(part) && part.Length > 3)
                                {
                                    translation.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if (translation.ContainsKey(part) && translation[part].Contains("NotFound"))
                                    translation[part] = Pos;
                            }
                        else if (!translation.ContainsKey(Komm))
                                translation.Add(Komm, Pos);
                        if (translation.ContainsKey(Komm) && translation[Komm].Contains("NotFound"))
                            translation[Komm] = Pos;
                        if (!tourEigentümer.ContainsKey(Pos))
                            tourEigentümer.Add(Pos, Usr);
                        else
                        {
                            string oldVal = "NULL";
                            tourEigentümer.TryGetValue(Pos, out oldVal);
                            tourEigentümer.Remove(Pos);
                            tourEigentümer.Add(Pos, $"{oldVal}, {Usr}");
                        }

                    }
                    catch (SqlException sqex) { MessageBox.Show(sqex.ToString()); }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    ind++;
                }

                read.Close();
                cnn.Close();
            }
            else if(flag == 2)
            {
                SqlConnection cnn2 = new SqlConnection(connectionString);

                cnn2.Open();

                this.Text = "SQL-Server-Abfrage 2/2";
                SqlCommand Transportnummern_Referenznummern = new SqlCommand("" +
                    $@"  SELECT ..."
                    , cnn2);
                Transportnummern_Referenznummern.CommandTimeout = 6000;

                SqlDataReader read2 = Transportnummern_Referenznummern.ExecuteReader(System.Data.CommandBehavior.Default);

                translationRef = new Dictionary<string, string>();
                tourEigentümerRef = new Dictionary<string, string>();

                string OutRefs = "";

                int ind2 = 1;
                while (read2.Read())
                {
                    try
                    {
                        /*
                        string Refs = "NotFound_" + ind2.ToString(),
                                Pos = "NotFound_" + ind2.ToString(),
                                Usr = "NotFound_" + ind2.ToString();

                        if (!read2.IsDBNull(0))
                            Refs = read2.GetString(0);
                        if (!read2.IsDBNull(1))
                            Pos = read2.GetString(1);
                        if (!read2.IsDBNull(2))
                            Usr = read2.GetString(2);

                        if (!translationRef.ContainsKey(Refs))
                            translationRef.Add(Refs, Pos);
                        if (!tourEigentümerRef.ContainsKey(Pos))
                            tourEigentümerRef.Add(Pos, Usr);
                        else
                        {
                            string oldVal = "NULL";
                            tourEigentümerRef.TryGetValue(Pos, out oldVal);
                            tourEigentümerRef.Remove(Pos);
                            tourEigentümerRef.Add(Pos, $"{oldVal}, {Usr}");
                        }
                        */
                        string Refs = "NotFound_" + ind2.ToString(),
                                Pos = "NotFound_" + ind2.ToString(),
                                Usr = "NotFound_" + ind2.ToString();

                        if (!read2.IsDBNull(0))
                            Refs = read2.GetString(0);
                        if (!read2.IsDBNull(1))
                            Pos = read2.GetString(1);
                        if (!read2.IsDBNull(2))
                            Usr = read2.GetString(2);


                        ///
                        OutRefs += Refs + '\n';
                        ///


                        if (Refs.Contains("////"))
                            foreach (string partRaw in Refs.Split("////".ToCharArray()))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translationRef.ContainsKey(part) && part.Length > 3)
                                {
                                    translationRef.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if (translationRef.ContainsKey(part) && translationRef[part].Contains("NotFound"))
                                    translationRef[part] = Pos;
                            }
                        else if (Refs.Contains("//"))
                            foreach (string partRaw in Refs.Split("//".ToCharArray()))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translationRef.ContainsKey(part) && part.Length > 3)
                                {
                                    translationRef.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if (translationRef.ContainsKey(part) && translationRef[part].Contains("NotFound"))
                                    translationRef[part] = Pos;
                            }
                        else if (Refs.Contains("/"))
                            foreach (string partRaw in Refs.Split('/'))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translationRef.ContainsKey(part) && part.Length > 3)
                                {
                                    translationRef.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if(translationRef.ContainsKey(part) && translationRef[part].Contains("NotFound"))
                                    translationRef[part] = Pos;
                            }
                        else if (Refs.Contains("-"))
                            foreach (string partRaw in Refs.Split('-'))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translationRef.ContainsKey(part) && part.Length > 3)
                                {
                                    translationRef.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if (translationRef.ContainsKey(part) && translationRef[part].Contains("NotFound"))
                                    translationRef[part] = Pos;
                            }
                        else if (Refs.Contains("+"))
                            foreach (string partRaw in Refs.Split('+'))
                            {
                                string part = partRaw.Trim(' ');
                                if (!translationRef.ContainsKey(part) && part.Length > 3)
                                {
                                    translationRef.Add(part, Pos);
                                    Console.WriteLine(part);
                                }
                                else if (translationRef.ContainsKey(part) && translationRef[part].Contains("NotFound"))
                                    translationRef[part] = Pos;
                            }
                        else if (!translationRef.ContainsKey(Refs))
                            translationRef.Add(Refs, Pos);
                        if (translationRef.ContainsKey(Refs) && translationRef[Refs].Contains("NotFound"))
                            translationRef[Refs] = Pos;
                        if (!tourEigentümerRef.ContainsKey(Pos))
                            tourEigentümerRef.Add(Pos, Usr);
                        else
                        {
                            string oldVal = "NULL";
                            tourEigentümerRef.TryGetValue(Pos, out oldVal);
                            tourEigentümerRef.Remove(Pos);
                            tourEigentümerRef.Add(Pos, $"{oldVal}, {Usr}");
                        }
                    }
                    catch (SqlException sqex) { MessageBox.Show(sqex.ToString()); }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    ind2++;
                }

                read2.Close();
                cnn2.Close();

            }
        }

        bool getKeyLike(string comparer, Dictionary<string, string> dic, out string inDicKey)
        {
            string comp = comparer;

            foreach(KeyValuePair<string, string> keyVal in dic)
                if(keyVal.Key.Contains(comp))
                {
                    inDicKey = keyVal.Key;
                    return true;
                }

            inDicKey = "";
            return false;
        }
    }
}
