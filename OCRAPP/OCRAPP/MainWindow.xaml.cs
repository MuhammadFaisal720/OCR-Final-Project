
using CsvHelper;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Tesseract;
using SautinSoft.Document;
using System.Windows.Media.TextFormatting;

namespace OCRAPP
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const bool V = true;
        string file, UploadDate, DueDateString;
        DateTime Uploaddates, duedates;
        string Filetext;
        List<string> Files = new List<string>();
        List<string> FileUploadDates = new List<string>();
        List<DateTime> FileDueDates = new List<DateTime>();
        private Timer timer;

        public MainWindow()
        {
            InitializeComponent();
            timer = new System.Timers.Timer();
            timer.Interval = 1000;
            timer.Elapsed += Timer_Elapsed;
            submitbtn.IsEnabled = false;
            setreminderbtn.IsEnabled = false;
            clearbtn.IsEnabled = false;
            saveasdocxbtn.IsEnabled = false;
            saveastxtbtn.IsEnabled = false;
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            DateTime currentTime = DateTime.Now;
            List<DateTime> dateTimes = FileDueDates;
            foreach(DateTime DateT in dateTimes)
            {
                if (DateTime.Today.Date == DateT)
                {
                    timer.Stop();
                    try {
                        int index = dateTimes.FindIndex(a => a == DateT);
                        MessageBox.Show(Files[index]+"Assignment Deadline.");
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message,"Message",MessageBoxButton.OK);
                    }
                }
            }
        }
        private void choosefilebtn_Click(object sender, RoutedEventArgs e)
        {
            file = "";
            filepathtxt.Text = "";
            inputimage.Source = null;
            resulttxtblk.Text = null;
            meanconfidencetxt.Text = null;
            wordresulttxt.Items.Clear();


            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                filepathtxt.Text = openFileDialog.FileName;
            file = openFileDialog.FileName;
            try
            {
                inputimage.Source = new BitmapImage(new Uri(file));
            }
            catch { }
            remindercalender.SelectedDates.Clear();
            submitbtn.IsEnabled = true;
;
        }

        private void submitbtn_Click(object sender, RoutedEventArgs e)
        {
            
            setreminderbtn.IsEnabled = true;
            clearbtn.IsEnabled = true;
            saveasdocxbtn.IsEnabled = true;
            saveastxtbtn.IsEnabled = true;
            string str;
            int wrd, l;
            Uploaddates = DateTime.Today.Date;
            UploadDate = Uploaddates.ToString("MM/dd/yyyy");
            uploaddatetxt.Text = UploadDate;
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
            path = Path.Combine(path, "tessdata");
            path = path.Replace("file:\\", "");
            if (file == null || file.Length == 0)
            {

                statustxt.Text = "File not Found";

            }
            else
                // string Text = new IronTesseract().Read(@"img\Screenshot.png").Text;
                using (var engine = new TesseractEngine(path, "eng", EngineMode.TesseractAndLstm))
                {
                    using (var image = new System.Drawing.Bitmap(file))
                    {
                        using (var pix = PixConverter.ToPix(image))
                        {
                            using (var page = engine.Process(pix))
                            {
                                l = 0;
                                wrd = 1;
                                str = page.GetText();
                                // string word = "";

                                while (l <= str.Length - 1)
                                {
                                    /* check whether the current character is white space or new line or tab character*/
                                    if (str[l] == ' ' || str[l] == '\n' || str[l] == '\t')
                                    {
                                        wrd++;
                                    }

                                    l++;
                                }
                                Dictionary<string, int> dictionary = new Dictionary<string, int>();

                                string sInput = str;
                                sInput = sInput.Replace(",", ""); //Just cleaning up a bit
                                sInput = sInput.Replace(".", ""); //Just cleaning up a bit
                                string[] arr = sInput.Split(' '); //Create an array of words

                                foreach (string words in arr) //let's loop over the words
                                {
                                    if (words.Length >= 3) //if it meets our criteria of at least 3 letters
                                    {
                                        if (dictionary.ContainsKey(words)) //if it's in the dictionary
                                            dictionary[words] = dictionary[words] + 1; //Increment the count
                                        else
                                            dictionary[words] = 1; //put it in the dictionary with a count 1
                                    }
                                }

                                resulttxtblk.Text = str;
                                meanconfidencetxt.Text = String.Format("{0:p}", page.GetMeanConfidence());
                                noofwordstxt.Text = Convert.ToString(wrd);
                                Dictionary<string, int> dict = dictionary.OrderByDescending(u => u.Value).ToDictionary(z => z.Key, y => y.Value);
                                foreach (KeyValuePair<string, int> pair in dict) //loop through the dictionary
                                    wordresulttxt.Items.Add(string.Format("Word: {0}, Count: {1}", pair.Key, pair.Value));
                                string testDate = str;
                                
                                try
                                {
                                    Regex rgx = new Regex(@"\d{2}-\d{2}-\d{4}");
                                    Match mat = rgx.Match(testDate);
                                    DueDateString = Convert.ToString(mat);
                                    DueDateString = DueDateString.Replace("-", "/");

                                    CultureInfo culture = new CultureInfo("en-US");
                                    duedates = DateTime.ParseExact(DueDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    DueDateString = duedates.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);


                                    // 
                                    //duedates = Convert.ToDateTime.(DueDateString);
                                    // DueDateString = duedates.ToString("dd/MM/yyyy");
                                    duedatetxt.Text = DueDateString;
                                }
                                catch {
                                    datastatustxt.Text = "No Due Date avaible you can add reminder.";
                                    setreminderbtn.IsEnabled = false; 
                                }
                                Filetext = str;
                            }
                        }
                    }
                }

        }

        private void clearbtn_Click(object sender, RoutedEventArgs e)
        {
            file = "";
            filepathtxt.Text = "";
            inputimage.Source = null;
            resulttxtblk.Text = null;
            meanconfidencetxt.Text = null;
            wordresulttxt.Items.Clear();
            dataset.Text = null;
            duedatetxt.Text = null;
            uploaddatetxt.Text = null;
            dataset.Text = null;
        }

        private void exitbtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void saveastxtbtn_Click(object sender, RoutedEventArgs e)
        {
            string newFileName = file + ".txt";
            TextWriter txt = new StreamWriter(newFileName);
            txt.Write(Filetext);
            txt.Close();
            savestatustxt.Text = "Text File has been saved at location of assignment image.";
        }

        private void reminder0ffbtn_Click(object sender, RoutedEventArgs e)
        {
            timer.Stop();
            datastatustxt.Text = "Reminders Stops";
        }

        private void saveasdocxbtn_Click_1(object sender, RoutedEventArgs e)
        {
            // Assume we already have a document 'dc'.
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = Encoding.ASCII.GetBytes(Filetext);
            string filePath = @file + ".docx";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert(Filetext);

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                dc.Save(ms, new DocxSaveOptions());
                fileData = ms.ToArray();
            }
            File.WriteAllBytes(filePath, fileData);
            savestatustxt.Text = "Docx File has been saved at location of assignment image.";
            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        private void meanconfidencetxt_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }

        private void saveasdocxbtn_Click(object sender, RoutedEventArgs e)
        {
            // var ookiiDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            //  if (ookiiDialog.ShowDialog() == true)
            //       root_folder_TextBox.Text = ookiiDialog.SelectedPath;
            //  file = openFileDialog.FileName;
            //  inputimage.Source = new BitmapImage(new Uri(file));
        }

        private void setreminderbtn_Click(object sender, RoutedEventArgs e)
        {
            if (duedates < Uploaddates)
            {
                datastatustxt.Text = "Due date is old deadline missed already.";
            }
            else
            {
                // string strSeperator = ",";
                string newFileName = "C:\\list\\ocr.txt";
                string newFileNameduedate = "C:\\list\\duedateocr.txt";
                //TextWriter txt = new StreamWriter(newFileName);
                //TextWriter txtduedate = new StreamWriter(newFileNameduedate);
                using (StreamWriter txtduedate = File.AppendText(newFileNameduedate))
                {
                    txtduedate.WriteLine(DueDateString);
                    txtduedate.Close();
                }
                 // txtduedate.WriteLine(DueDateString);
                string datas = file + "," + UploadDate + "," + DueDateString + ";";
                using (StreamWriter txt = File.AppendText(newFileName))
                {
                    txt.WriteLine(datas);
                    txt.Close();
                }
                //txt.WriteLine(datas);
                
                

                string newFileName1 = "C:\\list\\ocr.csv";
                string clientDetails = file + "," + UploadDate + "," + DueDateString + ";";
                
                //remindercalender.SelectedDates.Add(duedates);

                if (!File.Exists(newFileName1))
                {
                    string clientHeader = "File" + "," + "UploadDate" + "," + "DueDate" + Environment.NewLine;

                    File.WriteAllText(newFileName1, clientHeader);
                }
                try
                {
                    File.AppendAllText(newFileName1, clientDetails);
                    string path = newFileName;

                    string[] datalines = System.IO.File.ReadAllLines(path);
                    foreach (string dataline in datalines)
                    {
                        listtxtlist.Items.Add(dataline);

                    }
                }
                catch
                {
                    datastatustxt.Text = "Reminder Already Existed.";
                }
                remindercalender.SelectedDates.Clear();
                remindercalender.SelectedDates.Add(Convert.ToDateTime(duedates));
                remindercalender.SelectedDates.Contains(remindercalender.SelectedDate.Value);
                datastatustxt.Text = "Reminder has been added.";
                dataset.Text = file + "," + UploadDate + "," + DueDateString;
                string[] datelines = System.IO.File.ReadAllLines(newFileNameduedate);
                string[] lines = System.IO.File.ReadAllLines(newFileName);
                foreach (string line in lines)
                {

                    Files.Add(line);

                }
                foreach (string dt in datelines)
                {

                    FileDueDates.Add(Convert.ToDateTime(dt));

                }

                timer.Start();
            }

        }
    }
}





    

