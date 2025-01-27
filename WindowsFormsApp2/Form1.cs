using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronOcr;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Diagnostics;
using Accord;
using AForge.Video;
using AForge.Video.DirectShow;
using System.IO;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Globalization;
using System.IO.Ports;

//Debug line because I forget the syntax
//Debug.WriteLine();

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {

        List<string> CardsScannedlist = new List<string>();
        string[] CardImageStack = new string[12];

        FilterInfoCollection filterInfoCollection;
        VideoCaptureDevice videoCaptureDevice;      

        DataTable CardTable = new DataTable();

        public bool CamStartStop = true; 
        public bool SerialStop = true;
        public bool moveWaiting = true;

        // string serialDataIn = "1";

        decimal CardvalueTotal = 0;

        SerialPort CardScanPort = new SerialPort();

        public Form1()
        {
            //Start the form code
            InitializeComponent();
            GetListCameraUSB();

            //Set startup sate and image blanks
            button1.Enabled = false;
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox2.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox3.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox4.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox5.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox6.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox7.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox8.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox9.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox10.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox11.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox12.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox13.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox14.SizeMode = PictureBoxSizeMode.Zoom;

            pictureBox14.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox13.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox12.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox11.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox10.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox9.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox8.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox7.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox5.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox4.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox3.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
            pictureBox6.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");

            CardTable.Columns.Add("Card Name", typeof(string));
            CardTable.Columns.Add("Card Value", typeof(double));
            CardTable.Columns.Add("Card Rarity", typeof(string));
            CardTable.Columns.Add("Card CMC", typeof(int));
            CardTable.Columns.Add("Card Colour", typeof(string));

            CardsScannedlist.Insert(0, "Total Value \t $0.00");
        }

        public void moveWait()
        {
            SerialPort CardScanPort = new SerialPort(comboBox2.Text, 115200, Parity.None, 8, StopBits.One);
            CardScanPort.ReadTimeout = 5000;
            CardScanPort.Open();

            while (moveWaiting)
            {
                try
                {
                    string indata = CardScanPort.ReadLine();
                    Console.Write(indata);

                    if (indata.Contains("Y"))
                    {
                        moveWaiting = false;
                        Console.WriteLine("Move Complete");
                    }
                }
                catch (TimeoutException) {
                    moveWaiting = false;
                    Console.WriteLine("Timeout");
                }
            }

            moveWaiting = true;
            CardScanPort.Close();
        }

        private void StopCamera()
        {
            //If the camera is on, turn it off
            if (filterInfoCollection.Count != 0)
            {
                if (videoCaptureDevice.IsRunning == true)
                {
                    videoCaptureDevice.Stop();
                }
            }
        }

        private void SaveData()
        {
            var workbook = new XLWorkbook();

            workbook.Worksheets.Add(CardTable, "WorksheetName");

            string dummyFileName = "Save Here";

            SaveFileDialog sf = new SaveFileDialog();
            // Feed the dummy name to the save dialog
            sf.FileName = dummyFileName;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                // Now here's our save folder
                string savePath = Path.GetDirectoryName(sf.FileName);
                // Do whatever
                Debug.WriteLine(savePath);

                string SaveTime = string.Format("{0:yyyy-MM-dd_hh-mm-ss}", DateTime.Now);

                workbook.SaveAs(savePath + "\\CardList-" + SaveTime + ".xlsx");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Form close routine
            SaveData();
            StopCamera();
            if (CardScanPort.IsOpen)
            {
                SerialStop = true;
            }
                File.Delete("..\\..\\..\\res\\card.png");
        }

        private void GetListCameraUSB()
        {
            //Get a list of the avalible camears and populate the combobox
            filterInfoCollection = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            if (filterInfoCollection.Count != 0)
            {
                foreach (FilterInfo filterinfo in filterInfoCollection)
                {
                    comboBox1.Items.Add(filterinfo.Name);
                    comboBox1.SelectedIndex = 0;
                    videoCaptureDevice = new VideoCaptureDevice();
                }
            }
            else
            {
                //Text to add if no cameras are found
                comboBox1.Items.Add("No Cameras Found");
            }
            comboBox1.SelectedIndex = 0;
        }

        private void XYControl(int X, int Y)
        {
            SerialPort CardScanPort = new SerialPort(comboBox2.Text, 115200, Parity.None, 8, StopBits.One);
            CardScanPort.Open();
            CardScanPort.Write("<" + X.ToString() + "," + Y.ToString() + ">");
            CardScanPort.Close();
        }

        private void ScanCard()
        {
            int retryCount = 0;

            //Set text feilds while card is scanned
            richTextBox1.Text = "Data Laoding";
            richTextBox5.Text = "Data Laoding";

        //Pointer for rescan
        retryscan:

            //Where the magic happens
            //Turn off the scan button to disable starting a new scan if a scan is in progress
            button1.Enabled = false;

            //Save a frame from the webcam
            pictureBox2.Image.Save(@"..\\..\\..\\res\\card.png", System.Drawing.Imaging.ImageFormat.Png);

            //load prevously saved imagine into memory for use
            using (FileStream fs = new FileStream("..\\..\\..\\res\\card.png", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, (int)fs.Length);
                using (MemoryStream ms = new MemoryStream(buffer))
                    this.pictureBox1.Image = Image.FromStream(ms);
            }

            //OCR On currnet frame
            IronTesseract IronOcr = new IronTesseract();
            var Result = IronOcr.Read("..\\..\\..\\res\\card.png");
            string text = Result.Text;

            //Remove blank lines
            text = Regex.Replace(text, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline).TrimEnd();

            //Dispaly raw text
            richTextBox1.Text = text;

            //Split string by lines
            string[] result = text.Split('\n');

            var totalCount = result.Count();

            if (totalCount >= 2)
            {
                //Find last feilds from OCR where set name and number are found
                string setName = result[totalCount - 1];
                string setNum = result[totalCount - 2];
                string setNumOut;
                int setNumLen;

                //clean up set name feild to needed numbers/letters only
                Debug.WriteLine("setName");
                Debug.WriteLine(setName);
                setName = setName.Remove(3);
                setName = setName.ToLower();

                setNum = setNum.Replace("T", "");
                setNum = setNum.Replace("S", "");
                setNum = setNum.Replace("L", "");
                setNum = setNum.Replace("C", "");
                setNum = setNum.Replace("R", "");
                setNum = setNum.Replace("U", "");
                setNum = setNum.Replace("M", "");
                setNum = setNum.Replace("H", "");
                setNum = setNum.Trim();

                setNumLen = setNum.Length;

                if (setNumLen <= 3)
                {
                    setNum = setNum.Replace("/", "");
                    setNum = setNum.TrimStart('0');
                    setNumOut = setNum.Trim();
                }
                else if (setNumLen == 4)
                {
                    setNum = setNum.Replace("/", "");
                    setNum = setNum.TrimStart('0');
                    setNumOut = setNum.Trim();
                }
                else
                {
                    setNum = setNum.Remove(4);
                    setNum = setNum.Replace("/", "");
                    setNum = setNum.TrimStart('0');
                    setNumOut = setNum.Trim();
                }

                // Create an API URI
                string[] cardID = { "https://api.scryfall.com/cards/", setName, "/", setNumOut };

                string output = String.Join("", cardID);

                //Display the URI
                //richTextBox2.Text = output;

                //Call the URI
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "xxxxxxxxx");
                    client.DefaultRequestHeaders.Add("Accept", "*/*");
                    var endpoint = new Uri(output);
                    var scrfallAPI = client.GetAsync(endpoint).Result.Content.ReadAsStringAsync().Result;


                    Debug.WriteLine("check output");
                    Debug.WriteLine(endpoint);

                    Debug.WriteLine(scrfallAPI);


                    //Filter data for check return
                    var scrfallAPIAPICheck = scrfallAPI.Remove(20);
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace("/", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace("\n", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace(":", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace(" ", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace("{", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace(",", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace(".", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Replace("\"", "");
                    scrfallAPIAPICheck = scrfallAPIAPICheck.TrimEnd('i', 'd');
                    scrfallAPIAPICheck = scrfallAPIAPICheck.Remove(0, 6);




                    if (scrfallAPIAPICheck == "error" && retryCount < 5)
                    {
                        retryCount++;

                        string[] retryCountText = { "Scan failed, retying ", retryCount.ToString(), " times"};
                        textBox1.Text = String.Join("", retryCountText);
                        Debug.WriteLine(String.Join("", retryCount));

                        goto retryscan;
                    }


                    if (scrfallAPIAPICheck == "car")
                    {
                        //Pull useful data out of api return

                        //Card image URL
                        int CardImgLoc = scrfallAPI.IndexOf("\"normal\":") + "\"normal\":".Length;

                        string CardImg = scrfallAPI.Remove(0, CardImgLoc + 1);
                        CardImgLoc = CardImg.IndexOf("\"");

                        CardImg = CardImg.Remove(CardImgLoc);
                        Debug.WriteLine(CardImg);

                        //Card name
                        int CardNameLoc = scrfallAPI.IndexOf("\"name\":") + "\"name\":".Length;

                        string CardName = scrfallAPI.Remove(0, CardNameLoc + 1);
                        CardNameLoc = CardName.IndexOf("\"");

                        CardName = CardName.Remove(CardNameLoc);

                        //Card current value
                        int CardvalueLoc = scrfallAPI.IndexOf("\"prices\":") + "\"prices\":".Length;

                        string Cardvalue = scrfallAPI.Remove(0, CardvalueLoc + 8);
                        CardvalueLoc = Cardvalue.IndexOf("\"");
                        Cardvalue = Cardvalue.Remove(CardvalueLoc);

                        //Card rarity
                        int CardRarityLoc = scrfallAPI.IndexOf("\"rarity\":") + "\"rarity\":".Length;

                        string CardRarity = scrfallAPI.Remove(0, CardRarityLoc + 1);
                        CardRarityLoc = CardRarity.IndexOf("\"");
                        CardRarity = CardRarity.Remove(CardRarityLoc);
                        CardRarity = Regex.Replace(CardRarity, @"\b(\w)", m => m.Value.ToUpper());

                        //Card current value foil
                        int CardvaluefoilLoc = scrfallAPI.IndexOf("\"prices\":") + "\"prices\":".Length;

                        string Cardvaluefoil = scrfallAPI.Remove(0, CardvaluefoilLoc + 8); Debug.WriteLine(Cardvaluefoil);
                        CardvaluefoilLoc = Cardvaluefoil.IndexOf("\"");
                        Cardvaluefoil = Cardvaluefoil.Remove(CardvaluefoilLoc);

                        //Card CMC
                        int CardCMCLoc = scrfallAPI.IndexOf("\"cmc\":") + "\"cmc\":".Length;

                        string CardCMC = scrfallAPI.Remove(0, CardCMCLoc);
                        CardCMCLoc = CardCMC.IndexOf("\"");
                        CardCMC = CardCMC.Remove(CardCMCLoc - 3);

                        //Card colour
                        int CardcolourLoc = scrfallAPI.IndexOf("\"colors\":") + "\"colors\":".Length;

                        string Cardcolour = scrfallAPI.Remove(0, CardcolourLoc + 1);
                        CardcolourLoc = Cardcolour.IndexOf("]");
                        Cardcolour = Cardcolour.Remove(CardcolourLoc);

                        int W = Cardcolour.IndexOf("W");
                        int U = Cardcolour.IndexOf("U");
                        int B = Cardcolour.IndexOf("B");
                        int R = Cardcolour.IndexOf("R");
                        int G = Cardcolour.IndexOf("G");

                        List<string> Cardcolourlist = new List<string>();


                        if (W != -1)
                        {
                            Cardcolourlist.Add("White");
                        }
                        if (U != -1)
                        {
                            Cardcolourlist.Add("Blue");
                        }
                        if (B != -1)
                        {
                            Cardcolourlist.Add("Black");
                        }
                        if (R != -1)
                        {
                            Cardcolourlist.Add("Red");
                        }
                        if (G != -1)
                        {
                            Cardcolourlist.Add("Green");
                        }
                        if (W == -1 && U == -1 && B == -1 && R == -1 && G == -1)
                        {
                            Cardcolourlist.Add("Colorless");
                        }

                        String[] CardcolourArry = Cardcolourlist.ToArray();

                        //Display the raw URI results
                        richTextBox5.Text = "Card Name \t" + CardName;
                        if (Cardvalue == "ull,")
                        {
                            richTextBox5.AppendText("\n" + "Card Value \t" + "No Price for this Foiling");
                        }
                        else
                        {
                            richTextBox5.AppendText("\n" + "Card Value \t" + "$" + Cardvalue + " USD");
                        }
                        richTextBox5.AppendText("\n" + "Card Rarity \t" + CardRarity);
                        richTextBox5.AppendText("\n" + "Card CMC \t" + CardCMC);
                        richTextBox5.AppendText("\n" + "Card Colour \t" + String.Join("\n \t \t", CardcolourArry));

                        CardvalueTotal = CardvalueTotal + Convert.ToDecimal(Cardvalue);

                        CardsScannedlist.Insert(0, "Total Value \t \t $" + decimal.Round(CardvalueTotal, 2).ToString("0.00") + "\n" + "- - - - - - - - -");
                        CardsScannedlist.RemoveAt(1);

                        if (CardRarity == "Uncommon")
                        {
                            CardsScannedlist.Insert(1, CardName + "\n" + CardRarity + "\t \t $" + Cardvalue + "\n" + "- - - - - - - - -");
                        }
                        else
                        {
                            CardsScannedlist.Insert(1, CardName + "\n" + CardRarity + "\t \t \t $" + Cardvalue + "\n" + "- - - - - - - - -");
                        }
                        
                        String[] CardsScannedArry = CardsScannedlist.ToArray();

                        richTextBox2.Text = String.Join("\n", CardsScannedArry);

                        CardTable.Rows.Add(CardName, Convert.ToDouble(Cardvalue), CardRarity, Convert.ToInt32(CardCMC), String.Join(",", CardcolourArry));

                        for (int i = 11; i > 0; i--)
                        {
                            CardImageStack[i] = CardImageStack[i - 1];
                        }

                        CardImageStack[0] = CardImg;

                        pictureBox14.ImageLocation = CardImageStack[0];
                        pictureBox13.ImageLocation = CardImageStack[1];
                        pictureBox12.ImageLocation = CardImageStack[2];
                        pictureBox11.ImageLocation = CardImageStack[3];
                        pictureBox10.ImageLocation = CardImageStack[4];
                        pictureBox9.ImageLocation = CardImageStack[5];
                        pictureBox8.ImageLocation = CardImageStack[6];
                        pictureBox7.ImageLocation = CardImageStack[7];
                        pictureBox5.ImageLocation = CardImageStack[8];
                        pictureBox4.ImageLocation = CardImageStack[9];
                        pictureBox3.ImageLocation = CardImageStack[10];
                        pictureBox6.ImageLocation = CardImageStack[11];
                    }
                    else
                    {
                        //Display fail message if api return is in error
                        textBox2.Text = "--   Scan: Fail!   --";
                    }
                }
            }
            else
            {
                //Display fail message if no card deteced
                richTextBox1.Text = "No Vaild Data Found";
                richTextBox5.Text = "No Vaild Data Found";
            }
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ScanCard();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (CamStartStop == true)
            {
                //State of button when no camera active
                button2.Text = "Stop Camera";
                CamStartStop = false;

                //Start the camera and set other elements
                if (filterInfoCollection.Count != 0)
                {
                    
                    videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[comboBox1.SelectedIndex].MonikerString);
                    videoCaptureDevice.NewFrame += VideoCaptureDevice_NewFrame;
                    comboBox1.Enabled = false;
                    button1.Enabled = true;
                    videoCaptureDevice.VideoResolution = videoCaptureDevice.VideoCapabilities[0];
                    videoCaptureDevice.Start();

                    



                }
                else
                {
                    //State of button when camera active
                    string message = "No Camera Selected";
                    MessageBox.Show(message);
                }
            }
            else
            {
                button2.Text = "Start Camera";
                StopCamera();
                CamStartStop = true;
                pictureBox2.Image = Image.FromFile(@"..\\..\\..\\res\\camPlace.png");
                comboBox1.Enabled = true;
                button1.Enabled = false;
                richTextBox1.Text = "Waiting...";
                richTextBox5.Text = "Waiting...";
            }
        }

        private void VideoCaptureDevice_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            pictureBox2.Image = (Bitmap)eventArgs.Frame.Clone();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] portNames = SerialPort.GetPortNames();     //<-- Reads all available comPorts
            foreach (var portName in portNames)
            {
                comboBox2.Items.Add(portName);                  //<-- Adds Ports to combobox
            }

            if (comboBox2.Items.Count == 0)
            {
                comboBox2.Items.Add("No Ports");
                comboBox2.SelectedIndex = 0;
            }


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            SaveData();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (SerialStop == true)
            {
                //State of button when no camera active
                button4.Text = "Disconnect Scanner";
                SerialStop = false;
                comboBox2.Enabled = false;

                try
                {
                    SerialPort CardScanPort = new SerialPort(comboBox2.Text, 115200, Parity.None, 8, StopBits.One);
                    CardScanPort.Open();
                    CardScanPort.Close();
                }
                catch (Exception)
                {
                    button4.Text = "Connect Scanner";
                    SerialStop = true;
                    comboBox2.Enabled = true;
                    string message = "Port Busy";
                    MessageBox.Show(message);
                }

            }
            else
            {
                button4.Text = "Connect Scanner";
                SerialStop = true;
                comboBox2.Enabled = true;
            }
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            XYControl(0, 0);
            moveWait();
           // ScanCard();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            XYControl(10000, 10000);
            moveWait();
          //  ScanCard();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            XYControl(-10000, -10000);
            moveWait();
           // ScanCard();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            XYControl(-100, -100);

        }

        private void button8_Click(object sender, EventArgs e)
        {
            XYControl(100, 100);
            moveWait();
            //ScanCard();
        }
    }
}