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

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {

        FilterInfoCollection filterInfoCollection;
        VideoCaptureDevice videoCaptureDevice;

        public Form1()
        {
            InitializeComponent();
            getListCameraUSB();
        }

        private void getListCameraUSB()
        {

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
                comboBox1.Items.Add("No Cameras Found");
            }

            comboBox1.SelectedIndex = 0;
            //
        }

            private void button1_Click(object sender, EventArgs e)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

            OpenFileDialog open = new OpenFileDialog();
            // image filters  
            open.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp";
            if (open.ShowDialog() == DialogResult.OK)
            {
                var totalCount = 0;

                // display image in picture box  
                pictureBox1.Image = new Bitmap(open.FileName);
                // image file path  
                string PicPath = open.FileName;

                //OCR On currnet image
                IronTesseract IronOcr = new IronTesseract();
                var Result = IronOcr.Read(PicPath);
                //Dispaly raw text
                richTextBox1.Text = Result.Text;

                // split string 
                string text = Result.Text;                
                string[] result = text.Split('\n');

                //Format set name and numnber

                //Find last feilds from OCR where set name and number are found
                totalCount = result.Count();

                string setName = result[totalCount - 1];
                string setNum = result[totalCount - 2];
                string setNumOut;
                int setNumLen;

                //clean up set name feild to needed numbers only
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
                else if(setNumLen == 4)
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
                string[] cardID = { "https://api.scryfall.com/cards/",setName,"/",setNumOut};

                string output = String.Join("", cardID);

                //Display the URI
                richTextBox2.Text = output;

                //Call the URI
                using(var client = new HttpClient())
                {
                    var endpoint = new Uri(output);
                    var scrfallAPI = client.GetAsync(endpoint).Result.Content.ReadAsStringAsync().Result;
                    //Debug.WriteLine(scrfallAPI);

                    //Display the raw URI results
                    richTextBox3.Text = scrfallAPI;

                }
            }
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (filterInfoCollection.Count != 0)
            {
                videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[comboBox1.SelectedIndex].MonikerString);
                videoCaptureDevice.NewFrame += VideoCaptureDevice_NewFrame;
                videoCaptureDevice.Start();
            }
            else
            {
                string message = "No Camera Selected";
                MessageBox.Show(message);
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (filterInfoCollection.Count != 0)
            {
                if (videoCaptureDevice.IsRunning == true)
                {
                    videoCaptureDevice.Stop();
                }
            }
        }
    }
}

