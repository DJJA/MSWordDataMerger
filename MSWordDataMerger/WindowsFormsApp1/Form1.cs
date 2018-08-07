using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using GemBox.Document;
using MSWordDataMerger.Logic;
using Xceed.Words.NET;
using Paragraph = Xceed.Words.NET.Paragraph;
using ThreadState = System.Threading.ThreadState;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Merger merger = new Merger(KeyValuePairLoaderType.SeperateFileBased, TemplateEditorType.XceedSoftwareDocX, @"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\MSWordDataMerger\Offerte template");

            merger.Merge();
        }

        private Thread t;
        private bool canceled;

        private void button2_Click(object sender, EventArgs e)
        {
            canceled = false;
            t = new Thread(new ThreadStart(loop));
            t.Start();

            MessageBox.Show("Thread running...");
        }

        private void loop()
        {
            while (!canceled)
            {
                Thread.Sleep(1000);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            canceled = true;
            while (t.ThreadState != ThreadState.Stopped)
            {
                Thread.Sleep(250);
            }

            MessageBox.Show("Thread stopped.");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show(DateTime.Now.ToString("yyyyMMdd"));
        }

        private MSWordDataMerger msWordDataMerger = new MSWordDataMerger();

        private void btnStart_Click(object sender, EventArgs e)
        {
            msWordDataMerger.OnStart(new String[]{});
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            msWordDataMerger.OnStop();
        }
    }
}
