using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ScanExcel
{
    public partial class Form1 : Form
    {
        volatile public bool scanning = false;
        private Thread scanner;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            scanner = new Thread(Communicate);
            scanner.Start();
            scanning = true;
        }

        public void Communicate()
        {
            while (scanning)
            {
                textBox1.Invoke((MethodInvoker)(delegate { this.textBox1.Text = "Hello"; }));
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (scanning)
            {
                scanning = false;
                scanner.Join();
            }
        }
    }
}
