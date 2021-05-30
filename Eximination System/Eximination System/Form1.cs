using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using MicrosoftEdgecls;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Threading;

namespace Eximination_System
{
    public partial class Form1 : Form
    {
        MicrosoftEdgecls.MsEdgeOperations edge = new MsEdgeOperations();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        private void btn_submit_Click(object sender, EventArgs e)
        {
            try
            {
               
               edge.Invoke(textBox1.Text);

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Data.ToString());
            }
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (edge.GetResult(textBox1.Text))
                {
                    MessageBox.Show("he is sucesseded");
                }
                else
                    MessageBox.Show("he is Failed");
                edge.ResetToDefault(textBox1.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

    }

}



