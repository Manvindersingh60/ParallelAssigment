using System;
using System.Windows.Forms;

namespace ParallelAssigment
{
    public partial class ProcessingForm : Form
    {
        public ProcessingForm()
        {
            InitializeComponent();
            TopMost = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Utility.StopRequest();
        }

        private void ProcessingForm_Load(object sender, EventArgs e)
        {
        }
    }
}