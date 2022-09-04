using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var Centre_intimation = new Centre_intimation();
            Centre_intimation.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var CNS = new CNS();
            CNS.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var HE_Acceptance = new HE_Acceptance();
            HE_Acceptance.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var Self_Centre = new Self_Centre();
            Self_Centre.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var HE_letter = new HE_letter();
            HE_letter.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var Examiner_letter = new Examiner_letter();
            Examiner_letter.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var VOCPRAC_letter = new VOCPRAC_letter();
            VOCPRAC_letter.Show();
        }
    }
}
