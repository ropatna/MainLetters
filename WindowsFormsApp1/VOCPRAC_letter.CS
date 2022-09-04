using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class VOCPRAC_letter : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=COMP2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public VOCPRAC_letter()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            SqlCommand cmd2 = new SqlCommand("", sqlcon);
            SqlCommand cmd3 = new SqlCommand("", sqlcon);
            if (String.IsNullOrEmpty(textBox6.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[CITYCOORD22]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[CITYCOORD22] where SCH_CC='" + textBox6.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                SqlConnection sqlcon2 = new SqlConnection(connectionString: "Data Source=COMP2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
                sqlcon2.Open();
                cmd2 = new SqlCommand("select * FROM [LETTERS].[dbo].[VOC_SCH2] where sch_cc='" + dr["sch_cc"].ToString() + "'", sqlcon2);
                SqlDataReader dr2 = cmd2.ExecuteReader();
                Document doc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(doc, new FileStream("E:\\abhi\\pdf\\OTHER\\" + dr["SCH_CC"].ToString() + "_VOC.pdf", FileMode.Create));
                var header = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\header.png");
                var footer = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\FOOTER.png");
                var rosign = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\rosignpng.png");
                var header2 = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\acceptance_header.png");
                header2.ScaleToFit(900f, 60f);
                header2.Alignment = 1;
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(90f, 30f);
                rosign.SetAbsolutePosition(470, 200);
                doc.Open();
                doc.Add(header); //Adding Header
                doc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                iTextSharp.text.Font bold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 12);
                doc.AddTitle("SELF CENTRE CS Letter");
                Paragraph p = new Paragraph("===============================================================================\n");
                doc.Add(p);
                Paragraph p4 = new Paragraph(str: "No.CBSE/RO/PTN/Conf./Prac.Exam/2022/                                                   Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                doc.Add(p4);
                Paragraph p5 = new Paragraph(str: dr["ccname"].ToString() + "\nPrincipal cum City Co-ordinator,\n(" + dr["sch_cc"].ToString() + ")" + dr["cc_abbr"].ToString().ToUpper() + "\n") { Alignment = Element.ALIGN_LEFT };
                doc.Add(p5);
                Paragraph p1 = new Paragraph(str: "      Sub:    Arrangement for Practical Examiner in Vocational/Minor subjects in r/o Practical Exam- Term 2-2022 - reg.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                doc.Add(p1);
                Paragraph p2 = new Paragraph(str: "      Sir,\n      You are aware that the Practical Examination for All India Senior School Certificate Examination - 2022 has been commencing from 26th April, 2022 (Last date would be 10 days before the Examination of respective subject) and all the preparations regarding smooth conduct of Practical Examination has already been made by school under your supervision.\n\n      Some of the school candidates have opted the vocational /minor subjects for which Practical Examiners cannot be appointed due to insufficient teacher’s data available with the Regional Office, Patna, CBSE.\n\n      Therefore, you are requested to depute eligible Practical Examiners not below the rank of PGT's in concerned subject from Local Authority/Institutions under intimation to the undersigned so that approval to be given by the Competent Authority. The list of schools with subject is enclosed herewith providing suitable early necessary arrangement.\n\n      Further, on completion of practical examinations, the bills of the said examiners duly countersigned by you may be forwarded to this office for processing & payment of the same through RTGS / NEFT.\n\n      Your cooperation in this regard will be highly appreciated.\n\n      Kindly treat it most urgent.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                doc.Add(p2);
                Paragraph p17 = new Paragraph(str: "Yours faithfully,     \n") { Alignment = Element.ALIGN_RIGHT };
                doc.Add(p17);
                doc.Add(rosign);
                Paragraph p18 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n RO PATNA, CBSE") { Alignment = Element.ALIGN_RIGHT };
                doc.Add(p18);
                doc.NewPage();
                doc.Add(header);
                doc.Add(footer);
                int count2 = 1;
                PdfPTable tbl1 = new PdfPTable(4);
                tbl1.WidthPercentage = 95;
                tbl1.SetWidths(new float[] { 20f, 250f, 250f, 50f });
                tbl1.HorizontalAlignment = 1;
                tbl1.AddCell(new Phrase(str: "SN", bold));
                tbl1.AddCell(new Phrase(str: "SCHOOL NO/SCHOOL NAME", bold));
                tbl1.AddCell(new Phrase(str: "SUBJECT", bold));
                tbl1.AddCell(new Phrase(str: "NO OF CAND", bold));
                while (dr2.Read())
                {
                    tbl1.AddCell(new PdfPCell(new Phrase(count2.ToString() + "\n", bold)));
                    tbl1.AddCell(new Phrase(str: "(" + dr2["SCH_NO"].ToString() + ") " + dr2["ABBR_NAME"].ToString().ToUpper() + "\nCONTACT: " + dr2["contact"].ToString()));
                    tbl1.AddCell(new Phrase(str: "(" + dr2["SUB"].ToString() + ") " + dr2["SUBNAME"].ToString().ToUpper()));
                    tbl1.AddCell(new Phrase(str: dr2["NOC"].ToString(), arial));
                    count2++;
                }
                doc.Add(tbl1);
                Paragraph p3 = new Paragraph(str: "Copy to:-\n\n", bold) { Alignment = Element.ALIGN_LEFT };
                doc.Add(p3);
                sqlcon2.Close();
                sqlcon2.Open();
                cmd3 = new SqlCommand("select distinct sch_no,abbr_name FROM [LETTERS].[dbo].[VOC_SCH2] where sch_cc='" + dr["sch_cc"].ToString() + "'", sqlcon2);
                SqlDataReader dr3 = cmd3.ExecuteReader(); 
                int count = 1;
                while (dr3.Read())
                {
                    Paragraph p7 = new Paragraph(count.ToString() + ".      The Principal, (" + dr3["sch_no"].ToString() + ") " + dr3["abbr_name"].ToString()) { Alignment = Element.ALIGN_LEFT };
                    doc.Add(p7);
                    count++;
                }
                Paragraph p6 = new Paragraph(str: "For Information with a request to contact the above mentioned Principal/City Coordinator for smooth conduct the practical examination 2022 in vocational/left-out subjects within time frame/rule of the Board.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                doc.Add(p6);
                doc.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
