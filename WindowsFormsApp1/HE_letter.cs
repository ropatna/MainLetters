using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class HE_letter : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=COMP2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public HE_letter()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            SqlCommand cmd2 = new SqlCommand("", sqlcon);
            string database = "";
            string database2 = "";
            string cls = Microsoft.VisualBasic.Interaction.InputBox("ENTER CLASS FOR WHICH LETTER HAS TO BE GENERATED", "INPUT CLASS", "");
            if (cls == "10")
            {
                database = "revexm10";
                database2 = "hepat1022t2";
            }
            if (cls == "12")
            {
                database = "revexm12";
                database2 = "hepat2022t2";
            }
            if (String.IsNullOrEmpty(textBox4.Text))
            {
                cmd2 = new SqlCommand("select * FROM [LETTERS].[dbo].["+ database2 + "]", sqlcon);
            }
            else
            {
                cmd2 = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database2 + "] where heno='" + textBox4.Text + "'", sqlcon);
            }
            SqlDataReader dr2 = cmd2.ExecuteReader();
            while (dr2.Read())
            {
                SqlConnection sqlcon2 = new SqlConnection(connectionString: "Data Source=comp2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
                sqlcon2.Open();
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where HENO='" + dr2["heno"].ToString() + "'", sqlcon2);
                SqlDataReader dr = cmd.ExecuteReader();
                Document pdoc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(pdoc, new FileStream("E:\\abhi\\pdf\\HELETTER\\" + dr2["sch_no"].ToString() + "_" + dr2["HENO"].ToString() + "_HE2022.pdf", FileMode.Create));
                var header = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\header.png");
                var footer = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\FOOTER.png");
                var rosign = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\rosignpng.png");
                var header2 = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\acceptance_header.png");
                header2.ScaleToFit(900f, 60f);
                header2.Alignment = 1;
                header.ScaleToFit(900f, 60f);
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(90f, 30f);
                rosign.SetAbsolutePosition(470, 110);
                pdoc.Open();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Font arial = FontFactory.GetFont("Arial", 10);
                Font bold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 10);
                Font hbold = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                pdoc.AddTitle("HE Appointment Letter");
                //
                Paragraph p = new Paragraph("===============================================================================\n");
                pdoc.Add(p);
                Paragraph p2 = new Paragraph("CONFIDENTIAL", bold) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p2);
                Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p3);
                Paragraph p4 = new Paragraph(str: "Ref.No.:RO-PTN/EXAM/SPOT/Term 2 2022/" + dr2["hesub"].ToString() + " - " + dr2["heno"].ToString() + "/                                          Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p4);
                Paragraph p34 = new Paragraph(str: "He No." + dr2["heno"].ToString(),hbold) { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p34);
                Paragraph p5 = new Paragraph(str: dr2["HENAME"].ToString().ToUpper() + ", " + dr2["POST"].ToString().ToUpper() + "\n(" + dr2["sch_no"].ToString() + ") " + dr2["headd1"].ToString().ToUpper() + "\n" + dr2["headd2"].ToString().ToUpper() + "\n" + dr2["headd3"].ToString().ToUpper() + "\n" + dr2["headd4"].ToString().ToUpper() + ", " + dr2["headd5"].ToString().ToUpper() + ", PIN: " + dr2["hepin"].ToString()) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p5);
                Paragraph p6 = new Paragraph(str: "\nSUB.:  APPOINTMENT LETTER AND INTIMATION FOR SPOT EVALUATION FOR HEAD EXAMINER OF SUBJECT " + dr2["subname"].ToString() + "(" + dr2["hesub"].ToString() + ") OF CLASS " + dr2["heclass"].ToString() + " Term 2 Main Exam 2022") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p6);
                Paragraph p7 = new Paragraph(str: "\nDear Sir/Madam,") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p7);
                Paragraph p8 = new Paragraph(str: "      This is in  continuation  to  this  office  letter  no. RO(PTN)/CONF./H.E./" + dr2["hesub"].ToString() + "/" + dr2["HENO"].ToString() + "/2022 Term 2   and  subsequently  your  acceptence for appointment/assignment of Head Examiner in the subject " + dr2["subname"].ToString().ToUpper() + "(" + dr2["hesub"].ToString() + "),\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p8);
                Paragraph p38 = new Paragraph(str: "      In this connection, I am to inform you that the Evaluation in the subject  " + dr2["subname"].ToString().ToUpper() + "(" + dr2["hesub"].ToString() + ") is scheduled to   commence  on dt. " + dr2["SDATE"].ToString() + " at your Centre and the work shall have to  be  completed  before dt. " + dr2["EDATE"].ToString() + "  positively.",hbold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p38);
                Paragraph p9 = new Paragraph(str: "      You must be aware that the Evaluation Work is very sensitive in nature and the future of the students is going to be  affected by the  manner in which the Evaluation is done.  Therefore, it is requested that due care, attention and complete  dedication should be given to  the  Evaluation  of  Answer  Books  by ensuring accuracy & uniformity as per  Marking  Scheme  and performance.  Also it may be noted that  the  work  assigned  to  you is strictly time  bound & confidential and the entire Post - Examination work will have a direct bearing for timely completion of Evaluation Work and subsequently scheduled declaration  of result.  You are therefore requested to complete the task including coordination work within prescirbed period i.e. 08-10 days time(maximum) positively.\n      Though  the  examiners  have  been  appointed  by the Regional Office, Patna, CBSE.  However, Head Examiner can make alternate arrangements in consultation with the CNS when  the  requisite no. of  examiners appointed by the Board do not report for evaluation with a valid reason.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p9);
                Paragraph p10 = new Paragraph(str: "      The Head  Examiners  will have to ensure that  such appointment is made strictly as per the qualifications of examiners given under rule 55(ii) of exam by  laws  and as per  guidelines  of  Spot Evaluation Term 2 Main Exam 2022.However details of the appointments made at local level be sent  to  the  Regional  Office, Patna, CBSE  for necessary record and updations.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p10);
                pdoc.NewPage();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p11 = new Paragraph(str: "      The name(s) and  addresses of the Examiners appointed for the Evaluation work at your Centre are as follows  :-\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p11);
                PdfPTable tbl1 = new PdfPTable(3);
                tbl1.WidthPercentage = 95;
                tbl1.SetWidths(new float[] { 20f, 100f, 250f });
                tbl1.HorizontalAlignment = 1;
                int count = 1;
                tbl1.AddCell(new Phrase(str: "SN", hbold));
                tbl1.AddCell(new Phrase(str: "Examiner Name\nExaminer No.", hbold));
                tbl1.AddCell(new Phrase(str: "Examiner School No. and Name\nContact No.", hbold));
                while (dr.Read())
                {
                    if (count <= 20)
                    {
                        tbl1.AddCell(new PdfPCell(new Phrase(count.ToString() + "\n", bold)) { Rowspan = 2 });
                        //tbl1.AddCell(new Phrase(str: count.ToString()+"\n", arial));
                        tbl1.AddCell(new Phrase(str: dr["NAME"].ToString().ToUpper(), bold));
                        tbl1.AddCell(new Phrase(str: "(" + dr["SCH_NO"].ToString() + ") " + dr["EXABBRNAME"].ToString().ToUpper(), bold));
                        tbl1.AddCell(new Phrase(str: "EXAMINER NO:  " + dr["SLNO"].ToString(), arial));
                        tbl1.AddCell(new Phrase(str: "CONTACT:  " + dr["MOBILE"].ToString() + "       PRINCIPAL CONTACT:  " + dr["expMOBILE"].ToString(), arial));
                    }
                    count++;
                }
                pdoc.Add(tbl1);
                pdoc.NewPage();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                if (count >= 21)
                {
                    PdfPTable tbl2 = new PdfPTable(3);
                    tbl2.WidthPercentage = 95;
                    tbl2.SetWidths(new float[] { 20f, 100f, 250f });
                    tbl2.HorizontalAlignment = 1;
                    int count2 = 1;
                    tbl2.AddCell(new Phrase(str: "SN", hbold));
                    tbl2.AddCell(new Phrase(str: "Examiner Name\nExaminer No.", hbold));
                    tbl2.AddCell(new Phrase(str: "Examiner School No. and Name\nContact No.", hbold));
                    sqlcon2.Close();
                    sqlcon2.Open();
                    cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where HENO='" + dr2["heno"].ToString() + "'", sqlcon2);
                    SqlDataReader dr3 = cmd.ExecuteReader();
                    while (dr3.Read())
                    {
                        if (count2 >= 21)
                        {
                            tbl2.AddCell(new PdfPCell(new Phrase(count2.ToString() + "\n", bold)) { Rowspan = 2 });
                            tbl2.AddCell(new Phrase(str: dr3["NAME"].ToString().ToUpper(), bold));
                            tbl2.AddCell(new Phrase(str: "(" + dr3["SCH_NO"].ToString() + ") " + dr3["EXABBRNAME"].ToString().ToUpper(), bold));
                            tbl2.AddCell(new Phrase(str: "EXAMINER NO:  " + dr3["SLNO"].ToString(), arial));
                            tbl2.AddCell(new Phrase(str: "CONTACT:  " + dr3["MOBILE"].ToString() + "       PRINCIPAL CONTACT:  " + dr3["expMOBILE"].ToString(), arial));
                        }
                        count2++;
                    }
                    pdoc.Add(tbl2);
                }
                if (count <= 20)
                {
                    string space = "\n\n";
                    Paragraph p12 = new Paragraph(str: "\n\n      The   teachers  appointed as  Examiners  have already been informed separately to contact you before the start of Evaluation Work on " + dr2["sdate"].ToString() + ".   You are also requested to confirm the status from all the Examiners and inform them the date of commencement of Evaluation Work to enable them to attend the task as per schedule and also ensure  that  maximum no. of Examiners  turn up for evaluation." + space + "  THE GUIDELINE FOR SPOT EVALUATION 2022 WILL BE FORWARDED TO  YOU  IN DUE TIME FOR KIND  PERUSAL  AND  STRICT  COMPLIANCE  OF  THE  INSTRUCTIONS  AND   MAKING  ALL NECESSARY ARRANGEMENTS ACCORDINGLY.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p12);
                }
                else
                {
                    string space = "";
                    Paragraph p12 = new Paragraph(str: "\n\n      The   teachers  appointed as  Examiners  have already been informed separately to contact you before the start of Evaluation Work on " + dr2["sdate"].ToString() + ".   You are also requested to confirm the status from all the Examiners and inform them the date of commencement of Evaluation Work to enable them to attend the task as per schedule and also ensure  that  maximum no. of Examiners  turn up for evaluation." + space + "  THE GUIDELINE FOR SPOT EVALUATION 2022 WILL BE FORWARDED TO  YOU  IN DUE TIME FOR KIND  PERUSAL  AND  STRICT  COMPLIANCE  OF  THE  INSTRUCTIONS  AND   MAKING  ALL NECESSARY ARRANGEMENTS ACCORDINGLY.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p12);
                }
                Paragraph p13 = new Paragraph(str: "\nNAME AND ADDRESS OF CHIEF NODAL SUPERVISOR/VENUE OF SPOT EVALUATION IS:",hbold) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p13);
                PdfPTable tbl3 = new PdfPTable(2);
                tbl3.HorizontalAlignment = 1;
                tbl3.DefaultCell.Border = 0;
                tbl3.AddCell(new Phrase(str: "\n" + dr2["CNSNAME"].ToString() + "\nPOST-  PRINCIPAL\n(" + dr2["CNSSCH_NO"].ToString() + ") " + dr2["CNSADD1"].ToString() + "\n" + dr2["CNSADD2"].ToString() + "\n" + dr2["CNSADD3"].ToString() + "\n" + dr2["CNSADD4"].ToString() + "\n" + dr2["CNSADD5"].ToString() + "\nPIN: " + dr2["CNSPIN"].ToString() + "\n",hbold));
                tbl3.AddCell(new Phrase(str: "\n\nMobile :" + dr2["CNSmobile"].ToString() + "\n\nEmail : " + dr2["email"].ToString() + "\n",hbold));
                pdoc.Add(tbl3);
                pdoc.NewPage();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p14 = new Paragraph(str: "      APART FROM ABOVE, THE EXAMINERS  EVALUATING  THE  ANSWER BOOKS OF THE MEDIUM  OTHER THAN THE ONE THEY ARE TEACHING IN, MAY HAVE  SOME DIFFICULTY  IN UNDERSTANDING  THE  ANSWER   WHICH   MAY   LEAD   TO  WRONG  AWARD  OF  MARKS.  THEREFORE, FOR DOING FULL JUSTICE TO  THE  STUDENTS, ALL HEAD EXAMINERS SHOULD STRICTLY  CHECK  AND  ENSURE THAT NO ANSWER BOOK IS EVALUATED BY EXAMINERS WHO ARE TEACHING IN THE MEDIUM OTHER THAN THE ONE USED IN THE ANSWER BOOK i.e.  ALL ANSWER BOOKS SHOULD  BE  CHECKED BY THE  SUBJECT EXAMINERS OF THE SAME MEDIUM.\n\n      Further  Head  Examiner  should  check  whether  any  Answer Books of Physically Challanged children(Spastic, Blind, Physically  Handicapped  and Dyslexic children) have been erroneously received alongwith the  Answer  Booksof other candidates.  If the Answer Books of Physically Challanged children are found mixed with the Answer Books of other candidates, these must be  immediately returned to the undersigned UNEVALUATED through sealed insured Speed Post.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p14);
                Paragraph p15 = new Paragraph(str: "\n      The Head Examiners should ensure that all the  examiners appointed  for evaluation must report  for duty at the specified time and stay upto 7 - 8  hours for evaluating 30 - 35 answer - books daily.\n      It may be noted, where multiple sets  of  question  papers  are  in the subject  i.e. 3 sets  of question  papers, only one  set has to be assigned for evaluation to  individual  examiner as far as possible.  It  will  be   the responsibility of the concerned examiner to verify and ascertain that the answer books in hand does not belong to any of the other set, other than the one which he / she  has  been  allotted  for  evaluation.However, the  examiner concerned  evaluating  one  of  the set of the question paper should have the knowledge of  the other sets also which is not difficult being   part of  the  same  syllabus  and the academic subject.  Thus, it may also be ascertained  that all the answer  books be  evaluated  by  the correct set of question paper used by the examinee in the examination.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p15);
                Paragraph p16 = new Paragraph(str: "\n      Hope, you would pay  proper  attention  for conducting  and  completing  error free  evaluation as per  norms  of  the  Board.  Your cooperation  for sucessfull completion of the work is highly appreciated.\n\n      You may submit your consolidated bill of remuneration etc.  within 15 days after the completion of  the  evaluation  with  full  details  of  Bank Address, IFSC code, name of A / C holder to  the  Assistant Secretary(Accounts) directly for settlement of the bills and release of the payments.\n\n      Thanking you,") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p16);
                Paragraph p17 = new Paragraph(str: "Yours faithfully,     \n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p17);
                pdoc.Add(rosign);
                Paragraph p18 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p18);

                pdoc.Close();
                sqlcon2.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
