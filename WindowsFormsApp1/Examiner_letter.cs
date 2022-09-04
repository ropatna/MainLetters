using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class Examiner_letter : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=comp2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public Examiner_letter()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            SqlCommand cmd2 = new SqlCommand("", sqlcon);
            string database = "";
            string database2 = "";
            string cls = Microsoft.VisualBasic.Interaction.InputBox("ENTER CLASS FOR WHICH LETTER HAS TO BE GENERATED", "INPUT CLASS", "");
            if(cls == "10"){
                database = "revexm10";
                database2 = "hepat1022t2";
            }
            if(cls == "12"){
                database = "revexm12";
                database2 = "hepat2022t2";
            }
            if (String.IsNullOrEmpty(textBox5.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where slno='" + textBox5.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                SqlConnection sqlcon2 = new SqlConnection(connectionString: "Data Source=comp2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
                sqlcon2.Open();
                cmd2 = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database2 + "] where HENO='" + dr["heno"].ToString() + "'", sqlcon2);
                SqlDataReader dr2 = cmd2.ExecuteReader();
                if (dr2.Read())
                {
                    Document doc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                    PdfWriter pwriter = PdfWriter.GetInstance(doc, new FileStream("E:\\abhi\\pdf\\EXMLETTER\\" + dr["sch_no"].ToString() + "_" + dr["HENO"].ToString() + "_" + dr["slno"].ToString() + "_EX2022.pdf", FileMode.Create));
                    var header = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\header.png");
                    var footer = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\FOOTER.png");
                    var rosign = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\rosignpng.png");
                    var rosign2 = iTextSharp.text.Image.GetInstance("E:\\abhi\\WindowsFormsApp1\\WindowsFormsApp1\\images\\rosignpng.png");
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
                    rosign.SetAbsolutePosition(470, 225);
                    rosign2.ScaleToFit(90f, 30f);
                    rosign2.SetAbsolutePosition(470, 370);
                    doc.Open();
                    void exm(Document pdoc)
                    {
                        pdoc.Add(header); //Adding Header
                        pdoc.Add(footer); //Adding Foter
                        iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12); 
                        Font bold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 12);
                        Font hbold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 12);
                        pdoc.AddTitle("EXAMINER Appointment Letter");
                        //
                        Paragraph p = new Paragraph("===============================================================================\n");
                        pdoc.Add(p);
                        Paragraph p2 = new Paragraph("CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                        pdoc.Add(p2);
                        Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                        pdoc.Add(p3);
                        Paragraph p4 = new Paragraph(str: "Ref.No.:RO-PTN/EXAM/SPOT/Term 2 2022/" + dr2["hesub"].ToString() + " - " + dr["SLNO"].ToString() + "/                                       Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p4);
                        Paragraph p33 = new Paragraph(str: "Examiner No: " + dr["SLNO"].ToString(),hbold) { Alignment = Element.ALIGN_RIGHT };
                        pdoc.Add(p33);
                        //Paragraph p5 = new Paragraph(str: dr["SLNO"].ToString() + "\nSchool No: " + dr["sch_no"].ToString() + "  " + dr["add1"].ToString() + "\n" + dr["add2"].ToString() + "\n" + dr["add3"].ToString() + "\n" + dr["add4"].ToString() + "\n" + dr["add5"].ToString() + "\nPIN: " + dr["pin"].ToString()) { Alignment = Element.ALIGN_LEFT };
                        Paragraph p5 = new Paragraph(str: dr["name"].ToString() + "\nSchool No: " + dr["sch_no"].ToString() + "\n" + dr["exabbrname"].ToString()) { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p5);
                        Paragraph p6 = new Paragraph(str: "\nSUB.:  APPOINTMENT LETTER AND INTIMATION OF VENUE FOR SPOT EVALUATION FOR EXAMINERS IN THE SUBJECT " + dr2["subname"].ToString() + "(" + dr2["hesub"].ToString() + ") OF CLASS " + dr2["heclass"].ToString() + " Term 2 Exam 2022") { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p6);
                        Paragraph p7 = new Paragraph(str: "\nDear Sir/Madam,\n\n") { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p7);
                        Paragraph p8 = new Paragraph(str: "      I am to inform you that you have  been  appointed as  Examiner for Evaluation  of  Answer Books of the subject mentioned above for 2022 class " + dr2["heclass"].ToString() + " Term 2 Examinations.  The Evaluation work will be done at the following address.  The appointment as well as the information to  this  effect  will be kept strictly confidential by you.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p8);
                        Paragraph p1 = new Paragraph(str: "\nNAME AND ADDRESS OF CHIEF NODAL SUPERVISOR/VENUE OF SPOT EVALUATION",hbold) { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p1);
                        PdfPTable tbl1 = new PdfPTable(2);
                        tbl1.HorizontalAlignment = 1;
                        tbl1.DefaultCell.Border = 0;
                        tbl1.AddCell(new Phrase(str: "\n" + dr2["CNSNAME"].ToString() + "  PRINCIPAL\n(" + dr2["CNS_SCHNO"].ToString() + ") " + dr2["CNSADD1"].ToString() + "\n" + dr2["CNSADD2"].ToString() + "\n" + dr2["CNSADD3"].ToString() + "\n" + dr2["CNSADD4"].ToString() + "\n" + dr2["CNSADD5"].ToString() + "\nPIN: " + dr2["CNSPIN"].ToString() + "\n"));
                        tbl1.AddCell(new Phrase(str: "\n\nMobile :" + dr2["CNSmobile"].ToString() + "\n\nEmail : " + dr2["email"].ToString() + "\n"));
                        pdoc.Add(tbl1);
                        Paragraph p9 = new Paragraph(str: "YOUR HEAD EXAMINER WILL BE :",hbold) { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p9);
                        PdfPTable tbl2 = new PdfPTable(2);
                        tbl2.HorizontalAlignment = 1;
                        tbl2.DefaultCell.Border = 0;
                        tbl2.AddCell(new Phrase(str: dr2["HENAME"].ToString() + "  <" + dr2["heno"].ToString() + ">\n(" + dr2["SCH_NO"].ToString() + ") " + dr2["HEADD1"].ToString() + "\n" + dr2["HEADD2"].ToString() + "\n" + dr2["HEADD3"].ToString() + "\n" + dr2["HEADD4"].ToString() + "\n" + dr2["HEADD5"].ToString() + "\nPIN: " + dr2["HEPIN"].ToString() + "\n"));
                        tbl2.AddCell(new Phrase(str: "\nMobile : " + dr2["MOBILE"].ToString() + "\n\nEmail:  " + dr2["heemail"].ToString() + "\n"));
                        pdoc.Add(tbl2);
                        Paragraph p10 = new Paragraph(str: "      Further, tentative date for commencement of the Spot Evaluation  is " + dr2["sdate"].ToString() + " Thus, you are requested to please contact your CNS & HE before the scheduled date accordingly for evaluation of answer books.  You will work  under the Supervision and Control  of H.E. & Chief Nodal Supervisor and the  time  of Evaluation will be decided by him / her.  Please bring with you a red ink ball pen for corrections.  \n") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p10);
                        pdoc.NewPage();
                        pdoc.Add(header);
                        pdoc.Add(footer);
                        Paragraph p11 = new Paragraph(str: "\n      On  the  basis  of the  past  experience  it  has been observed  that the teachers  are  normally  sending  their  acceptance  but  when  the actual Evaluation work  commences, they  do not report  for  duty which  jeopardises  the  schedule of Evaluation Work.  It may be noted that as per provisions of affiliation / examination bye - laws of the Board, the assignment offered  by  the Board in regard with conduct of examination, evaluation etc. are  mandatory / obligatory on the part of each affiliated institutions and their teachers.  \n\n      Hence non  reporting  for evaluation duties as assigned by the Board  may lead to strict action as deemed fit.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p11);
                        Paragraph p12 = new Paragraph(str: "\n      FURTHER, AS PER PAST EXPERIENCE IT HAS BEEN OBSERVED THAT QUALITY OF EVALUATION DONE AT SOME OF THE SPOT EVALUATION CENTRES WAS NOT FOUND SATISFACTORY AND ALSO NOT UPTO THE  DESIRED  LEVEL, AS  LARGE NUMBER OF MISTAKE  CASES WERE  DETECTED DURING THE COURSE OF SCRUTINY  ON ACCOUNT OF  EVALUATION  OF  PREVIOUS EXAMINATION.  NON PERFORMING  THE DUTIES  PROPERLY  RAISED  QUESTIONS  ON  THE CREDIBILITY  ON WORKING OF EXAMINERS WHO PARTICIPATED IN THE EVALUATION  WORK, AS YOU ARE AWARE THAT  THE  STUDENTS  LOOK  AT THESE EVALUATION  AS  A   FINAL EVALUATION OF THEIR ACADEMIC PERFORMANCE.  THE COMPETENT AUTHORITY OF THE BOARD HAS  TAKEN  SERIOUS  VIEW  ON  LARGE  NO.OF MISTAKES FOUND  DURING  PREVIOUS YEAR.  ALSO FROM EXAM  2012 THE  STUDENTS / EXAMINEES CAN  TAKE  PHOTOCOPY  OF THEIR ANSWER SHEETS UNDER RTI  ACT 2005 AS  PER ORDERS OF THE HON'BLE SUPREME COURT OF  INDIA.  THEREFORE,  IT  IS  REQUESTED  THAT  PROPER ATTENTION  TOWARDS  EVALUATION  SHOULD  BE   GIVEN  AND  ANSWER  BOOKS  OF   THE  SUBJECT BE EVALUATED IN PERFECT MANNER STRICTLY IN ACCORDANCE WITH THE MARKING SCHEME.") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p12);
                        Paragraph p13 = new Paragraph(str: "\n      APART FROM ABOVE, THE EXAMINERS  EVALUATING  THE  ANSWER BOOKS OF THE  MEDIUM OTHER THAN THE ONE THEY ARE TEACHING IN, MAY HAVE  SOME DIFFICULTY  IN UNDERSTANDING THE  ANSWER   WHICH   MAY   LEAD   TO  WRONG  AWARD  OF  MARKS.  THEREFORE, FOR  DOING  FULL  JUSTICE TO  THE  STUDENTS, EXAMINERS MAY ENSURE EVALUATION OF THE ANSWER BOOK WITH THE SAME MEDIUM IN WHICH THEY ARE TEACHING.\n\n      It may be noted that each examiner should devote 7-8  hours  daily  for evaluation of 30 - 35 answer books and ensure that all the answer books evaluated strictly in accordance With the marking scheme provided by the Board and question paper set used by the examinee.") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p13);
                        pdoc.NewPage();
                        pdoc.Add(header);
                        pdoc.Add(footer);
                        Paragraph p14 = new Paragraph(str: "\n       While evaluating the answer books, the examiner concerned  must  ensure that where multiple sets of question papers are  in the subject  i.e. 3 sets of  question  papers, only one  set's answer book have to be taken  for  evaluation as far as possible.") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p14);
                        Paragraph p34 = new Paragraph(str: "\n       It will be the responsibility of the examiner to verify and ascertain that the answer books in hand does not belong to any of the other set other than the one which he / she has been allotted for evaluation.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p34);
                        Paragraph p35 = new Paragraph(str: "\n       However, in case of the multiple  sets  of  question paper's  answer books the examiner should have the knowledge of the other sets also which is not difficult being the part of  the same syllabus and the academic subject.  Thus, the  evaluation be done accordingly.") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p35);
                        Paragraph p15 = new Paragraph(str: "\n       REMUNERATION\\CONVEYANCE IS ADMISSIBLE AS PER LAST YEAR OR AS PER CS GUIDELINES-2022:\n        =================================================================================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 10)) { Alignment = Element.ALIGN_LEFT };
                        pdoc.Add(p15);
                        PdfPTable tbl3 = new PdfPTable(2);
                        tbl3.HorizontalAlignment = 0;
                        tbl3.DefaultCell.Border = 0;
                        tbl3.AddCell(new Phrase("Evaluation of A/B of theory paper"));
                        tbl3.AddCell(new Phrase(""));
                        tbl3.AddCell(new Phrase("    For Class XII"));
                        tbl3.AddCell(new Phrase(":  @Rs.30/- per answer book"));
                        tbl3.AddCell(new Phrase("    For Class X"));
                        tbl3.AddCell(new Phrase(":  @Rs.25/- per answer book"));
                        tbl3.AddCell(new Phrase("    Conveyance"));
                        tbl3.AddCell(new Phrase(":  @Rs.250/- per day"));
                        tbl3.AddCell(new Phrase("    Refreshment charges at the\n    Spot Evaluation Center"));
                        tbl3.AddCell(new Phrase(":  @Rs. 75/- per day"));
                        pdoc.Add(tbl3);
                        Paragraph p16 = new Paragraph(str: "\n      Your acceptance as Examiner may please be sent to the concerned Head  Examiner immediately on receipt  of this letter.  You are also  required  to  produce a relieving certificate at the time joining / releiving issued accordingly.\n\n      Hope, you would pay proper attention towards most important  work of whole examination system and  ensure  error  free  evaluation as per  norms / guidelines of the Board in the larger academic interest of the students.") { Alignment = Element.ALIGN_JUSTIFIED };
                        pdoc.Add(p16);
                        Paragraph p22 = new Paragraph(str: "\nYours faithfully,     \n") { Alignment = Element.ALIGN_RIGHT };
                        pdoc.Add(p22);
                    }
                    exm(doc);
                    doc.Add(rosign);
                    Paragraph p24 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\nCBSE, PATNA") { Alignment = Element.ALIGN_RIGHT };
                    doc.Add(p24);
                    doc.NewPage();
                    doc.Add(header);
                    doc.Add(footer);
                    Paragraph p25 = new Paragraph(str: "\n\n\nCopy for information and necessary action:\n\nThe Principal,(" + dr["SCH_NO"].ToString() + ")\n" + dr["exabbrname"].ToString()) { Alignment = Element.ALIGN_LEFT };
                    doc.Add(p25);
                    Paragraph p26 = new Paragraph(str: "\n\n       With  the  request  to  relieve  above  VP/PGT/PET/TGT  from the School  to  act as HEAD EXAMINER/SUB EXAMINER at the above Spot Evaluation Centre  for 2022 Term 2 Exam as per  the  undertaking/data forwarded  by you.  The status of releiving of the teacher concerned must be confirmed to the  undersigned  on priority as it is mandatory.\n\n       AS PER CLAUSE 12.2.10 OF AFFILIATION BYE-LAWS.  THE BOARD MAY IMPOSE ALL OR ANY OF THE PENALTIES IN CLAUSE  12.1.1 TO 12.1.9 FOR NOT NOMINATING AND RELIEVING TEACHERS FOR EVALUATION OF ANSWER SCRIPTS OF THE BOARDS EXAMINATION AND OTHER ANCILLARY ACTIVITIES AS PER REQUIREMENTS OF THE BOARD.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    doc.Add(p26);
                    doc.Add(rosign2);
                    Paragraph p27 = new Paragraph(str: "\n\n\n(JAGADISH BARMAN) \nREGIONAL OFFICER\nCBSE, PATNA") { Alignment = Element.ALIGN_RIGHT };
                    doc.Add(p27);
                    doc.Close();
                }
                sqlcon2.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
