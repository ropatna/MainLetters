using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class CNS : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=COMP2\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public CNS()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            string database = Microsoft.VisualBasic.Interaction.InputBox("ENTER NAME OF DATABASE FROM WHICH LETTER HAS TO BE GENERATED", "INPUT DATABASE NAME", "CNS2022SQL");
            if (String.IsNullOrEmpty(textBox3.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where cns_schno='" + textBox3.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Document pdoc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(pdoc, new FileStream("E:\\abhi\\pdf\\CNS\\" + dr["cns_schno"].ToString() + "_CNS.pdf", FileMode.Create));
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
                rosign.ScaleToFit(100f, 40f);
                rosign.SetAbsolutePosition(470, 330);
                pdoc.Open();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                iTextSharp.text.Font bold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 12);
                pdoc.AddTitle("CNS Acceptance Letter");
                //
                Paragraph p = new Paragraph("===============================================================================\n");
                pdoc.Add(p);
                Paragraph p2 = new Paragraph("CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p2);
                Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p3);
                Paragraph p4 = new Paragraph(str: "No.:CBSE/RO(PTN)/CONF./NODAL/" + dr["cns_schno"].ToString() + "/TERM 2-2022/                                            Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p4);
                Paragraph p5 = new Paragraph(str: "School Code: " + dr["cns_schno"].ToString() + "\n" + dr["CNSNAME"].ToString() + ", PRINCIPAL\n" + dr["cnsadd1"].ToString() + "\n" + dr["cnsadd2"].ToString() + "\n" + dr["cnsadd2"].ToString() + "\n" + dr["cnsadd3"].ToString() + "\n" + dr["cnsadd4"].ToString() + "\n" + dr["cnsadd5"].ToString() + "\nPIN: " + dr["cnspin"].ToString() + "\n\nSUB. : APPOINTMENT OF CHIEF NODAL SUPERVISOR(CNS) FOR SPOT EVALUATION FOR AISSE/AISSCE-2022 (TERM 2).\n\nSir/Madam,\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p5);
                Paragraph p6 = new Paragraph(str: "      This is to inform you that the Competent Authority of the Board is pleased to appoint you as Chief Nodal Supervisor at your school / nodal center to ensure proper supervision and timely completion of Evaluation of Answer books with perfect  accuracy. The Examination of AISSE/AISSCE-2022 TERM 2 is scheduled to be held from 26-04-2022 to 15-06-2022. The Chief Nodal Supervisor(CNS) would be a Principal of the school and Vice Principal/PGT be  appointed  as  Head Examiner  under his/her supervision. The  Chief Nodal Supervisor would supervise the  work of  Head Examiners appointed under him/her.  The Head  Examiner would do the Evaluation Work in the School of Chief Nodal Supervisor.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p6);
                Paragraph p7 = new Paragraph(str: "\n      The evaluation will be done at your school and the timing of evaluation will be 8 to 10 hours minimum for HE / AHE(Evaluation) and AHE(Coord) / Evaluators in accordance with the schedule.  Since evaluation is a time bound activity, it is, therefore, desirable that adequate time be devoted by each teacher to complete the Evaluation not only  in time but should  be  in a perfect and professional manner under  each Chief Nodal Supervisor, Head  Examiners  and  about  12 - 16 Examiners are also being appointed with each Head Examiner.  For evaluation work one seperate room  be provided to  each Head Examiner. The tentative  dates for Evaluation are between 02-05-2022  to  02-07-2022. The  AHE(Coord.) who will do coordination work of the Answer Books(uploading of marks) and AHE(Evaluation) will be  appointed by  the Head Examiners as per provisions / guidelines amongst PGTs  and TGTs for class XII and X respectively in consultation with you.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p7);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p8 = new Paragraph(str: "      The Evaluation/Coordination work will be done strictly as per instructions given in the guidelines for Spot Evaluation.\n      The  delivery  of  Answer  Books  will be made approximately after 04-06 days from the commencement of the Examination of  the   concerned  subject and evaluation shall be got started immediately on  receipt of Answer Books by you at  Nodal Centre. The appointment as well as the information in this regard will be kept secret.\n      As a Chief  Nodal Supervisor, you will  have to  ensure that  the Head Examiners and AHE / Examiners working under you are Evaluating the Answer Scripts strictly in accordance with the  Marking  Scheme, leaving  no  scope  for  any allegations etc.except the real merit of the Examinees.  It may also be ensured that the AHE(Evaluation) / AHE(Coord) will perform their duties with utmost care and  sense of responsibility and will see  to it that  evaluation / coordination work done by them is absolutely without any error and prejudice.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p8);
                Paragraph p24 = new Paragraph(str: "\nREMUNERATION\\CONVEYANCE IS ADMISSIBLE AS PER LAST YEAR OR AS PER SPOT GUIDELINES 2022 (Term 2):\n        =======================================================================================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 10)) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p24);
                PdfPTable tbl1 = new PdfPTable(2);
                tbl1.HorizontalAlignment = 0;
                tbl1.DefaultCell.Border = 0;
                tbl1.AddCell(new Phrase("1. Chief Nodal Supervisor"));
                tbl1.AddCell(new Phrase(":  Rs.2000/- per HE ( Remuneration is\n   per HE for entire duration of\n   evaluation, subject to maximum of\n   Rs. 10000.) "));
                tbl1.AddCell(new Phrase("2. Head Examiner"));
                tbl1.AddCell(new Phrase(":  Rs.1000/- per day"));
                tbl1.AddCell(new Phrase("3. Addl. Head Examiner"));
                tbl1.AddCell(new Phrase(":  Rs.900/- per day"));
                tbl1.AddCell(new Phrase("4. Evaluating A/Bs (Theory paper)"));
                tbl1.AddCell(new Phrase(""));
                tbl1.AddCell(new Phrase("    i)   For Senior Exams"));
                tbl1.AddCell(new Phrase(":  Rs.30/- per A/B"));
                tbl1.AddCell(new Phrase("    ii)  For Secondary Exams"));
                tbl1.AddCell(new Phrase(":  Rs.25/- per A/B"));
                tbl1.AddCell(new Phrase("5. Conveyance"));
                tbl1.AddCell(new Phrase(":  Rs.250/- per day (For CNS/HE/AHE\n   (Evaluation/Coordination)/Examiner)"));
                tbl1.AddCell(new Phrase("6. AHE(Coord) ( Class XII)"));
                tbl1.AddCell(new Phrase(":  Rs.7.5/-  per A/B"));
                tbl1.AddCell(new Phrase("    ( Class X  )"));
                tbl1.AddCell(new Phrase(":  Rs.6.25/- per A/B \n   AHE(Coord) will be paid remuneration\n   as above and fixed conveyance of Rs.\n   250 for finalizing minimum 80 Answer\n   Books in a day."));
                tbl1.AddCell(new Phrase("7. Admissibility of clerical and Class IV\n    staff (each one)"));
                tbl1.AddCell(new Phrase(""));
                tbl1.AddCell(new Phrase("    i)  Clerk (Regular employee of the\n    School)"));
                tbl1.AddCell(new Phrase(":  Rs.200/- per day"));
                tbl1.AddCell(new Phrase("    ii) Class IV (from the school)"));
                tbl1.AddCell(new Phrase(":  Rs.100/- per day"));
                tbl1.AddCell(new Phrase("8.  Refreshment charges Rs. 75/- per\n    head per day"));
                tbl1.AddCell(new Phrase(""));
                pdoc.Add(tbl1);
                Paragraph p9 = new Paragraph(str: "      FURTHER, AS PER PAST EXPERIENCE IT HAS BEEN OBSERVED THAT QUALITY OF EVALUATION DONE AT SOME OF THE SPOT EVALUATION CENTRES WAS NOT FOUND SATISFACTORY AND ALSO NOT UPTO THE  DESIRED LEVEL AS PER MARKING SCHEME.  AS LARGE NUMBER OF MISTAKE CASES WERE DETECTED DURING THE COURSE OF SCRUTINY ON ACCOUNT OF EVALUATION OF PREVIOUS EXAMINATION, WHICH WERE NOT DONE",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p9);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p28 = new Paragraph(str: "PROPERLY & RAISED  QUESTIONS ON THE CREDIBILITY ON WORKING OF HEAD EXAMINERS AND EXAMINERS WHO PARTICIPATED IN THE EVALUATION WORK.\n\n", bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p28);
                Paragraph p10 = new Paragraph(str: "      AS YOU ARE AWARE THAT STUDENTS LOOK AT THESE EXAMINATION AS A FINAL EVALUATION OF THEIR ACADEMIC PERFORMANCE. THE COMPETENT AUTHORITY OF THE BOARD HAS TAKEN IT SERIOUSLY OWING TO LARGE NO. OF MISTAKES DURING RE-EVALUATION. ALSO FROM EXAM 2012 STUDENTS / EXAMINEES MAY TAKE PHOTOCOPY OF THEIR ANSWER SHEETS UNDER RTI ACT 2005 AS PER ORDERS OF THE HON'BLE SUPREME COURT OF INDIA ALSO CAN GET THEIR A/Bs EVALUATED FOR OPTED SUBJECTS, THEREFORE  IT IS REQUESTED THAT PROPER ATTENTION TOWARDS EVALUATION SHOULD BE GIVEN AND ANSWER BOOKS OF THE SUBJECT BE EVALUATED  IN PERFECT MANNER STRICTLY IN ACCORDANCE WITH THE MARKING SCHEME.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p10);
                Paragraph p11 = new Paragraph(str: "      APART FROM ABOVE, THE  EXAMINERS  EVALUATING  THE  ANSWER BOOKS OF THE MEDIUM  OTHER THAN THE ONE THEY ARE TEACHING IN, MAY HAVE  SOME DIFFICULTY  IN UNDERSTANDING  THE  ANSWER  WHICH  MAY  LEAD  TO  WRONG  EVALUATION  OF MARKS. THEREFORE, FOR DOING FULL JUSTICE TO  THE  STUDENTS, ALL HEAD EXAMINERS SHOULD STRICTLY  CHECK  AND  ENSURE THAT NO ANSWER BOOK IS EVALUATED BY EXAMINERS WHO ARE TEACHING IN THE MEDIUM OTHER THAN THE ONE USED IN THE ANSWER BOOK i.e.ALL ANSWER BOOKS SHOULD  BE  CHECKED BY THE  SUBJECT EXAMINERS OF THE SAME MEDIUM.\n\n      Further, Head  Examiner  should  check  whether  any  Answer Books of Physically Challanged children(Spasctic, Blind, Physically  Handicapped  and Dyslexic candidate) have been erroneously received alongwith the  Answer Books of other candidates.  If the Answer Books of Physically Challanged candidate are found mixed with the Answer Books of other candidates, these  be  immediately returned to the undersigned without being  evaluated  through  sealed insured Speed Post parcel and information must be acknowledged to undersigned through email.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p11);
                Paragraph p12 = new Paragraph(str: "      After evaluation of all  the  Answer  Books  pertaining  to  the  H.E.s concerned be serialised in ascending order century wise  and should be packed / sealed in  respective Answer Book bags. \n\n      It may  also  be  noted  that  the  Evaluation  is  the  most  important part  of the whole Examination system and it determines  the  future career of the  students.  Therefore you are  requested  to  take  every possible care to ensure objective and  judicious  Evaluation  to  safeguard the interest of the students and also to avoid any future complications / allegations.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p12);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p13 = new Paragraph(str: "      A copy  of the guidelines for  Spot Evaluation 2022  Exam Will be  e-mailed to you  in due course for your reference and strict compliance.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p13);
                Paragraph p14 = new Paragraph(str: "\n      You are also requested  to  expedite the  acceptance of  Head  Examiners appointed under your supervisions.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p14);
                Paragraph p25 = new Paragraph(str: "\n      If  there is any change  on  account  of transfer, ward  appearing  in the  subject  etc.the  same  may  be informed immediately with replacement as per experience / eligibility, so  that necessary changes be done accordingly in time.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p25);
                Paragraph p26 = new Paragraph(str: "\n      You are requested to send the bill of HE(s) including yourself within 15 days  after  completion of work, with following details of your Bank A / C. All bills pertaining to HE(s) of various subject will be submitted by you  to the undersigned at a time.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p26);
                PdfPTable tbl2 = new PdfPTable(1);
                tbl2.HorizontalAlignment = 1;
                tbl2.DefaultCell.Border = 0;
                tbl2.AddCell(new Phrase("a) A/C holder name\nb) IFSC code\nc) A/C No.\nd) Bank name branch & location\n\n"));
                pdoc.Add(tbl2);
                Paragraph p15 = new Paragraph(str: "      Your acceptance of the assignment offered by the Board in the enclosed proforma  duly  completed  in all  respects  should  reach  to the undersigned within 03 days after receipt of this letter through Speed Post / Email: ropatna.cbse@nic.in / abcell.ropatna@cbseshiksha.in.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p15);
                Paragraph p16 = new Paragraph(str: "      Yours faithfully,\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p16);
                pdoc.Add(rosign);
                Paragraph p17 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\nRO PATNA, CBSE") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p17);
                pdoc.NewPage();
                pdoc.Add(header2); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p18 = new Paragraph(str: "FORM OF ACCEPTANCE FOR CHEIF NODAL SUPERVISOR FOR SR./SEC. SCHOOL EXAMINATION 2022 TERM 2\n-----------------------------------------------------------------------------------------\n(To be sent by Speed Post/Email or through Messanger)\nIMMEDIATE & CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p18);
                Paragraph p19 = new Paragraph(str: "Dated:________________") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p19);
                Paragraph p20 = new Paragraph(str: "THE REGIONAL OFFICER\nCENTRAL BOARD OF SECONDARY EDUCATION,\nREGIONAL OFFICE\nAMBIKA COMPLEX, BEHIND SBI COLONY NEAR BRAHMSTHAN\nSHEIKHPURA, BAILEY ROAD PATNA, BIHAR - 800014\n\nSir,\n\nWith  reference  to  your  Confidential  Letter  No.:CBSE/RO(PTN)/CONF./NODAL / " + dr["cns_schno"].ToString() + " / 2022 TERM 2 / Dtd  " + date_str + ". I  hereby  accept  to   act as Cheif Nodal Supervisor for Sr. as well as Sec.School AISSE/AISSCE 2022 TERM 2.\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p20);
                Paragraph p21 = new Paragraph(str: "      My appointment and any information which may come to my notice during the discharge of my duties as Cheif Nodal supervisor will be kept confidential.  I undertake to do this work with perfect efficiency and according to  the instructions issued by the Board from time to time.\n\n      I  CERTIFY THAT I HAVE NO NEAR  RELATION / WARD  INTENDING / APPEARING  IN THE SUBJECTS ALLOTTED AT MY NODAL CENTRE AT  THE  AFORESAID  EXAMINATION.  I  ALSO CERTIFY THAT I HAVE NOT  WRITTEN  ANY HELP  BOOK OR NOTES FOR THE  EXAMINATION OF THE BOARD.  I UNDERTAKE TO  COMPLETE  THE  WORK ENTRUSTED  TO  ME WITHIN THE STIPULATED TIME GIVEN IN THE LETTER OF APPOINTMENT.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p21);
                Paragraph p23 = new Paragraph(str: "Yours faithfully,     \n\nSignature ..................................\n\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p23);
                PdfPTable tbl3 = new PdfPTable(2);
                tbl3.WidthPercentage = 100f;
                tbl3.HorizontalAlignment = 1;
                tbl3.DefaultCell.Border = 0;
                tbl3.AddCell(new Phrase("Telephone (Office): " + dr["CNSSTD"].ToString() + "  " + dr["CNSteleres"].ToString() + "\n\n(Resi.if any):\n(Mobile No.): " + dr["CNSmobile"].ToString() + "\n"));
                tbl3.AddCell(new Phrase("Name: " + dr["CNSname"].ToString() + "\n\nSchool: " + dr["cnsadd1"].ToString() + "\n             " + dr["cnsadd2"].ToString() + "\n             " + dr["cnsadd3"].ToString() + "\n             " + dr["cnsadd4"].ToString() + "\n             (" + dr["cnsadd5"].ToString() + " - " + dr["cnspin"].ToString() + ")"));
                pdoc.Add(tbl3);
                //
                pdoc.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
