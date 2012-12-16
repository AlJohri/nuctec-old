using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using MongoDB.Bson;
//using MongoDB.Driver;
using System.Web;



namespace WindowsFormsApplication1
{
    /*
    public class CTEC
    {
        public string Name { get; set; }
        public string Age { get; set; }
        public string ID { get; set; }
    }

    public class Course
    {
        public CTEC ctec { get; set; }
    }
    */

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void web()
        {
            // Log In
            webBrowser1.Navigate("http://www.northwestern.edu/caesar");
            wait();
            webBrowser1.Document.GetElementById("userid").SetAttribute("value", "netid");
            webBrowser1.Document.GetElementById("pwd").SetAttribute("value", "password");
            wait();
            webBrowser1.Navigate("javascript: document.login.action = \"https://ses.ent.northwestern.edu/psp/caesar/?cmd=login&amp;languageCd=ENG\"; document.login.submit();");
            wait();

            // Navigate to CTEC Page
            webBrowser1.Navigate("https://ses.ent.northwestern.edu/psc/caesar/EMPLOYEE/HRMS/c/NWCT.NW_CT_PUBLIC_VIEW.GBL");
            wait();

            // Set First ComboBox on Page to "Undergraduate" and submit form
            MessageBox.Show("set first combobox to 'UGRD'");
            runJS("javascript: document.getElementById('NW_CT_PB_SRCH_ACAD_CAREER').value = 'UGRD'");
            runJS("javascript: submitAction_win0(win0,'NW_CT_PB_SRCH_ACAD_CAREER')");
            //wait();

            // Set Second ComboBox to "EECS" and submit form
            MessageBox.Show("set second combobox to 'EECS'");
            runJS("javascript: document.getElementById('NW_CT_PB_SRCH_SUBJECT').value = 'EECS'");
            runJS("javascript: submitAction_win0(document.win0,'NW_CT_PB_SRCH_SUBJECT')");
            //wait();
            
            // Click Search
            MessageBox.Show("click Search");
            runJS("javascript: submitAction_win0(document.win0,'NW_CT_PB_SRCH_SRCH_BTN')");
            //wait();

            MessageBox.Show("getting number of courses");
            int numberOfCourses = webBrowser1.Document.GetElementById("win0divNW_CT_PV_DRV$0").GetElementsByTagName("tr").Count - 1;
            

            for (int i = 0; i <= numberOfCourses; i++)
            {
                // Select class in list (EECS 101, EECS 110, EECS 111)
                MessageBox.Show("loop index @ course " + i);
                runJS("submitAction_win0(document.win0,'MYLINK$" + i + "');");
                //wait();

                MessageBox.Show("getting number of CTECS for course " + i);
                int numberOfCTECS = webBrowser1.Document.GetElementById("win0divNW_CT_PV4_DRV$0").GetElementsByTagName("tr").Count - 2;

                for (int j = 0;j <= numberOfCTECS; j++)
                {
                    MessageBox.Show("loop index @ ctec " + j);
                    runJS("submitAction_win0(document.win0,'MYLINK1$" + j + "');");
                    //wait();
                    
                    MessageBox.Show("pause");

                    var courseTitle = webBrowser1.Document.GetElementById("NW_CT_DERIVED_0_NW_CT_COURSE").InnerText;
                    var department = webBrowser1.Document.GetElementById("NW_CT_DERIVED_0_NW_CT_DEPARTMENT").InnerText;
                    var academicTerm = webBrowser1.Document.GetElementById("NW_CT_DERIVED_0_NW_CT_TERM_DESCR").InnerText;
                    var enrollmentCount = webBrowser1.Document.GetElementById("NW_CT_DERIVED_0_NW_CT_ENROLLCOUNT").InnerText;
                    var responseCount = webBrowser1.Document.GetElementById("NW_CT_DERIVED_0_NW_CT_RESPONSECNT").InnerText;
                    var instructor = webBrowser1.Document.GetElementById("NW_CT_PV_NAME_NAME").InnerText;

                    var instructionRating = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$0").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    var courseRating = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$1").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    var amountLearned = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$2").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    var challengeIntellectually = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$3").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    var stimulateInterest = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$4").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    //var averageHours = webBrowser1.Document.GetElementById("win0divNW_CT_PV2_DRV_DESCRLONG$5").GetElementsByTagName("td")[0].GetElementsByTagName("div")[0].GetElementsByTagName("div")[0].InnerText;
                    var averageHours = "";

                    var CTEC = webBrowser1.Document.GetElementById("win0divNW_CT_PV3_DRV_DESCRLONG$0").GetElementsByTagName("font")[1].InnerText.Split('\n')[2];
                    //var CTEC = "";

                    MessageBox.Show(

                        "Course Title: " + courseTitle + Environment.NewLine + 
                        "Department: " + department + Environment.NewLine +
                        "Academic Term : " + academicTerm + Environment.NewLine +
                        "Enrollment Count : " + enrollmentCount + Environment.NewLine +
                        "Response Count : " + responseCount + Environment.NewLine +
                        "Instructor : " + instructor + Environment.NewLine + 
                        
                        Environment.NewLine +

                        "Instructor Rating : " + instructionRating + Environment.NewLine +
                        "Course Rating : " + courseRating + Environment.NewLine +
                        "Amount Learned : " + amountLearned + Environment.NewLine +
                        "Challenge Intellectually : " + challengeIntellectually + Environment.NewLine +
                        "Stimulate Interest : " + stimulateInterest + Environment.NewLine +
                        "Average Hours (to implement) : " + averageHours + Environment.NewLine + 
                        
                        Environment.NewLine + 
                        
                        "CTEC:" + Environment.NewLine + CTEC

                        );

                    //System.Web.Script.Serialization.JavaScriptSerializer oSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

                    //wait();
                    runJS("submitAction_win0(document.win0,'NW_CT_PV_NAME_RETURN_PB');");
                    wait();
                }

                // Click Search Again to Reset Class List
                MessageBox.Show("click Search Again");
                runJS("javascript: submitAction_win0(document.win0,'NW_CT_PB_SRCH_SRCH_BTN')");
                wait();
            }
        }

        private void wait()
        {
            while (webBrowser1.ReadyState != WebBrowserReadyState.Complete) //(webBrowser1.IsBusy)
            {
                Application.DoEvents();
            }
        }

        private void wait2()
        {
            for (long i = 0; i < 1000000; i++) { Application.DoEvents(); }
        }

        private void runJS(string JScript)
        {
            object[] args = { JScript };
            //webBrowser1.Document.InvokeScript("eval", args);
            webBrowser1.Document.InvokeScript("execScript", args);
        }


        // public delegate void InvokeDelegate();

        private void button1_Click(object sender, EventArgs e)
        {
            web();
        }


    }
}


/*

Course Title: win0divNW_CT_DERIVED_0_NW_CT_DEPARTMENT
Department: win0divNW_CT_DERIVED_0_NW_CT_DEPARTMENT
Academic Term: win0divNW_CT_DERIVED_0_NW_CT_TERM_DESCR
Enrollment Count: win0divNW_CT_DERIVED_0_NW_CT_ENROLLCOUNT
Response Count: win0divNW_CT_DERIVED_0_NW_CT_RESPONSECNT
Instructor: win0divNW_CT_PV_NAME_NAME

Multiple Choice Questions: win0divNW_CTEC_M_QUESTION$0
Provide an overall rating of the instruction: win0divNW_CT_PV2_DRV_DESCRLONG$0
Provide an overall rating of the course: win0divNW_CT_PV2_DRV_DESCRLONG$1
Estimate how much you learned in the course: win0divNW_CT_PV2_DRV_DESCRLONG$2
Rate the effectiveness of the course in challenging you intellectually: win0divNW_CT_PV2_DRV_DESCRLONG$3
Rate the effectiveness of the instructor(s) in stimulating your interest in the subject: win0divNW_CT_PV2_DRV_DESCRLONG$4
Estimate the average number of hours per week you spent on this course outside of class and lab time: win0divNW_CT_PV2_DRV_DESCRLONG$5

Please summarize your reaction to this course focusing on the aspects that were most important to you: 
 * win0divNW_CTEC_COMMENTS$0
 * NW_CTEC_COMMENTS$scroll$0
 * win0divGPNW_CTEC_COMMENTS$0
 * ACE_NW_CTEC_COMMENTS$0
 * win0divNW_CT_PV3_DRV_DESCRLONG$0

Course and Teacher Evaluations - Demographic Questions: win0divNW_CTEC_DEMOGRPAHI$0
Your School: win0divNW_CT_PV5_DRV_DESCRLONG$0
Your Class: win0divNW_CT_PV5_DRV_DESCRLONG$1
Your reason for taking course (mark all that apply): win0divNW_CT_PV5_DRV_DESCRLONG$2
Interest in subject before taking course: win0divNW_CT_PV5_DRV_DESCRLONG$3


Back: javascript:submitAction_win0(document.win0,'NW_CT_PV_NAME_RETURN_PB');

Sleep: System.Threading.Thread.Sleep(2000);
*/
