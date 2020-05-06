using System;
using System.IO;
using PowerTerm;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;

namespace medNecUpdated
{
    class Program
    {
        private static Conversation pt;
        private static string channel = "";
        private static string user = "";
        private static string pass = "";

        private static void signon(string region)
        {
            pt = new Conversation();
            pt.Open(channel);

            pt.ReConnect("hma-dar-mf.hma.com");
            System.Threading.Thread.Sleep(2000);
            pt.WaitSystem();

            string chcstr;

            if(user == "" || pass == "")
            {
                Console.WriteLine("Please enter username.");
                user = Console.ReadLine().Trim();
                Console.WriteLine("Please enter password.");
                pass = Console.ReadLine().Trim();
            }

            chcstr = pt.Screen(2 , 6 , 13);

            if (chcstr == "COMPUTER")
            {

                pt.Cursor(20 , 18);
                pt.Send(region);
                pt.Enter();
                pt.WaitSystem();
                System.Threading.Thread.Sleep(2559);

                chcstr = pt.Screen(18 , 2 , 5);

                if (chcstr == "USER")
                {
                    pt.Cursor(18 , 11);                    
                    pt.Send(user);
                    pt.Cursor(19 , 11);                  
                    pt.Send(pass);
                    pt.Enter();
                    pt.WaitSystem();

                    chcstr = pt.Screen(3 , 4 , 8);

                    if (chcstr == "Enter")
                    {
                        pt.Cursor(24 , 6);
                        pt.Send("1");
                        pt.Enter();
                        pt.WaitSystem();
                    }
                    else
                    {
                        Console.WriteLine("PASSWORD RESET NEEDED OR RE-ENTRY NEEDED");
                        user = "";
                        pass = "";
                    }
                }
            }
        }

        private static void takeNoteMed(string noteMed)
        {

            pt = new Conversation();
            pt.Open(channel);

            string chcstr;
            string noteMed1;
            string noteMed2;

            chcstr = pt.Screen(2 , 33 , 40);

            if (chcstr == "COMMENTS")
            {
                pt.Cursor(7 , 6);
                pt.Send("ASSIST");

                if (Strings.Len(noteMed) > 199)
                {
                    noteMed1 = Strings.Left(noteMed , 199);
                    noteMed2 = Strings.Mid(noteMed , 200 , 199);
                }
                else
                {
                    noteMed1 = noteMed;
                    noteMed2 = null;
                }

                if (noteMed1 != null)
                {
                    pt.Cursor(7 , 16);
                    pt.Send(noteMed1);
                    pt.Enter();
                    pt.WaitSystem();

                }

                if (noteMed2 != null)
                {
                    pt.Cursor(7 , 6);
                    pt.Send("ASSIST");
                    pt.Cursor(7 , 16);
                    pt.Send(noteMed2);
                    pt.Enter();
                    pt.WaitSystem();
                }

                pt.F7();
                pt.WaitSystem();
            }
        }

        private static void writeoff953(string adjAmt953,string comment)
        {
            pt = new Conversation();
            pt.Open(channel);

            string chcstr;
            string adjCode;
            
            adjCode = "953";

            comment = comment + adjCode;

            chcstr = pt.Screen(2 , 2 , 8);

            if (chcstr == "ACCOUNT")
            {
                pt.F13();
                pt.WaitSystem();

                do
                {
                    System.Threading.Thread.Sleep(0050);
                    chcstr = pt.Screen(15 , 53 , 62);

                } while (chcstr != "ADJUSTMENT");

                if (chcstr == "ADJUSTMENT")
                {
                    pt.Cursor(18 , 54);
                    pt.FieldExit();
                    pt.Cursor(18 , 54);
                    pt.Send(adjCode);
                    pt.Cursor(19 , 54);
                    pt.FieldExit();
                    pt.Cursor(19 , 54);
                    pt.Send(adjAmt953);
                    pt.Cursor(20 , 54);
                    pt.FieldExit();
                    pt.Cursor(20 , 54);
                    pt.Send(comment);
                    pt.F1();
                    pt.WaitSystem();
                    pt.F3();
                    pt.WaitSystem();
                    System.Threading.Thread.Sleep(0500);
                }

                pt.F7();
                pt.WaitSystem();
            }
        }

        private static void writeoff986(string adjAmt986, string comment)
        {
            pt = new Conversation();
            pt.Open(channel);

            string chcstr;
            string adjCode;
            
            adjCode = "986";

            comment = comment + adjCode;

            chcstr = pt.Screen(2 , 2 , 8);

            if (chcstr == "ACCOUNT")
            {
                pt.F13();
                pt.WaitSystem();

                do
                {
                    System.Threading.Thread.Sleep(0050);
                    chcstr = pt.Screen(15 , 53 , 62);

                } while (chcstr != "ADJUSTMENT");


                if (chcstr == "ADJUSTMENT")
                {
                    pt.Cursor(18 , 54);
                    pt.FieldExit();
                    pt.Cursor(18 , 54);
                    pt.Send(adjCode);
                    pt.Cursor(19 , 54);
                    pt.FieldExit();
                    pt.Cursor(19 , 54);
                    pt.Send(adjAmt986);
                    pt.Cursor(20 , 54);
                    pt.FieldExit();
                    pt.Cursor(20 , 54);
                    pt.Send(comment);
                    pt.F1();
                    pt.WaitSystem();
                    pt.F3();
                    pt.WaitSystem();
                    System.Threading.Thread.Sleep(0500);
 
                }

                pt.F7();
                pt.WaitSystem();

            }
        }

        static Boolean deciderMed()
        {
            pt = new Conversation();
            pt.Open(channel);

         /*   Boolean noteFound;
            string chcstr;
            string darNote;

            noteFound = false;

            do
            {
                chcstr = pt.Screen(23 , 2 , 7);

                for (int i = 6; i < 23; i++)
                {
                    darNote = pt.Screen(i , 12 , 80);

                    if (Strings.InStr(darNote,"986") > 0)
                    {
                        if (Strings.InStr(darNote,adjAmt986)>0)
                        {
                            noteFound = true;
                        }
                    }

                    if (Strings.InStr(darNote , "953") > 0)
                    {
                        if (Strings.InStr(darNote,adjAmt953) > 0)
                        {
                            noteFound = true;
                        }
                    }
                }

                pt.Enter();
                pt.WaitSystem();

            } while (chcstr != "BOTTOM");
            */
            pt.F7();
            pt.WaitSystem();

            return false;
        }

        static void Main(string[] args)
        {
            pt = new Conversation();
            channel = pt.getNextClient("DAR");
            //channel = "a";
            pt.Open(channel);
            Boolean decider;
            string dpath;
            string fpath;
            DateTime localDate = DateTime.Now;
            string lcdt = localDate.ToString("MMddyy");
            //string lcdt = "040516";

            // dpath = @"S:\ARSC Scripting Backup\Med Nec\";
            //fpath = dpath + "Med Nec Report " + lcdt + ".xlsx";
            //fpath = dpath + "Med Nec Report" + " 03CO16"  + ".xlsx";
            Console.WriteLine("Please drag and drop the report you wish to process.");
            fpath = Console.ReadLine( ).Trim( );
            //string fname = Path.GetFileName(fpath.Replace("\"", ""));

            Excel.Worksheet workingMed;
            Excel.Worksheet bypassMed;

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.ScreenUpdating = false;
            excelApp.DisplayAlerts = false;
            Excel.Workbook newWorkbook = excelApp.Workbooks.Open(fpath);

            workingMed = ((Excel.Worksheet)newWorkbook.Sheets["working"]);
            bypassMed = ((Excel.Worksheet)newWorkbook.Sheets["bypass"]);

            int tRows;
            int tRows2;
            int force1 = 0, force2 = 0;
            int x = 0, y = 0;
            string chcstr;
            string adjAmt953;
            string adjAmt986;
            string coidMed;
            string acctMed;
            string noteMed;
            string comment;
            string adjAmt9532;
            string adjAmt9862;
            string needsForce = "";

            tRows = workingMed.UsedRange.Rows.Count;
            tRows2 = bypassMed.UsedRange.Rows.Count;

            string[] regions = new string[] { "A" , "C" , "I" };

            foreach (string region in regions)
            {
                signon(region);

                chcstr = pt.Screen(7 , 16 , 22);

                if (chcstr == "PATIENT")
                {
                    Console.WriteLine("Working");
                    for (int i = 2; i <= tRows; i++)
                    //for (int i = 2; i < 11; i++)
                    {
                        decider = false;
                        coidMed = Convert.ToString(workingMed.get_Range("A" + i , "A" + i).Value2);
                        acctMed = Convert.ToString(workingMed.get_Range("B" + i , "B" + i).Value2);
                        adjAmt953 = Convert.ToString(workingMed.get_Range("D" + i , "D" + i).Value2);
                        adjAmt986 = Convert.ToString(workingMed.get_Range("C" + i , "C" + i).Value2);
                        noteMed = Convert.ToString(workingMed.get_Range("E" + i, "E" + i).Value2);
                        needsForce = Convert.ToString(workingMed.get_Range("Y" + i, "Y" + i).Value2);
                        comment = "DENIAL ";
                        if(needsForce != "FORCE")
                            coidMed = null;
                        else
                        {
                            force1++;
                            x = Console.CursorLeft;
                            y = Console.CursorTop;
                            Console.WriteLine(force1 + " Forced.");
                            Console.CursorLeft = x;
                            Console.CursorTop = y;
                        }
                        if (adjAmt953 != null)
                        {
                            adjAmt953 = adjAmt953.Trim();
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt986 = adjAmt986.Trim();
                        }

                        if (adjAmt953 != null)
                        {
                            adjAmt953 = adjAmt953.Replace("$" , " ");
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt986 = adjAmt986.Replace("$" , " ");
                        }

                        if (adjAmt953 != null)
                        {
                            adjAmt9532 = Convert.ToDouble(adjAmt953).ToString("F2");
                            adjAmt953 = Convert.ToString(adjAmt9532);
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt9862 = Convert.ToDouble(adjAmt986).ToString("F2");
                            adjAmt986 = Convert.ToString(adjAmt9862);
                        }

                        if (noteMed != null)
                        {
                            noteMed = noteMed.Trim();
                        }

                        //acctMed = Strings.Right(acctMed,7);

                        if (coidMed != null)
                        {
                            pt.Cursor(1 , 76);
                            pt.Send(coidMed);
                            pt.Cursor(7 , 32);
                            pt.Send(acctMed);
                            pt.F8();
                            pt.WaitSystem();

                            chcstr = pt.Screen(2 , 42 , 48);

                            if (chcstr == "INQUIRY")
                            {
                                decider = deciderMed();


                                if (decider == false)
                                {
                                    if (adjAmt953 != null && adjAmt953 != "N")
                                    {
                                        pt.Cursor(1 , 76);
                                        pt.Send(coidMed);
                                        pt.Cursor(7 , 32);
                                        pt.Send(acctMed);
                                        pt.F1();
                                        pt.WaitSystem();

                                        writeoff953(adjAmt953 , comment);

                                        workingMed.Cells[i , 24] = "953 TAKEN";
                                    }


                                    if (adjAmt986 != null && adjAmt986 != "N")
                                    {
                                        pt.Cursor(1 , 76);
                                        pt.Send(coidMed);
                                        pt.Cursor(7 , 32);
                                        pt.Send(acctMed);
                                        pt.F1();
                                        pt.WaitSystem();

                                        writeoff986(adjAmt986 , comment);

                                        workingMed.Cells[i , 23] = "986 TAKEN";
                                    }


                                    if (noteMed != null)
                                    {
                                        pt.Cursor(1 , 76);
                                        pt.Send(coidMed);
                                        pt.Cursor(7 , 32);
                                        pt.Send(acctMed);
                                        pt.F9();
                                        pt.WaitSystem();

                                        takeNoteMed(noteMed);

                                        workingMed.Cells[i , 22] = "NOTED";
                                    }
                                }
                                else
                                {
                                    workingMed.Cells[i , 21] = "ERROR ADJ FOUND: MANUAL REVIEW";
                                }
                            }
                        }
                    }
                }
                coidMed = null;
                acctMed = null;
                adjAmt953 = null;
                adjAmt986 = null;
                noteMed = null;
                comment = null;
                adjAmt9532 = null;
                adjAmt9862 = null;

                chcstr = pt.Screen(7 , 16 , 22);

                if (chcstr == "PATIENT")
                {
                    Console.WriteLine("BYPASS");
                    for (int i = 2; i <= tRows2; i++)
                    //for (int i = 2; i < 11; i++)
                    {
                        decider = false;
                        coidMed = Convert.ToString(bypassMed.get_Range("A" + i , "A" + i).Value2);
                        acctMed = Convert.ToString(bypassMed.get_Range("B" + i , "B" + i).Value2);
                        adjAmt953 = Convert.ToString(bypassMed.get_Range("D" + i , "D" + i).Value2);
                        adjAmt986 = Convert.ToString(bypassMed.get_Range("C" + i , "C" + i).Value2);
                        noteMed = Convert.ToString(bypassMed.get_Range("E" + i , "E" + i).Value2);
                        needsForce = Convert.ToString(bypassMed.get_Range("Y" + i, "Y" + i).Value2);
                        if(needsForce != "FORCE" || coidMed == "875")
                            coidMed = null;
                        else
                        {
                            force2++;
                            x = Console.CursorLeft;
                            y = Console.CursorTop;
                            Console.WriteLine(force2 + " Forced.");
                            Console.CursorLeft = x;
                            Console.CursorTop = y;
                        }
                        comment = "NON COV CHARGE ";

                        if (adjAmt953 != null)
                        {
                            adjAmt953 = adjAmt953.Trim();
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt986 = adjAmt986.Trim();
                        }

                        if (adjAmt953 != null)
                        {
                            adjAmt953 = adjAmt953.Replace("$" , " ");
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt986 = adjAmt986.Replace("$" , " ");
                        }

                        if (adjAmt953 != null)
                        {
                            adjAmt9532 = Convert.ToDouble(adjAmt953).ToString("F2");
                            adjAmt953 = Convert.ToString(adjAmt9532);
                        }

                        if (adjAmt986 != null)
                        {
                            adjAmt9862 = Convert.ToDouble(adjAmt986).ToString("F2");
                            adjAmt986 = Convert.ToString(adjAmt9862);
                        }

                        if (noteMed != null)
                        {
                            noteMed = noteMed.Trim();
                        }

                        //acctMed = Strings.Right(acctMed , 7);

                        if (coidMed != null)
                        {

                            pt.Cursor(1 , 76);
                            pt.Send(coidMed);
                            pt.Cursor(7 , 32);
                            pt.Send(acctMed);
                            pt.F8();
                            pt.WaitSystem();

                            chcstr = pt.Screen(2 , 42 , 48);

                            if (chcstr == "INQUIRY")
                            {
                                decider = deciderMed();


                                if (decider == false)
                                {
                                    if (adjAmt953 != null && adjAmt953 != "N")
                                    {
                                        pt.Cursor(1 , 76);
                                        pt.Send(coidMed);
                                        pt.Cursor(7 , 32);
                                        pt.Send(acctMed);
                                        pt.F1();
                                        pt.WaitSystem();

                                        writeoff953(adjAmt953 , comment);

                                        bypassMed.Cells[i , 24] = "953 TAKEN";
                                    }


                                    if (noteMed != null)
                                    {
                                        pt.Cursor(1 , 76);
                                        pt.Send(coidMed);
                                        pt.Cursor(7 , 32);
                                        pt.Send(acctMed);
                                        pt.F9();
                                        pt.WaitSystem();

                                        takeNoteMed(noteMed);

                                        bypassMed.Cells[i , 22] = "NOTED";
                                    }
                                }
                                else
                                {
                                    bypassMed.Cells[i , 21] = "ERROR ADJ FOUND: MANUAL REVIEW";
                                }
                            }
                        }
                    }
                }
                coidMed = null;
                acctMed = null;
                adjAmt953 = null;
                adjAmt986 = null;
                noteMed = null;
                comment = null;
                adjAmt9532 = null;
                adjAmt9862 = null;
            }

            Random random = new Random();
            int randomNumber = random.Next(0 , 10000);

            excelApp.Visible = true;
            excelApp.ScreenUpdating = true;
            newWorkbook.SaveAs(@"S:\ASSIST - Reporting Dashboard\General Reports\Medical Necessity\" +
                DateTime.Today.ToString("yyyy-MM-dd") + " Med Nec Forced " + DateTime.Now.ToString("HHmmss") + ".xlsx");

            Assist_MailSend.MailSend endEmail = new Assist_MailSend.MailSend();
            string fname = Path.GetFileName(fpath.Replace("\"", ""));
            //MsOutlook.Attachment oAttach = mail.Attachments.Add(newWorkbook);

            endEmail.SendUploadEmail("MED NEC FORCES HAVE BEEN ADDED TO REPORTING DASHBOARD - " + fname, "Please see ASSIST Reporting Dashboard for this report.  Thank you.", "User");

            newWorkbook.Close();
            excelApp.Quit();
            //pt.ProcessHandle.CloseMainWindow();
        }
    }
}
