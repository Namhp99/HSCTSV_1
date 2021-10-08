using DevExpress.Pdf;
using DevExpress.XtraReports.Design;
using HSCTSV.ConnectDB;
using HSCTSV.Model;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace HSCTSV.View
{
    public partial class PaperStudentF : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public PaperStudentF()
        {
            InitializeComponent();
        }
        HSStudentInfo student = new HSStudentInfo();
        private void ribbonStatusBar_Click(object sender, EventArgs e)
        {
        }
        public async void getUStudentByID()
        {
            var objReq = new Dictionary<string, string>()
                {
                    {"UNumberId",txtMSSV.Text },
                    {"TokenCode",Login.userLogin.TokenCode }
                };
            var objRes = await new StudentDB().GetStudentById(objReq);
            student = objRes.HSStudentInfo;
            if (objRes.RespCode == 0)
            {
                rtbinfo.Text = "Họ và tên: " + objRes.HSStudentInfo.UFullName + "\n\n" + "MSSV: " + objRes.HSStudentInfo.UNumberId + "\n\n" + "Lớp: " +
                    objRes.HSStudentInfo.UClass + "\n\n" + "Khoa: " + objRes.HSStudentInfo.UFaculty;
            }
            else
            {
                MessageBox.Show(objRes.RespText);
            }
        }
        public async void GetPaperStudentById()
        {
            var objReq = new Dictionary<string, string>()
                {
                    {"IDStudent",txtMSSV.Text },
                    {"TokenCode",Login.userLogin.TokenCode }
                };
            var objRes = await new PaperDB().GetPaperByIdStudentLst(objReq);
            var objRes1 = await new PaperDB().GetPaperByIdStudentLstAccept(objReq);
            if (objRes.RespCode == 0)
            {
                grCtrPaperWait.DataSource = objRes.HSPaperStudentInforLst.ToList();
            }
            if (objRes1.RespCode == 0)
            {
                grCtrPaperAccept.DataSource = objRes1.HSPaperStudentInforLst.ToList();
            }
            else
            {
                MessageBox.Show(objRes.RespText);
            }
        }


        HSStudentPaperById paper = new HSStudentPaperById();
        public async void AcceptPaperStudent1()
        {
            Int32[] seclectedRowHandles = gvpaperwait.GetSelectedRows();
            for (int i = 0; i < seclectedRowHandles.Length; i++)
            {
                var HSAcceptPaper = gvpaperwait.GetRow(seclectedRowHandles[i]) as HSStudentPaperById;

                var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowID",HSAcceptPaper.RowID.ToString()},
                {"Note",txtNote.Text }
            };
                var objRes = await new PaperDB().AcceptPaperStudent(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
            }
        }
        public async void AcceptPaperStudent()
        {
            Int32[] seclectedRowHandles = gvpaperaccept.GetSelectedRows();
            for (int i = 0; i < seclectedRowHandles.Length; i++)
            {
                var HSAcceptPaper = gvpaperaccept.GetRow(seclectedRowHandles[i]) as HSStudentPaperById;

                var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowID",HSAcceptPaper.RowID.ToString()},
                {"Note",txtNote.Text }
            };
                var objRes = await new PaperDB().AcceptPaperStudent(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
            }
        }
        public async void rePrint()
        {
            if (MessageBox.Show("Xác nhận phê duyệt và In?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles1 = gvpaperaccept.GetSelectedRows();
            for (int k = 0; k < selectedRowsHandles1.Length; k++)
            {
                var hSStudentPaperInfo1 = gvpaperaccept.GetRow(selectedRowsHandles1[k]) as HSStudentPaperById;
                var objReq3 = new Dictionary<string, string>()
                        {
                           {"TokenCode",Login.userLogin.TokenCode },
                            {"ID",hSStudentPaperInfo1.RowID.ToString() }
                        };
                var objRes3 = await new PaperDB().GetPaperDetailByID(objReq3);
                if (objRes3.RespCode != 0)
                {
                    MessageBox.Show(objRes3.RespText);
                    return;
                }
                bool folderExists = File.Exists("C:/Giay_To/" + objRes3.Data.FileName);
                if (!folderExists)
                {
                    MessageBox.Show("Chưa có mẫu giấy, bấm Tải mẫu !");
                    return;
                }
                try
                {
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes3.Data.FileName);
                    application.Visible = true;
                    /////////
                    async void GetStudentPaper(string Type, Microsoft.Office.Interop.Word.Document Document)
                    {
                        string dateBirth;
                        string monthBirth;
                        string yearBirth;
                        var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"IDStudent",txtMSSV.Text }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",txtMSSV.Text},
                {"TokenCode",Login.userLogin.TokenCode }
            };
                        var objRes2 = await new StudentDB().GetStudentById(objReq2);
                        var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowHeader",Type },

            };
                        var objRes = await new PaperDB().GetPaperByRowHeader(objReq);
                        if (objRes.RespCode == 0)
                        {
                            if (student.UBirthDay == null)
                            {
                                dateBirth = "";
                                monthBirth = "";
                                yearBirth = "";
                            }
                            else
                            {
                                char[] spearator = { '/' };
                                String[] strlist = student.UBirthDay.Split(spearator);
                                dateBirth = strlist[0];
                                monthBirth = strlist[1];
                                yearBirth = strlist[2];
                                DateTime now = DateTime.Now;
                                foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                {
                                    if (field.Code.Text.Contains("name"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFullName);
                                    }
                                    else if (field.Code.Text.Contains("address"))
                                    {
                                        field.Select();
                                        string addr = "";
                                        addr = student.UCity.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

                                        application.Selection.TypeText(myTI.ToTitleCase(addr));
                                    }
                                    else if (field.Code.Text.Contains("city"))
                                    {
                                        field.Select();
                                        string city = "";
                                        city = objRes1.Infor.Address.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                                        application.Selection.TypeText(myTI.ToTitleCase(city));
                                    }
                                    else if (field.Code.Text.Contains("birthDay"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UBirthDay);
                                    }

                                    else if (field.Code.Text.Contains("db"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(dateBirth);
                                    }
                                    else if (field.Code.Text.Contains("mb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(monthBirth);
                                    }
                                    else if (field.Code.Text.Contains("yb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(yearBirth);
                                    }
                                    else if (field.Code.Text.Contains("class"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UClass);
                                    }
                                    else if (field.Code.Text.Contains("kStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UGroupLevel);
                                    }
                                    else if (field.Code.Text.Contains("idStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UNumberId);
                                    }
                                    else if (field.Code.Text.Contains("dn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Day.ToString());
                                    }
                                    else if (field.Code.Text.Contains("mn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Month.ToString());
                                    }
                                    else if (field.Code.Text.Contains("yn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Year.ToString());
                                    }
                                    else if (field.Code.Text.Contains("sex"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.Usex);
                                    }
                                    else if (field.Code.Text.Contains("nganhStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("marjor"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("dantoc"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Nation);
                                    }
                                    else if (field.Code.Text.Contains("LevelProgram"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.LevelProgram);
                                    }
                                    else if (field.Code.Text.Contains("program"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.Program);
                                    }
                                    else if (field.Code.Text.Contains("emailStudent"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Email);
                                    }
                                    else if (field.Code.Text.Contains("level"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.LevelStudent);
                                    }
                                    else if (field.Code.Text.Contains("Semester"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Semester);
                                    }
                                    else if (field.Code.Text.Contains("cmt"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.CMT);
                                    }
                                    else if (field.Code.Text.Contains("BHYT"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.BHYT);
                                    }
                                }
                                for (int i = 0; i < objRes.HSStudentQuestionPaperLst.Count; i++)
                                {
                                    foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                    {

                                        if (field.Code.Text.Contains(i.ToString()))
                                        {
                                            field.Select();
                                            application.Selection.TypeText(objRes.HSStudentQuestionPaperLst[i].Answer);
                                        }
                                    }

                                }
                               
                            }
                        }
                    }
                    /////////
                    AcceptPaperStudent();
                    getUStudentByID();
                    GetStudentPaper(objRes3.Data.RowID, document);
                    GetPaperStudentById();
                    
                   
                    txtNote.Text = null;

                }
                catch
                {
                    MessageBox.Show(@"Cấp quyền ghi đọc file ở đường dẫn C:\Giay_To\" + objRes3.Data.FileName);
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes3.Data.FileName);
                    application.Visible = true;
                    async void GetStudentPaper(string Type, Microsoft.Office.Interop.Word.Document Document)
                    {
                        string dateBirth;
                        string monthBirth;
                        string yearBirth;
                        var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"IDStudent",txtMSSV.Text }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",txtMSSV.Text },
                {"TokenCode",Login.userLogin.TokenCode }
            };
                        var objRes2 = await new StudentDB().GetStudentById(objReq2);
                        var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowHeader",Type },

            };
                        var objRes = await new PaperDB().GetPaperByRowHeader(objReq);
                        if (objRes.RespCode == 0)
                        {
                            if (student.UBirthDay == null)
                            {
                                dateBirth = "";
                                monthBirth = "";
                                yearBirth = "";
                            }
                            else
                            {
                                char[] spearator = { '/' };
                                String[] strlist = student.UBirthDay.Split(spearator);
                                dateBirth = strlist[0];
                                monthBirth = strlist[1];
                                yearBirth = strlist[2];
                                DateTime now = DateTime.Now;
                                foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                {
                                    if (field.Code.Text.Contains("name"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFullName);
                                    }
                                    else if (field.Code.Text.Contains("address"))
                                    {
                                        field.Select();
                                        string addr = "";
                                        addr = student.UAddress.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

                                        application.Selection.TypeText(myTI.ToTitleCase(addr));
                                    }
                                    else if (field.Code.Text.Contains("city"))
                                    {
                                        field.Select();
                                        string city = "";
                                        city = objRes1.Infor.Address.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                                        application.Selection.TypeText(myTI.ToTitleCase(city));
                                    }
                                    else if (field.Code.Text.Contains("birthDay"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UBirthDay);
                                    }

                                    else if (field.Code.Text.Contains("db"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(dateBirth);
                                    }
                                    else if (field.Code.Text.Contains("mb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(monthBirth);
                                    }
                                    else if (field.Code.Text.Contains("yb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(yearBirth);
                                    }
                                    else if (field.Code.Text.Contains("class"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UClass);
                                    }
                                    else if (field.Code.Text.Contains("kStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UGroupLevel);
                                    }
                                    else if (field.Code.Text.Contains("idStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UNumberId);
                                    }
                                    else if (field.Code.Text.Contains("dn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Day.ToString());
                                    }
                                    else if (field.Code.Text.Contains("mn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Month.ToString());
                                    }
                                    else if (field.Code.Text.Contains("yn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Year.ToString());
                                    }
                                    else if (field.Code.Text.Contains("sex"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.Usex);
                                    }
                                    else if (field.Code.Text.Contains("nganhStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("marjor"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("dantoc"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Nation);
                                    }
                                    else if (field.Code.Text.Contains("LevelProgram"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.LevelProgram);
                                    }
                                    else if (field.Code.Text.Contains("program"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.Program);
                                    }
                                    else if (field.Code.Text.Contains("emailStudent"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Email);
                                    }
                                    else if (field.Code.Text.Contains("level"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.LevelStudent);
                                    }
                                    else if (field.Code.Text.Contains("Semester"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Semester);
                                    }
                                    else if (field.Code.Text.Contains("cmt"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.CMT);
                                    }
                                    else if (field.Code.Text.Contains("BHYT"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.BHYT);
                                    }
                                }
                                for (int i = 0; i < objRes.HSStudentQuestionPaperLst.Count; i++)
                                {
                                    foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                    {

                                        if (field.Code.Text.Contains(i.ToString()))
                                        {
                                            field.Select();
                                            application.Selection.TypeText(objRes.HSStudentQuestionPaperLst[i].Answer);
                                        }
                                    }

                                }

                            }
                        }
                    }
                    AcceptPaperStudent();
                    getUStudentByID();
                    GetStudentPaper(objRes3.Data.RowID, document);
                    GetPaperStudentById();
                    txtNote.Text = null;
                }
            }

        }

        public async void PrintPaper()
        {
            if (MessageBox.Show("Xác nhận phê duyệt và In?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles1 = gvpaperwait.GetSelectedRows();
            for (int k = 0; k < selectedRowsHandles1.Length; k++)
            {
                var hSStudentPaperInfo1 = gvpaperwait.GetRow(selectedRowsHandles1[k]) as HSStudentPaperById;
                var objReq3 = new Dictionary<string, string>()
                        {
                           {"TokenCode",Login.userLogin.TokenCode },
                            {"ID",hSStudentPaperInfo1.RowID.ToString() }
                        };
                var objRes3 = await new PaperDB().GetPaperDetailByID(objReq3);
                bool folderExists = File.Exists("C:/Giay_To/" + objRes3.Data.FileName);
                if (!folderExists)
                {
                    MessageBox.Show("Chưa có mẫu giấy, bấm Tải mẫu !");
                    return;
                }
                try
                {
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes3.Data.FileName);
                    application.Visible = true;
                    /////////
                    async void GetStudentPaper(string Type, Microsoft.Office.Interop.Word.Document Document)
                    {
                        string dateBirth;
                        string monthBirth;
                        string yearBirth;
                        var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"IDStudent",txtMSSV.Text }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",txtMSSV.Text },
                {"TokenCode",Login.userLogin.TokenCode }
            };
                        var objRes2 = await new StudentDB().GetStudentById(objReq2);
                        var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowHeader",Type },

            };
                        var objRes = await new PaperDB().GetPaperByRowHeader(objReq);
                        if (objRes.RespCode == 0)
                        {
                            if (student.UBirthDay == null)
                            {
                                dateBirth = "";
                                monthBirth = "";
                                yearBirth = "";
                            }
                            else
                            {
                                char[] spearator = { '/' };
                                String[] strlist = student.UBirthDay.Split(spearator);
                                dateBirth = strlist[0];
                                monthBirth = strlist[1];
                                yearBirth = strlist[2];
                                DateTime now = DateTime.Now;
                                foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                {
                                    if (field.Code.Text.Contains("name"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFullName);
                                    }
                                    else if (field.Code.Text.Contains("address"))
                                    {
                                        field.Select();
                                        string addr = "";
                                        addr = student.UCity.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

                                        application.Selection.TypeText(myTI.ToTitleCase(addr));
                                    }
                                    else if (field.Code.Text.Contains("city"))
                                    {
                                        field.Select();
                                        string city = "";
                                        city = objRes1.Infor.Address.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                                        application.Selection.TypeText(myTI.ToTitleCase(city));
                                    }
                                    else if (field.Code.Text.Contains("birthDay"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UBirthDay);
                                    }

                                    else if (field.Code.Text.Contains("db"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(dateBirth);
                                    }
                                    else if (field.Code.Text.Contains("mb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(monthBirth);
                                    }
                                    else if (field.Code.Text.Contains("yb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(yearBirth);
                                    }
                                    else if (field.Code.Text.Contains("class"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UClass);
                                    }
                                    else if (field.Code.Text.Contains("kStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UGroupLevel);
                                    }
                                    else if (field.Code.Text.Contains("idStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UNumberId);
                                    }
                                    else if (field.Code.Text.Contains("dn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Day.ToString());
                                    }
                                    else if (field.Code.Text.Contains("mn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Month.ToString());
                                    }
                                    else if (field.Code.Text.Contains("yn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Year.ToString());
                                    }
                                    else if (field.Code.Text.Contains("sex"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.Usex);
                                    }
                                    else if (field.Code.Text.Contains("nganhStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("marjor"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("dantoc"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Nation);
                                    }
                                    else if (field.Code.Text.Contains("LevelProgram"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.LevelProgram);
                                    }
                                    else if (field.Code.Text.Contains("program"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.Program);
                                    }
                                    else if (field.Code.Text.Contains("emailStudent"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Email);
                                    }
                                    else if (field.Code.Text.Contains("level"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.LevelStudent);
                                    }
                                    else if (field.Code.Text.Contains("Semester"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Semester);
                                    }
                                    else if (field.Code.Text.Contains("cmt"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.CMT);
                                    }
                                    else if (field.Code.Text.Contains("BHYT"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.BHYT);
                                    }
                                }
                                for (int i = 0; i < objRes.HSStudentQuestionPaperLst.Count; i++)
                                {
                                    foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                    {

                                        if (field.Code.Text.Contains(i.ToString()))
                                        {
                                            field.Select();
                                            application.Selection.TypeText(objRes.HSStudentQuestionPaperLst[i].Answer);
                                        }
                                    }

                                }

                            }
                        }
                    }
                    /////////
                    AcceptPaperStudent1();
                    getUStudentByID();
                    GetStudentPaper(objRes3.Data.RowID, document);
                    GetPaperStudentById();
                    txtNote.Text = null;

                }
                catch
                {
                    MessageBox.Show(@"Cấp quyền ghi đọc file ở đường dẫn C:\Giay_To\" + objRes3.Data.FileName);
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes3.Data.FileName);
                    application.Visible = true;
                    async void GetStudentPaper(string Type, Microsoft.Office.Interop.Word.Document Document)
                    {
                        string dateBirth;
                        string monthBirth;
                        string yearBirth;
                        var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"IDStudent",txtMSSV.Text }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",txtMSSV.Text },
                {"TokenCode",Login.userLogin.TokenCode }
            };
                        var objRes2 = await new StudentDB().GetStudentById(objReq2);
                        var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowHeader",Type },

            };
                        var objRes = await new PaperDB().GetPaperByRowHeader(objReq);
                        if (objRes.RespCode == 0)
                        {
                            if (student.UBirthDay == null)
                            {
                                dateBirth = "";
                                monthBirth = "";
                                yearBirth = "";
                            }
                            else
                            {
                                char[] spearator = { '/' };
                                String[] strlist = student.UBirthDay.Split(spearator);
                                dateBirth = strlist[0];
                                monthBirth = strlist[1];
                                yearBirth = strlist[2];
                                DateTime now = DateTime.Now;
                                foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                {
                                    if (field.Code.Text.Contains("name"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFullName);
                                    }
                                    else if (field.Code.Text.Contains("address"))
                                    {
                                        field.Select();
                                        string addr = "";
                                        addr = student.UAddress.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

                                        application.Selection.TypeText(myTI.ToTitleCase(addr));
                                    }
                                    else if (field.Code.Text.Contains("city"))
                                    {
                                        field.Select();
                                        string city = "";
                                        city = objRes1.Infor.Address.ToLower();
                                        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                                        application.Selection.TypeText(myTI.ToTitleCase(city));
                                    }
                                    else if (field.Code.Text.Contains("birthDay"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UBirthDay);
                                    }

                                    else if (field.Code.Text.Contains("db"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(dateBirth);
                                    }
                                    else if (field.Code.Text.Contains("mb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(monthBirth);
                                    }
                                    else if (field.Code.Text.Contains("yb"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(yearBirth);
                                    }
                                    else if (field.Code.Text.Contains("class"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UClass);
                                    }
                                    else if (field.Code.Text.Contains("kStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UGroupLevel);
                                    }
                                    else if (field.Code.Text.Contains("idStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UNumberId);
                                    }
                                    else if (field.Code.Text.Contains("dn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Day.ToString());
                                    }
                                    else if (field.Code.Text.Contains("mn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Month.ToString());
                                    }
                                    else if (field.Code.Text.Contains("yn"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(now.Year.ToString());
                                    }
                                    else if (field.Code.Text.Contains("sex"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.Usex);
                                    }
                                    else if (field.Code.Text.Contains("nganhStudent"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("marjor"))
                                    {
                                        field.Select();
                                        application.Selection.TypeText(student.UFaculty);
                                    }
                                    else if (field.Code.Text.Contains("dantoc"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Nation);
                                    }
                                    else if (field.Code.Text.Contains("LevelProgram"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.LevelProgram);
                                    }
                                    else if (field.Code.Text.Contains("program"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes2.HSStudentInfo.Program);
                                    }
                                    else if (field.Code.Text.Contains("emailStudent"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Email);
                                    }
                                    else if (field.Code.Text.Contains("level"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.LevelStudent);
                                    }
                                    else if (field.Code.Text.Contains("Semester"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.Semester);
                                    }
                                    else if (field.Code.Text.Contains("cmt"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.CMT);
                                    }
                                    else if (field.Code.Text.Contains("BHYT"))
                                    {
                                        field.Select();

                                        application.Selection.TypeText(objRes1.Infor.BHYT);
                                    }
                                }
                                for (int i = 0; i < objRes.HSStudentQuestionPaperLst.Count; i++)
                                {
                                    foreach (Microsoft.Office.Interop.Word.Field field in document.Fields)
                                    {

                                        if (field.Code.Text.Contains(i.ToString()))
                                        {
                                            field.Select();
                                            application.Selection.TypeText(objRes.HSStudentQuestionPaperLst[i].Answer);
                                        }
                                    }

                                }

                            }
                        }
                    }
                    AcceptPaperStudent1();
                    getUStudentByID();
                    GetStudentPaper(objRes3.Data.RowID, document);
                    GetPaperStudentById();
                    txtNote.Text = null;
                }
            }

        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            getUStudentByID();
            GetPaperStudentById();
            // gvpaperaccept.BestFitColumns();

            ptBImage.LoadAsync("https://ctt-sis.hust.edu.vn/Content/Anh/" + "anh_" + txtMSSV.Text + ".jpg");
            tbDetail.Text = null;
        }

        private void txtMSSV_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
            {
                btnSearch.PerformClick();
            }
        }

        private void tbDetail_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void ptBImage_Click(object sender, EventArgs e)
        {
            ViewImage view = new ViewImage();
            view.setupImage(ptBImage.Image);
            view.Show();
        }

        private void gvpaperaccept_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void gvpaperwait_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private async void gvpaperwait_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"ID",gvpaperwait.GetRowCellValue(gvpaperwait.FocusedRowHandle, "RowID").ToString() }
                };
                var objRes2 = await new PaperDB().GetPaperDetailByID(objReq);
                if (objRes2.RespCode != 0)
                {
                    MessageBox.Show(objRes2.RespText);
                    return;
                }
                tbDetail.Text = null;
                int i = 1;
                int j = 1;
                foreach (var a in objRes2.Data.QuestionLst)
                {
                    tbDetail.AppendText(i++ + ". " + a.Question + "  " + a.Answer + "\n");
                    tbDetail.Find(j++ + ". " + a.Question);
                    tbDetail.SelectionColor = Color.Gray;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void gvpaperaccept_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                tbDetail.Text = null;
                var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"ID",gvpaperaccept.GetRowCellValue(gvpaperaccept.FocusedRowHandle, "RowID").ToString() }
                };
                var objRes2 = await new PaperDB().GetPaperDetailByID(objReq);
                if (objRes2.RespCode != 0)
                {
                    MessageBox.Show(objRes2.RespText);
                    return;
                }
                int i = 1;
                int j = 1;
                foreach (var a in objRes2.Data.QuestionLst)
                {
                    tbDetail.AppendText(i++ + ". " + a.Question + "  " + a.Answer + "\n");
                    tbDetail.Find(j++ + ". " + a.Question);
                    tbDetail.SelectionColor = Color.Blue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btexel_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog saveFileExel = new SaveFileDialog();
            saveFileExel.InitialDirectory = @"C\";
            saveFileExel.Title = "Save text Files";
            saveFileExel.Filter = "Text file (*.xlsx)|*.xlsx|Tất cả các file(*.*)|*.*";
            saveFileExel.RestoreDirectory = true;
            if (saveFileExel.ShowDialog() == DialogResult.OK)
            {
                gvpaperaccept.ExportToXlsx(saveFileExel.FileName);
            }
        }

        private async void btDownLoadFile_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode }
            };
            var objRes = await new PaperDB().GetPaperLst(objReq);
            try
            {
                if (objRes.HSPaperLst == null)
                {
                    MessageBox.Show("Không tìm thấy file.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }


                bool folderExists = Directory.Exists(@"C:/Giay_To");
                if (!folderExists)
                {
                    Directory.CreateDirectory(@"C:/Giay_To");
                }
                foreach (var a in objRes.HSPaperLst.ToList())
                {
                    Thread thread = new Thread(() =>
                    {
                        WebClient client1 = new WebClient();
                        //client1.DownloadProgressChanged += new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
                        //client1.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
                        client1.DownloadFileAsync(new Uri("http://202.191.56.101/HSCTSV/File/Download?FileName=" + a.FilePath + "&TokenCode=" + Login.userLogin.TokenCode), "C:/Giay_To/" + a.FilePath);
                    });
                    thread.Start();
                }
                MessageBox.Show("Tải file hoàn tất, xin hãy kiểm tra thư mục: C:/ Giay_To", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ////Properties.Settings.Default.saveDocument = selectPath;
                ////Properties.Settings.Default.Save();
                //MessageBox.Show(Properties.Settings.Default.saveDocument);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btPrint_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (gvpaperwait.SelectedRowsCount > 0)
            {
                PrintPaper();
                GetPaperStudentById();
            }
            if (gvpaperaccept.SelectedRowsCount > 0)
            {
                rePrint();
                getUStudentByID();
            }
        }

        private void btLoad_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            ListStudentF_Load();
        }

        private void ListStudentF_Load()
        {
            GetPaperStudentById();
        }

        private async void btnRepair_Click(object sender, EventArgs e)
        {
            if (gvpaperaccept.SelectedRowsCount <= 0 & gvpaperwait.SelectedRowsCount<=0) 
            {
                
                return;
            }
            if (MessageBox.Show("Xác nhận chỉnh sửa?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles = gvpaperaccept.GetSelectedRows();
            for (int i = 0; i < selectedRowsHandles.Length; i++)
            {
                var hSStudentPaperinfo = gvpaperaccept.GetRow(selectedRowsHandles[i]) as HSStudentPaperById;
                var objReq = new Dictionary<string, string>()
                {
                    { "TokenCode",Login.userLogin.TokenCode},
                    { "RowID",hSStudentPaperinfo.RowID},
                    { "Note",txtNote.Text}
                };
                var objRes = await new PaperDB().PausePaperStudent(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
            }
            Int32[] selectedRowsHandles1 = gvpaperwait.GetSelectedRows();
            for (int i = 0; i < selectedRowsHandles1.Length; i++)
            {
                var hSStudentPaperinfo1 = gvpaperwait.GetRow(selectedRowsHandles1[i]) as HSStudentPaperById;
                var objReq = new Dictionary<string, string>()
                {
                    { "TokenCode",Login.userLogin.TokenCode},
                    { "RowID",hSStudentPaperinfo1.RowID},
                    { "Note",txtNote.Text}
                };
                var objRes = await new PaperDB().PausePaperStudent(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
            }
            GetPaperStudentById();
            txtNote.Text = null;
        }

        private void ListStudentF_Load(object sender, EventArgs e)
        {
            GetPaperStudentById();
        }
        
        
        
    }
}