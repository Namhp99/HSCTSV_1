using DevExpress.XtraReports.UI;
using HSCTSV.ConnectDB;
using HSCTSV.Model;
using HSCTSV.Report;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace HSCTSV.View
{
    public partial class ListPaperWaitStudent : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public ListPaperWaitStudent()
        {
            InitializeComponent();
            DateStart.Value = DateTime.Today.AddDays(-14);
            DateEnd.Value = DateTime.Today.AddDays(+1);
            gvInforPaperStudent.BestFitColumns();
        }
        HSStudentPaperById accept = new HSStudentPaperById();
        public async void DonePaperStudent()
        {
            
            Int32[] seclectedRowHandles = gvInforPaperStudent.GetSelectedRows();
            for (int i = 0; i < seclectedRowHandles.Length; i++)
            {
                var HSAcceptPaper = gvInforPaperStudent.GetRow(seclectedRowHandles[i]) as HSStudentPaperById;
                accept = HSAcceptPaper;
                var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowID",HSAcceptPaper.RowID.ToString()},
                {"Note",txtNote.Text }
            };               
                var objRes = await new PaperDB().DonePaperStudent(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
                if (objRes.RespCode == 0)
                {
                    HSAcceptPaper.Note = txtNote.Text;
                    HSAcceptPaper.MailAccept = Login.userLogin.Email;
                    HSAcceptPaper.UserAccept = accept.UserAccept;
                    //gvInforPaperStudent.SetRowCellValue(i, "Note", txtNote.Text);
                }
            }
            gvInforPaperStudent.RefreshData();

        
        }
        public async void GetByDate()
        {
            var objReq6 = new Dictionary<string, string>()
                {
                   {"TokenCode",Login.userLogin.TokenCode },
                   {"DateStart",DateStart.Value.ToString("yyyy-MM-dd HH:mm:ss") },
                   {"DateEnd",DateEnd.Value.ToString("yyyy-MM-dd HH:mm:ss") }
                };
            var objRes6 = await new PaperDB().GetPaperByDate(objReq6);
            grCtrInforPaperStudent.DataSource = objRes6.HSPaperStudentInforLst.ToList();
            gvInforPaperStudent.BestFitColumns();
            grCtrDetail.DataSource = null;
            
        }
        public async void DelPaperStudent()
        {
            Int32[] seclectedRowHandles = gvInforPaperStudent.GetSelectedRows();
            for (int i = 0; i < seclectedRowHandles.Length; i++)
            {
                var HSAcceptPaper = gvInforPaperStudent.GetRow(seclectedRowHandles[i]) as HSStudentPaperById;
                accept = HSAcceptPaper;
                var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"RowID",HSAcceptPaper.RowID.ToString() }
            };
                var objRes = await new PaperDB().DelStudentPaper(objReq);
                if (objRes.RespCode != 0)
                {
                    MessageBox.Show(objRes.RespText);
                    return;
                }
            }
        }

        private async void btnWait_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                };
            var objRes = await new PaperDB().GetPaperAllNoAccept(objReq);
            grCtrInforPaperStudent.DataSource = objRes.HSPaperStudentInforLst.ToList();
            gvInforPaperStudent.BestFitColumns();
        }

        private void btnExportExel_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog saveFileExel = new SaveFileDialog();
            saveFileExel.InitialDirectory = @"C\";
            saveFileExel.Title = "Save text Files";
            saveFileExel.Filter = "Text file (*.xlsx)|*.xlsx|Tất cả các file(*.*)|*.*";
            saveFileExel.RestoreDirectory = true;
            if (saveFileExel.ShowDialog() == DialogResult.OK)
            {
                //gvInforPaperStudent.OptionsView.ColumnAutoWidth=true;
                gvInforPaperStudent.ExportToXlsx(saveFileExel.FileName);

            }
        }

        private void btnreload_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            GetByDate();
            grCtrDetail.DataSource = null;
            
        }

        private void gvInforPaperStudent_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        

        private void gvInforPaperStudent_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName == "Status")
            {
                if (Convert.ToDecimal(e.Value) == 0) e.DisplayText = "Mới tạo";
                if (Convert.ToDecimal(e.Value) == 1) e.DisplayText = "Đã xem";
                if (Convert.ToDecimal(e.Value) == 2) e.DisplayText = "Đã in";
                if (Convert.ToDecimal(e.Value) == 100) e.DisplayText = "Chỉnh sửa";
                if (Convert.ToDecimal(e.Value) == -1) e.DisplayText = "Đã hủy";
            }
        }



        private void btnOK_Click(object sender, EventArgs e)
        {
            if (gvInforPaperStudent.SelectedRowsCount <= 0) return;
            if (MessageBox.Show("Xác nhận thêm ghi chú?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            DonePaperStudent();
            
            //GetByDate();
            
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (gvInforPaperStudent.SelectedRowsCount <= 0) return;
            try
            {
                if (MessageBox.Show("Xác nhận xóa yêu cầu?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    DelPaperStudent();
                    GetByDate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void btnRepair_Click(object sender, EventArgs e)
        {
            if (gvInforPaperStudent.SelectedRowsCount <= 0) return;
            if (MessageBox.Show("Xác nhận chỉnh sửa?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles = gvInforPaperStudent.GetSelectedRows();
            for (int i = 0; i < selectedRowsHandles.Length; i++)
            {
                var hSStudentPaperinfo = gvInforPaperStudent.GetRow(selectedRowsHandles[i]) as HSStudentPaperById;
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
            GetByDate();
        }



        private void ListPaperWaitStudent_Load(object sender, EventArgs e)
        {
            GetByDate();
        }



        private void gvDetail_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName == "Status")
            {
                if (Convert.ToDecimal(e.Value) == 0) e.DisplayText = "Mới tạo";
                if (Convert.ToDecimal(e.Value) == 1) e.DisplayText = "Đã xem";
                if (Convert.ToDecimal(e.Value) == 2) e.DisplayText = "Đã in";
                if (Convert.ToDecimal(e.Value) == 100) e.DisplayText = "Chỉnh sửa";
                if (Convert.ToDecimal(e.Value) == -1) e.DisplayText = "Đã hủy";
            }
        }

        private void gvDetail_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }



        private async void btnPrintP_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (gvInforPaperStudent.SelectedRowsCount <= 0) return;
            HSPrint.DetailC = null;           
            if (MessageBox.Show("Xác nhận phê duyệt và In phiếu hẹn?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles = gvInforPaperStudent.GetSelectedRows();
            for (int i = 0; i < selectedRowsHandles.Length; i++)
            {
                var HSpaper = gvInforPaperStudent.GetRow(selectedRowsHandles[i]) as HSStudentPaperById;
                var objReq = new Dictionary<string, string>()
                {
                  {"UNumberId",HSpaper.IDStudent },
                  {"TokenCode",Login.userLogin.TokenCode }
                };
                var objRes = await new StudentDB().GetStudentById(objReq);
                var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"ID",HSpaper.RowID }
                };
                var objRes1 = await new PaperDB().GetPaperDetailByID(objReq1);


                foreach (var a in objRes1.Data.QuestionLst)
                {

                    HSPrint.DetailC += "-" + a.Question + " " + a.Answer + "\n";

                }
                HSPrint.ClassC = objRes.HSStudentInfo.UClass;
                HSPrint.FacultyC = objRes.HSStudentInfo.UFaculty;
                HSPrint.NameC = HSpaper.NameStudent;
                HSPrint.IDC = HSpaper.IDStudent;

                HSPrint.DecritionC = HSpaper.Description;
                HSPrint.NoteC = HSpaper.Note;


            }
            PrintReq printreq = new PrintReq();
            ReportPrintTool print = new ReportPrintTool(printreq);
            print.ShowPreview();
        }
        // In
        public async void DownLoadFile()
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
                MessageBox.Show("Tải file hoàn tất, xin hãy kiểm tra thư mục: C:/ Giay_To", "Thông báo", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        HSStudentInfo student = new HSStudentInfo();
        public async void getUStudentByID()
        {
            var objReq = new Dictionary<string, string>()
                {
                    {"UNumberId",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle,"IDStudent").ToString() },
                    {"TokenCode",Login.userLogin.TokenCode }
                };
            var objRes = await new StudentDB().GetStudentById(objReq);
            student = objRes.HSStudentInfo;
            
        }
        
            HSStudentPaperById paper = new HSStudentPaperById();
        public async void AcceptPaperStudent()
        {
            Int32[] seclectedRowHandles = gvInforPaperStudent.GetSelectedRows();
            for (int i = 0; i < seclectedRowHandles.Length; i++)
            {
                var HSAcceptPaper = gvInforPaperStudent.GetRow(seclectedRowHandles[i]) as HSStudentPaperById;

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
        public async void PrintPaper()
        {
            if (gvInforPaperStudent.SelectedRowsCount <= 0) return;
            if (MessageBox.Show("Xác nhận phê duyệt và In?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Int32[] selectedRowsHandles = gvInforPaperStudent.GetSelectedRows();
            for (int j = 0; j < selectedRowsHandles.Length; j++)
            {
                var hSStudentPaperInfo = gvInforPaperStudent.GetRow(selectedRowsHandles[j]) as HSStudentPaperById;
                paper = hSStudentPaperInfo;
                var objReq5 = new Dictionary<string, string>()
                        {
                           {"TokenCode",Login.userLogin.TokenCode },
                            {"ID",hSStudentPaperInfo.RowID.ToString()  }
                        };
                var objRes5 = await new PaperDB().GetPaperDetailByID(objReq5);
                bool folderExists = File.Exists("C:/Giay_To/" + objRes5.Data.FileName);
                if (!folderExists)
                {
                    MessageBox.Show("Chưa có mẫu giấy, đang Tải mẫu !");
                    DownLoadFile();
                    return;

                }
                try
                {
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes5.Data.FileName);
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
                    {"IDStudent",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle,"IDStudent").ToString() }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle,"IDStudent").ToString() },
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
                    /////////
                    AcceptPaperStudent();
                    getUStudentByID();
                    GetStudentPaper(objRes5.Data.RowID, document);
                    GetByDate();
                    txtNote.Text = null;

                }
                catch
                {
                    MessageBox.Show(@"Cấp quyền ghi đọc file ở đường dẫn C:\Giay_To\" + objRes5.Data.FileName);
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
                    document = application.Documents.Add(Template: @"C:\Giay_To\" + objRes5.Data.FileName);
                    application.Visible = true;
                    async void GetStudentPaper(string Type, Microsoft.Office.Interop.Word.Document Document)
                    {
                        string dateBirth;
                        string monthBirth;
                        string yearBirth;
                        var objReq1 = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"IDStudent",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle,"IDStudent").ToString() }
                };
                        var objRes1 = await new StudentDB().GetInfoByIdStudent(objReq1);
                        var objReq2 = new Dictionary<string, string>()
            {
                {"UNumberId",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle,"IDStudent").ToString() },
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
                    GetStudentPaper(objRes5.Data.RowID, document);
                    GetByDate();
                }

            }
        }
        private void btPrint_Click(object sender, EventArgs e)
        {
            PrintPaper();
            GetByDate();
        }

        private void btok_Click_1(object sender, EventArgs e)
        {
            GetByDate();
        }

        

        private async void gvInforPaperStudent_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            grCtrDetail.DataSource = null;
            try
            {

                var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"ID",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "RowID").ToString() }
                };
                var objRes2 = await new PaperDB().GetPaperDetailByID(objReq);
                int i = 1;
                int j = 1;
                tbGhiChu.Text = null;
                foreach (var a in objRes2.Data.QuestionLst)
                {
                    tbGhiChu.AppendText(i++ + ". " + a.Question + "  " + a.Answer + "\n");
                    tbGhiChu.Find(j++ + ". " + a.Question);
                    tbGhiChu.SelectionColor = Color.Blue;

                }
                var objReq5 = new Dictionary<string, string>()
                {
                   {"UNumberId",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "IDStudent").ToString() },
                   {"TokenCode",Login.userLogin.TokenCode }
                };
                var objRes5 = await new StudentDB().GetStudentById(objReq5);
                var b = gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "TimeCreate").ToString();              
                b = DateTime.ParseExact(b, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture).ToString("HH:mm:ss dd/MM/yyyy");               
                rtbinfoo.Text = "\nHọ và tên: " + objRes5.HSStudentInfo.UFullName
                               + "\t" + "MSSV: " + gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "IDStudent").ToString() + "\n\n"
                               + "Yêu cầu: " + gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "Description").ToString() + "\n"
                                 +"Ngày tạo: " + b ;
                txtNote.Text = objRes2.Data.Note;          
                var objReq8 = new Dictionary<string, string>()
                {
                {"IDStudent",gvInforPaperStudent.GetRowCellValue(gvInforPaperStudent.FocusedRowHandle, "IDStudent").ToString() },
                {"TokenCode",Login.userLogin.TokenCode }
                };
                var objRes8 = await new PaperDB().GetPaperAllByIdStudentLst(objReq8);
                grCtrDetail.DataSource = objRes8.HSPaperStudentInforLst.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void grCtrInforPaperStudent_Click(object sender, EventArgs e)
        {

        }
    }
}