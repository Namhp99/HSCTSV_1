using HSCTSV.ConnectDB;
using HSCTSV.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace HSCTSV
{
    public partial class ListStudent : Form
    {
        public ListStudent()
        {
            InitializeComponent();
        }
        private async void btnImport_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Text file (*.xlsx)|*.xlsx|Tất cả các file(*.*)|*.*";
            openFile.InitialDirectory = @"C:\";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = openFile.FileName;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                List<HSStudentInfo> userList = new List<HSStudentInfo>();
                List<string> listmssv = new List<string>();
                try
                {
                    // mở file excel
                    var fileInfo = new FileInfo(@"" + fileName + "");
                    var package = new ExcelPackage(fileInfo);

                    // lấy ra sheet đầu tiên để thao tác
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.First();

                    // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
                    for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                    {
                        try
                        {
                            // biến j biểu thị cho một column trong file
                            int j = 2;

                            // lấy ra cột họ tên tương ứng giá trị tại vị trí [i, 1]. i lần đầu là 2
                            // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                            string idstudent = workSheet.Cells[i, j++].Value.ToString();

                            // lấy ra cột ngày sinh tương ứng giá trị tại vị trí [i, 2]. i lần đầu là 2
                            // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                            // lấy ra giá trị ngày tháng và ép kiểu thành DateTime                      
                            listmssv.Add(idstudent);
                            //InForStudent user = new InForStudent()
                            //{
                            //    IDStudent = idstudent
                            //};

                            //// add UserInfo vào danh sách userList
                            //userList.Add(user);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    foreach (var item in listmssv)
                    {
                        var objReq = new Dictionary<string, string>()
                        {
                            {"TokenCode",Login.userLogin.TokenCode },
                            {"UNumberId",item }
                        };
                        var objRes = await new StudentDB().GetStudentById(objReq);
                        if (objRes.RespCode != 0)
                        {
                            MessageBox.Show(objRes.RespText);
                            return;
                        }
                        userList.Add(objRes.HSStudentInfo);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                gCtrListUser.DataSource = userList;
            }
        }
        private void gvlistuser_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
        private void exportExel_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog saveFileExel = new SaveFileDialog();
            saveFileExel.InitialDirectory = @"C\";
            saveFileExel.Title = "Save text Files";
            saveFileExel.Filter = "Text file (*.xlsx)|*.xlsx|Tất cả các file(*.*)|*.*";
            saveFileExel.RestoreDirectory = true;
            if (saveFileExel.ShowDialog() == DialogResult.OK)
            {
                gvlistuser.ExportToXlsx(saveFileExel.FileName);
            }
        }

        private void DownTemp_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog saveFileExel = new SaveFileDialog();
            saveFileExel.InitialDirectory = @"C\";
            saveFileExel.Title = "Save text Files";
            saveFileExel.Filter = "Text file (*.xlsx)|*.xlsx|Tất cả các file(*.*)|*.*";
            saveFileExel.RestoreDirectory = true;
            if (saveFileExel.ShowDialog() == DialogResult.OK)
            {
                Thread thread = new Thread(() =>
                {
                    WebClient client1 = new WebClient();
                    //client1.DownloadProgressChanged += new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
                    //client1.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
                    client1.DownloadFileAsync(new Uri("https://firebasestorage.googleapis.com/v0/b/svdbv-90466.appspot.com/o/Template1.xlsx?alt=media&token=0ba33321-4559-4b72-8394-df0663f912f1"), saveFileExel.FileName);
                });
                thread.Start();
                MessageBox.Show("Tải file hoàn tất, xin hãy kiểm tra thư mục: " + saveFileExel.FileName, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
