using HSCTSV.ConnectDB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace HSCTSV.View
{
    public partial class ListWait : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public ListWait()
        {
            InitializeComponent();
            GetPaperAllNoAccept();
        }
        public async void GetPaperAllNoAccept()
        {
            var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                };
            var objRes = await new PaperDB().GetPaperAllNoAccept(objReq);
            
            grCtrListWait.DataSource = objRes.HSPaperStudentInforLst.ToList();
            gvListWait.BestFitColumns();
            
        }

        

        private void gvListWait_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
                //gvInforPaperStudent.IndicatorWidth = e.Info.DisplayText.Length;
            }
        }
       
        private void btrefresh_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            GetPaperAllNoAccept();
            rtbchitiet.Text = null;
            rtbinfoo2.Text = null;
        }

       

        

        

        private async void gvListWait_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            try
            {
                rtbchitiet.Text = null;
                var objReq = new Dictionary<string, string>()
                {
                    {"TokenCode",Login.userLogin.TokenCode },
                    {"ID",gvListWait.GetRowCellValue(gvListWait.FocusedRowHandle, "RowID").ToString() }
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
                    // label5.Visible = true;
                    rtbchitiet.AppendText(i++ + ". " + a.Question + "  " + a.Answer + "\n");
                    rtbchitiet.Find(j++ + ". " + a.Question);
                    rtbchitiet.SelectionColor = Color.Blue;
                }
                var objReq5 = new Dictionary<string, string>()
                {
                   {"UNumberId",gvListWait.GetRowCellValue(gvListWait.FocusedRowHandle, "IDStudent").ToString() },
                   {"TokenCode",Login.userLogin.TokenCode }
                };

                var objRes5 = await new StudentDB().GetStudentById(objReq5);
                if (objRes5.RespCode != 0)
                {
                    MessageBox.Show(objRes5.RespText);
                    return;
                }
                string b = gvListWait.GetRowCellValue(gvListWait.FocusedRowHandle, "TimeCreate").ToString();
                b = DateTime.ParseExact(b, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture).ToString("HH:mm:ss dd/MM/yyyy");
                rtbinfoo2.Text = "\nHọ và tên: " + objRes5.HSStudentInfo.UFullName
                               + "\t" + "MSSV: " + gvListWait.GetRowCellValue(gvListWait.FocusedRowHandle, "IDStudent").ToString() + "\n\n"
                               + "Yêu cầu: " + gvListWait.GetRowCellValue(gvListWait.FocusedRowHandle, "Description").ToString() + "\n" +
                                 "Ngày tạo: " + b;
            }
            catch (Exception ex)
            {
                rtbinfoo2.Text = null;
                MessageBox.Show(ex.Message);
            }
        }
    }
}