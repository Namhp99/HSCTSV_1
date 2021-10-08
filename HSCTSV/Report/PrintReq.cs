using HSCTSV.ConnectDB;
using HSCTSV.Model;
using System;
using System.Collections.Generic;

namespace HSCTSV.Report
{
    public partial class PrintReq : DevExpress.XtraReports.UI.XtraReport
    {
        public PrintReq()
        {
            InitializeComponent();
            GetPaper();
        }
        public async void GetPaper()
        {
            lbDate.Text = string.Format("Ngày :{0}", DateTime.Now.ToString("dd/MM/yyyy"));
            var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"DateStart",HSPaperDetail.saveDateStart },
                {"DateEnd",HSPaperDetail.saveDateEnd }
            };
            var objRes = await new PaperDB().GetPaperByDate(objReq);
            cellName.Text = HSPrint.NameC;
            cellID.Text = HSPrint.IDC;
            cellClass.Text = HSPrint.ClassC;
            cellFaculty.Text = HSPrint.FacultyC;
            cellReq.Text = HSPrint.DecritionC;
            cellDetail.Text = HSPrint.DetailC;
            cellNote.Text = HSPrint.NoteC;
        }
    }
}
