using HSCTSV.View;
using System.Windows.Forms;

namespace HSCTSV
{
    public partial class Home : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Home()
        {
            InitializeComponent();
            FirstView();

        }
        public void FirstView()
        {
            plHome.Controls.Clear();
            ListPaperWaitStudent view = new ListPaperWaitStudent();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
        }


        private void rcpiInfor_ItemClick(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            InfoStudent view = new InfoStudent();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }


        private void rcpiReq_ItemClick_1(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            ListPaperWaitStudent view = new ListPaperWaitStudent();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }


        private void rpiwaitaccept_ItemClick(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            ListWait view = new ListWait();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }


        private void rcpiInfor_ItemClick_1(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            InfoStudent view = new InfoStudent();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }

        private void rpiListInfo_ItemClick(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            ListStudent view = new ListStudent();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }

        private void rpiPrintPaper_ItemClick(object sender, DevExpress.XtraBars.Ribbon.RecentItemEventArgs e)
        {
            plHome.Controls.Clear();
            PaperStudentF view = new PaperStudentF();
            view.TopLevel = false;
            view.Dock = DockStyle.Fill;
            view.Visible = true;
            view.FormBorderStyle = FormBorderStyle.None;
            plHome.Controls.Add(view);
            bvCtrHome.Close();
        }


    }
}