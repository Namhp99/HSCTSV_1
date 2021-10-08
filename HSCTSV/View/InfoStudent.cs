using HSCTSV.ConnectDB;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HSCTSV.View
{
    public partial class InfoStudent : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public InfoStudent()
        {
            InitializeComponent();
        }
        public async void GetInfoStudent()
        {
            var objReq = new Dictionary<string, string>()
            {
                {"TokenCode",Login.userLogin.TokenCode },
                {"IDStudent",txtFind.Text }
            };
            var objRes = await new StudentDB().GetInfoByIdStudent(objReq);
            if (objRes.RespCode != 0)
            {
                MessageBox.Show(objRes.RespText);
                return;
            }
            rtbinfo1.Text = "Họ và tên: " + objRes.Infor.NameStudent + "\n\n" + "MSSV: " + objRes.Infor.IDStudent + "\n\n" + "Ngày sinh: " + objRes.Infor.BirthDay + "\n\n" + "Lớp: " + objRes.Infor.Class;
            rtbinfo2.Text = "\nKhoa/Viện: " + objRes.Infor.Faculty + "\n\n" + "Số CMTND: " + objRes.Infor.CMT + "\n\n" + "Địa chỉ: " + objRes.Infor.Address + "\n\n" + "Họ tên bố: " + objRes.Infor.Father + "\n\n" + "Họ tên mẹ: " + objRes.Infor.Mother + "\n\n" + "Dân tộc: " + objRes.Infor.Nation;
            rtbinfo3.Text = "MSSV: " + objRes.Infor.IDStudent + "\n\n" + "Kì học: " + objRes.Infor.Semester + "\n\n" + "Tình trạng học tập: " + objRes.Infor.StatusInfor + "\n\n" + "Trung bình tích lũy: " + objRes.Infor.PointTB + "\n\n" + "Số TC tích lũy: " + objRes.Infor.TC_All + "\n\n" + "Trình độ sinh viên: " + objRes.Infor.LevelStudent + "\n\n" + "Mức cảnh báo: " + objRes.Infor.LevelWarning;
        }

        private void btnref_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            GetInfoStudent();
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            GetInfoStudent();
            picImage.LoadAsync("https://ctt-sis.hust.edu.vn/Content/Anh/" + "anh_" + txtFind.Text + ".jpg");
        }

        private void txtFind_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
            {
                btSearch.PerformClick();
            }
        }

        private void picImage_Click(object sender, EventArgs e)
        {
            ViewImage view = new ViewImage();
            view.setupImage(picImage.Image);
            view.Show();
        }

        private void InfoStudent_Load(object sender, EventArgs e)
        {

        }
    }
}