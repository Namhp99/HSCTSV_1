using System;
using System.Drawing;
using System.Windows.Forms;

namespace HSCTSV
{
    public partial class ViewImage : Form
    {
        public ViewImage()
        {
            InitializeComponent();
        }
        public void setupImage(Image imageStudent)
        {
            ptImage.Image = imageStudent;
        }

        private void ptImage_Click(object sender, EventArgs e)
        {

        }
    }
}
