using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LearningSystem
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            skinEngine1.SkinFile = p.AppFolder + @"\MacOS.ssk";
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            loadUI();
        }

        private void loadUI()
        {
            // skinEngine1.SkinFile = p.AppFolder + @"\MacOS.ssk";
            this.Text = "Compare FTP files & DB Files,Ver:" + Application.ProductVersion + "(Edward_song@yeah.net)";
        }

    }
}
