using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSV_2_PDF
{
    public partial class frmProgress : Form
    {
        public frmProgress()
        {
            InitializeComponent();
        }

        private int progress;

        public int Progress
        {
            get
            {
                return progress;
            }

            set
            {
                progress = value;
                progressBar1.Value = value;
                progressBar1.Refresh();
            }
        }
        public string Message
        {
            set
            {
                label1.Text = value;label1.Refresh();
            }
        }

        private void frmProgress_Load(object sender, EventArgs e)
        {

        }
    }
}
