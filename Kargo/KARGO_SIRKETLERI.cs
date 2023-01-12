using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kargo
{
    public partial class KARGO_SIRKETLERI : Form
    {
        public KARGO_SIRKETLERI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            YURTICI_KARGO YK=new YURTICI_KARGO();
            YK.Show();
            this.Hide();
        }
    }
}
