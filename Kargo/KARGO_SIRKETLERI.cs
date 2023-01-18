using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
namespace Kargo
{
    public partial class KARGO_SIRKETLERI : Form
    {
        private Rectangle originalFormRect;
        private Rectangle originalButton1Rect;
        private Rectangle originalButton2Rect;
        private Rectangle originalButton3Rect;
        private Rectangle originalButton4Rect;
        private Rectangle originalLabel1Rect;

        private float originalButton1FontSize;
        private float originalButton2FontSize;
        private float originalButton3FontSize;
        private float originalButton4FontSize;
        private float originalLabel1FontSize;

        private float fontScale = 1;

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

        private void button2_Click(object sender, EventArgs e)
        {
            ARAS_KARGO ARAS=new ARAS_KARGO();
            ARAS.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SURAT_KARGO SURAT = new SURAT_KARGO();
            SURAT.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            MNG_KARGO MNG = new MNG_KARGO();
            MNG.Show();
            this.Hide();
        }

        private void KARGO_SIRKETLERI_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                e.Cancel = true;//Çıkışı durdur

        }

        private void KARGO_SIRKETLERI_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            label1.Text = "ERMED TIP MEDİKAL KARGO FİYAT HESABI...";


            //originalFormRect=new Rectangle(this.Location,this.Size);
            //originalLabel1Rect=new Rectangle(label1.Location,label1.Size);
            //originalButton1Rect=new Rectangle(button1.Location,button1.Size);
            //originalButton2Rect = new Rectangle(button2.Location, button2.Size);
            //originalButton3Rect = new Rectangle(button3.Location, button3.Size);
            //originalButton4Rect = new Rectangle(button4.Location, button4.Size);

            //originalLabel1FontSize=label1.Font.Size;
            //originalButton1FontSize=button1.Font.Size;
            //originalButton2FontSize=button2.Font.Size;
            //originalButton3FontSize=button3.Font.Size;
            //originalButton4FontSize=button4.Font.Size;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = label1.Text.Substring(1) + label1.Text.Substring(0, 1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ANKARA_KARGO ANKR = new ANKARA_KARGO();
            ANKR.Show();
            this.Hide();

        }

        //private void KARGO_SIRKETLERI_Resize(object sender, EventArgs e)
        //{
        //    ResizeChildrenControls();
        //}
        //private void ResizeChildrenControls()
        //{
        //    ResizeControl(button1, originalButton1Rect, originalButton1FontSize);
        //    ResizeControl(button2, originalButton2Rect, originalButton2FontSize);
        //    ResizeControl(button3, originalButton3Rect, originalButton3FontSize);
        //    ResizeControl(button4, originalButton4Rect, originalButton4FontSize);
        //    ResizeControl(label1, originalLabel1Rect, originalLabel1FontSize);
        //}
        //private void ResizeControl(Control control,Rectangle originalControlRect,float originalFontSize)
        //{
        //    float xRatio=(float)ClientRectangle.Width / originalFormRect.Width;
        //    float yRatio=(float)ClientRectangle.Height/ originalFormRect.Height;

        //    float newX=originalControlRect.Location.X* xRatio;
        //    float newY = originalControlRect.Location.Y * xRatio;

        //    control.Location=new Point((int)newX, (int)newY);
        //    control.Width=(int)(originalControlRect.Width*xRatio);
        //    control.Height = (int)(originalControlRect.Height * xRatio);

        //float ratio = xRatio;
        //if(xRatio>=yRatio) 
        //{
        //    ratio = yRatio;
        //}
        //float newFontSize = originalFontSize * ratio * fontScale;
        //Font newFont=new Font(control.Font.FontFamily,newFontSize);
        //control.Font = newFont;
        //}
    }
}
