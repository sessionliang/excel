using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;

namespace Excel
{
    /// <summary>
    /// 自定义控件:半透明控件
    /// </summary>

    [ToolboxBitmap(typeof(LoadingCtl))]
    public class LoadingCtl : System.Windows.Forms.Control
    {
        private bool _transparentBG = true;//是否使用透明
        private int _alpha = 125;//设置透明度
        private SolidBrush fontBrush = new SolidBrush(Color.FromArgb(206, 94, 94, 94));

        private System.ComponentModel.Container components = new System.ComponentModel.Container();

        public LoadingCtl()
            : this(125, true)
        {
        }

        /// <summary>
        /// 显示遮罩层
        /// </summary>
        /// <param name="form">要显示的父窗体</param>
        public void ShowLoading(Form form)
        {

            try
            {
                // SetStyle(System.Windows.Forms.ControlStyles.Opaque, true);
                base.CreateControl();
                form.Controls.Add(this);
                // this._alpha = Alpha;
                // if (isShowLoadingImage)
                //{
                // PictureBox pictureBox_Loading = new PictureBox();
                // pictureBox_Loading.BackColor = System.Drawing.Color.White;
                // pictureBox_Loading.Image = 加载中.Properties.Resources.loading;
                // pictureBox_Loading.Name = "pictureBox_Loading";
                // pictureBox_Loading.Size = new System.Drawing.Size(48, 48);
                // pictureBox_Loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
                // Point Location = new Point(this.Location.X + (this.Width - pictureBox_Loading.Width) / 2, this.Location.Y + (this.Height - pictureBox_Loading.Height) / 2);//居中
                // pictureBox_Loading.Location = Location;
                // pictureBox_Loading.Anchor = AnchorStyles.None;
                // this.Controls.Add(pictureBox_Loading);
                //}

                this.Dock = DockStyle.Fill;
                this.BringToFront();
                this.Enabled = true;
                this.Visible = true;
                Refresh();
                Invalidate();

            }
            catch { }

        }


        public LoadingCtl(int Alpha, bool IsShowLoadingImage)
        {
            SetStyle(System.Windows.Forms.ControlStyles.Opaque, true);
            base.CreateControl();

            this._alpha = Alpha;
            if (IsShowLoadingImage)
            {
                PictureBox pictureBox_Loading = new PictureBox();
                pictureBox_Loading.BackColor = System.Drawing.Color.White;
                pictureBox_Loading.Image = Image.FromFile("wait.gif");
                pictureBox_Loading.Name = "pictureBox_Loading";
                pictureBox_Loading.Size = new System.Drawing.Size(48, 48);
                pictureBox_Loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
                Point Location = new Point(this.Location.X + (this.Width - pictureBox_Loading.Width) / 2, this.Location.Y + (this.Height - pictureBox_Loading.Height) / 2);//居中
                pictureBox_Loading.Location = Location;
                pictureBox_Loading.Anchor = AnchorStyles.None;
                this.Controls.Add(pictureBox_Loading);
                Invalidate();
            }
        }


        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!((components == null)))
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// 自定义绘制窗体
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            float vlblControlWidth;
            float vlblControlHeight;

            Pen labelBorderPen;
            SolidBrush labelBackColorBrush;

            if (_transparentBG)
            {
                Color drawColor = Color.FromArgb(this._alpha, this.BackColor);
                labelBorderPen = new Pen(drawColor, 0);
                labelBackColorBrush = new SolidBrush(drawColor);
            }
            else
            {
                labelBorderPen = new Pen(this.BackColor, 0);
                labelBackColorBrush = new SolidBrush(this.BackColor);
            }
            base.OnPaint(e);
            vlblControlWidth = this.Size.Width;
            vlblControlHeight = this.Size.Height;
            e.Graphics.DrawRectangle(labelBorderPen, 0, 0, vlblControlWidth, vlblControlHeight);
            e.Graphics.FillRectangle(labelBackColorBrush, 0, 0, vlblControlWidth, vlblControlHeight);
            e.Graphics.DrawString("正在加载...", new Font("黑体", 10), fontBrush, vlblControlWidth / 2 - 30, vlblControlHeight / 2 + 40);

        }


        protected override CreateParams CreateParams//v1.10 
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00000020; //0x20; // 开启 WS_EX_TRANSPARENT,使控件支持透明
                return cp;
            }
        }

        /*
        * [Category("LoadingCtl"), Description("是否使用透明,默认为True")]
        * 一般用于说明你自定义控件的属性（Property）。
        * Category用于说明该属性属于哪个分类，Description自然就是该属性的含义解释。
        */
        [Category("LoadingCtl"), Description("是否使用透明,默认为True")]
        public bool TransparentBG
        {
            get
            {
                return _transparentBG;
            }
            set
            {
                _transparentBG = value;
                this.Invalidate();
            }
        }

        [Category("LoadingCtl"), Description("设置透明度")]
        public int Alpha
        {
            get
            {
                return _alpha;
            }
            set
            {
                _alpha = value;
                this.Invalidate();
            }
        }
    }


}

/*

外部窗体调用：

public Form1()
{
InitializeComponent();
bgWorker.DoWork += BgWorker_DoWork;
bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
}

BackgroundWorker bgWorker = new BackgroundWorker();
wgscd.LoadingCtl loadingCtl = new wgscd.LoadingCtl(162,true);//定义加载控件
private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
{

System.Threading.Thread.Sleep(3333);

}

private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
{
loadingCtl.Hide();

}

 


private void btnShow_Click(object sender, EventArgs e)
{
//显示加载动画
loadingCtl.ShowLoading(this);
Application.DoEvents();
bgWorker.RunWorkerAsync();

}

 

 


*/