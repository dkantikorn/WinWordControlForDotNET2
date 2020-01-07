using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using msWord = Microsoft.Office.Interop.Word;
using WinWordControl;

namespace WinWordTestApp
{
    /// <summary>
    /// just testing
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {

        // ==============================================================================================
        // User32.dll
        // ==============================================================================================
        [DllImport("user32.dll")]
        public static extern int FindWindow(string strclassName, string strWindowName);

        [DllImport("user32.dll")]
        static extern int SetParent(int hWndChild, int hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        static extern bool SetWindowPos(
            int hWnd,               // handle to window
            int hWndInsertAfter,    // placement-order handle
            int X,                  // horizontal position
            int Y,                  // vertical position
            int cx,                 // width
            int cy,                 // height
            uint uFlags             // window-positioning options
        );

        [DllImport("user32.dll", EntryPoint = "MoveWindow")]
        static extern bool MoveWindow(
            int hWnd,
            int X,
            int Y,
            int nWidth,
            int nHeight,
            bool bRepaint
        );


        const int SWP_DRAWFRAME = 0x20;
        const int SWP_NOMOVE = 0x2;
        const int SWP_NOSIZE = 0x1;
        const int SWP_NOZORDER = 0x4;

        // ==============================================================================================

        private System.Windows.Forms.Button load;
        private MyWordControl winWordControl1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button Restore;
        private System.Windows.Forms.Button close;
        private Button btnNewForm;

        private string formName;

        private Microsoft.Office.Interop.Word.Application wd;
        private int wordWnd = 0;

        /// <summary>
        /// needed designer variable
        /// </summary>
        private System.ComponentModel.Container components = null;

        public Form1()
        {
            this.formName = getFormRunningNumber();
            this.wd = new Microsoft.Office.Interop.Word.Application();
            InitializeComponent();
        }

        /// <summary>
        /// cleanuup ressources
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            // just to be shure!
            winWordControl1.CloseControl();

            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.load = new System.Windows.Forms.Button();
            this.winWordControl1 = new MyWordControl();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.Restore = new System.Windows.Forms.Button();
            this.close = new System.Windows.Forms.Button();
            this.btnNewForm = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // load
            // 
            this.load.Location = new System.Drawing.Point(592, 24);
            this.load.Name = "load";
            this.load.Size = new System.Drawing.Size(68, 32);
            this.load.TabIndex = 1;
            this.load.Text = "load";
            this.load.Click += new System.EventHandler(this.load_Click);
            // 
            // winWordControl1
            // 
            this.winWordControl1.Dock = System.Windows.Forms.DockStyle.Left;
            this.winWordControl1.Location = new System.Drawing.Point(0, 0);
            this.winWordControl1.Name = "winWordControl1";
            this.winWordControl1.Size = new System.Drawing.Size(560, 389);
            this.winWordControl1.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "WordDateien (*.doc)|*.doc";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(592, 152);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(68, 32);
            this.button1.TabIndex = 3;
            this.button1.Text = "PreActivate";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Restore
            // 
            this.Restore.Location = new System.Drawing.Point(592, 208);
            this.Restore.Name = "Restore";
            this.Restore.Size = new System.Drawing.Size(68, 32);
            this.Restore.TabIndex = 4;
            this.Restore.Text = "Restore Word";
            this.Restore.Click += new System.EventHandler(this.Restore_Click);
            // 
            // close
            // 
            this.close.Location = new System.Drawing.Point(592, 72);
            this.close.Name = "close";
            this.close.Size = new System.Drawing.Size(68, 32);
            this.close.TabIndex = 5;
            this.close.Text = "Close";
            this.close.Click += new System.EventHandler(this.close_Click);
            // 
            // btnNewForm
            // 
            this.btnNewForm.Location = new System.Drawing.Point(592, 267);
            this.btnNewForm.Name = "btnNewForm";
            this.btnNewForm.Size = new System.Drawing.Size(68, 32);
            this.btnNewForm.TabIndex = 6;
            this.btnNewForm.Text = "New Form";
            this.btnNewForm.Click += new System.EventHandler(this.btnNewForm_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(672, 389);
            this.Controls.Add(this.btnNewForm);
            this.Controls.Add(this.close);
            this.Controls.Add(this.Restore);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.winWordControl1);
            this.Controls.Add(this.load);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Activated += new System.EventHandler(this.OnActivate);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// main entry for your Application
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }

        private void OnActivate(object sender, System.EventArgs e)
        {
            //if (wordWnd == 0) wordWnd = FindWindow(null, this.Text);
            //if (wordWnd != 0)
            //{
            //    SetParent(wordWnd, this.Handle.ToInt32());
            //}
        }

        private void load_Click(object sender, System.EventArgs e)
        {

            //int wordWnd = FindWindow(String.Empty, this.formName);
            //SetWindowPos(wordWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width - 20, this.Bounds.Height - 20, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME);

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                //IntPtr wordWnd = FindWindow(null, this.Text);
                ////SetWindowPos(wordWnd, IntPtr.Zero, 0, 0, this.Bounds.Width, this.Bounds.Height, SWP_NOZORDER | SWP_NOMOVE);
                //SetParent(IntPtr.Zero, wordWnd);

                //IntPtr wordWnd = FindWindow(null, this.Text);
                //SetWindowPos(wordWnd, IntPtr.Zero, 0, 0, this.Bounds.Width, this.Bounds.Height, SWP_NOZORDER | SWP_NOMOVE);
                //SetParent(IntPtr.Zero, wordWnd);

                //if (wordWnd == 0) wordWnd = FindWindow("Opusapp", null);
                //if (wordWnd == 0) wordWnd = FindWindow(null,this.Text);
                //if(wordWnd != 0){
                //    SetParent(wordWnd, this.Handle.ToInt32());
                //}
                //if (wordWnd == 0) wordWnd = FindWindow(null, this.Text);
                //SetWindowPos(wordWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width + 20, this.Bounds.Height + 20, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME);
                //MoveWindow(wordWnd, -5, -33, this.Bounds.Width + 10, this.Bounds.Height + 57, true);
                //SetWindowPos(wordWnd, this.Handle.ToInt32(), 0, 0,this.Bounds.Width - 20, this.Bounds.Height - 20,SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME);
                
                winWordControl1.LoadDocument(openFileDialog1.FileName);

                //if (wordWnd == 0) wordWnd = FindWindow(null, this.Text);
                //if (wordWnd != 0)
                //{
                //    SetParent(wordWnd, this.Handle.ToInt32());
                //}

                

            }
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            winWordControl1.PreActivate();
        }

        private void Restore_Click(object sender, System.EventArgs e)
        {
            winWordControl1.RestoreWord();
        }

        private void close_Click(object sender, System.EventArgs e)
        {
            winWordControl1.CloseControl();
        }

        private void btnNewForm_Click(object sender, EventArgs e)
        {
            //int wordWnd = FindWindow(String.Empty, this.formName);
            //SetWindowPos(wordWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width - 20, this.Bounds.Height - 20, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME);

            Form1 frmForm1 = new Form1();
            frmForm1.Show();
        }


        private string getFormRunningNumber()
        {
            Random ran = new Random();
            int runningNum = ran.Next(1000);
            return "Form_winWordControl:" + runningNum.ToString();
            //return runningNum.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //int wordWnd = FindWindow(String.Empty, this.formName);
            //SetWindowPos(wordWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width - 20, this.Bounds.Height - 20, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME);

            this.Text = this.formName;

            //winWordControl1.PreActivate();
            //WinWordControl.WinWordControl.FindWindow(null, this.Text);
        }
    }
}
