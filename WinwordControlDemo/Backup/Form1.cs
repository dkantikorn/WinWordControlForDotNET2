using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace WinWordTestApp
{
	/// <summary>
	/// just testing
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button load;
		private WinWordControl.WinWordControl winWordControl1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button Restore;
		private System.Windows.Forms.Button close;
		/// <summary>
		/// needed designer variable
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			InitializeComponent();
		}

		/// <summary>
		/// cleanuup ressources
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			// just to be shure!
			winWordControl1.CloseControl();

			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		///
		/// </summary>
		private void InitializeComponent()
		{
			this.load = new System.Windows.Forms.Button();
			this.winWordControl1 = new WinWordControl.WinWordControl();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.button1 = new System.Windows.Forms.Button();
			this.Restore = new System.Windows.Forms.Button();
			this.close = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// load
			// 
			this.load.Location = new System.Drawing.Point(592, 24);
			this.load.Name = "load";
			this.load.Size = new System.Drawing.Size(56, 32);
			this.load.TabIndex = 1;
			this.load.Text = "load";
			this.load.Click += new System.EventHandler(this.load_Click);
			// 
			// winWordControl1
			// 
			this.winWordControl1.Dock = System.Windows.Forms.DockStyle.Left;
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
			this.button1.Size = new System.Drawing.Size(56, 32);
			this.button1.TabIndex = 3;
			this.button1.Text = "PreActivate";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Restore
			// 
			this.Restore.Location = new System.Drawing.Point(592, 208);
			this.Restore.Name = "Restore";
			this.Restore.Size = new System.Drawing.Size(56, 32);
			this.Restore.TabIndex = 4;
			this.Restore.Text = "Restore Word";
			this.Restore.Click += new System.EventHandler(this.Restore_Click);
			// 
			// close
			// 
			this.close.Location = new System.Drawing.Point(592, 72);
			this.close.Name = "close";
			this.close.Size = new System.Drawing.Size(56, 32);
			this.close.TabIndex = 5;
			this.close.Text = "Close";
			this.close.Click += new System.EventHandler(this.close_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(672, 389);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.close,
																		  this.Restore,
																		  this.button1,
																		  this.winWordControl1,
																		  this.load});
			this.Name = "Form1";
			this.Text = "Form1";
			this.Activated += new System.EventHandler(this.OnActivate);
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
			
		}

		private void load_Click(object sender, System.EventArgs e)
		{
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				winWordControl1.LoadDocument(openFileDialog1.FileName);
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
	}
}
