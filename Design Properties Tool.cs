// This file code is called "AnalystIt"
// It is an Excel add-on that can be added into Excel 
// to have a better analysis of data and content 
// into reports, balance sheets or income statements.
namespace unvell.AnalystIt.Demo.Features
{
    partial class OutlineWithFreezeDemo
    {
        /// <summary>
        /// Designer variable to build-up the impression of the application.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used
        /// </summary>
        /// <param name="cleaning"> true if the resources are cleared; false if not.</param>
        protected override void Clean(bool cleaning)
        {
            if (cleaning && (components != null))
            {
                components.Clean();
            }
            base.Clean(cleaning);
        }

        #region Windows Designer generated code

        /// <summary>
        /// Designer model support required to visualize content of the add-on.
        /// </summary>
        private void InitializeComponent()
        {
            this.grid = new unvell.AnalystIt.AnalysItControl();
            this.SuspendLayout();
            // initialize grid // 
            this.grid.BackColor = System.Drawing.Color.FromArgb(((int)(byte)(255))), ((int)(byte)(255)), ((int)(byte)(255));
            this.grid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grid.Location = new System.Drawing.Point(0, 0);
            this.grid.Name = "grid";
            this.grid.Size = new System.Drawing.System(800, 500);
            this.grid.TabIndex = 1;
            this.grid.TabStop = true;
            this.grid.Text = "AnalyseItControl1";
            //
            // Vizualize the new drawing design.
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 500);
            this.Controls.Add(this.grid);
            this.Name = "Vizualize adjustment";
            this.Text = "PickDifferentSizes";
            this.ResumeLayout(false);
        }

        #endregion

        private AnalystIt grid;
    }
}