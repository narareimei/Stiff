namespace Stiff
{
    partial class StiffForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);

            if (disposing && (stiffer != null))
            {
                stiffer.Dispose();
            }
            return;
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.bookGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.bookGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // bookGrid
            // 
            this.bookGrid.AllowUserToAddRows = false;
            this.bookGrid.AllowUserToDeleteRows = false;
            this.bookGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.bookGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.bookGrid.Location = new System.Drawing.Point(12, 56);
            this.bookGrid.Name = "bookGrid";
            this.bookGrid.ReadOnly = true;
            this.bookGrid.RowTemplate.Height = 21;
            this.bookGrid.Size = new System.Drawing.Size(839, 308);
            this.bookGrid.TabIndex = 1;
            // 
            // StiffForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(862, 378);
            this.Controls.Add(this.bookGrid);
            this.Controls.Add(this.button1);
            this.Name = "StiffForm";
            this.Text = "StiffForm";
            this.Load += new System.EventHandler(this.StiffForm_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.StiffForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.StiffForm_DragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.bookGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView bookGrid;
    }
}