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
            this.btUnification = new System.Windows.Forms.Button();
            this.bookGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.bookGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // btUnification
            // 
            this.btUnification.Location = new System.Drawing.Point(12, 12);
            this.btUnification.Name = "btUnification";
            this.btUnification.Size = new System.Drawing.Size(75, 23);
            this.btUnification.TabIndex = 0;
            this.btUnification.Text = "設定更新";
            this.btUnification.UseVisualStyleBackColor = true;
            this.btUnification.Click += new System.EventHandler(this.btUnification_Click);
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
            this.bookGrid.Size = new System.Drawing.Size(1041, 308);
            this.bookGrid.TabIndex = 1;
            // 
            // StiffForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1064, 378);
            this.Controls.Add(this.bookGrid);
            this.Controls.Add(this.btUnification);
            this.Name = "StiffForm";
            this.Text = "StiffForm";
            this.Load += new System.EventHandler(this.StiffForm_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.StiffForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.StiffForm_DragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.bookGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btUnification;
        private System.Windows.Forms.DataGridView bookGrid;
    }
}