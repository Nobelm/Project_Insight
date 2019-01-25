namespace Project_Insight
{
    partial class DB_Form
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
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.Eld_Grid = new System.Windows.Forms.DataGridView();
            this.timer_refresh = new System.Windows.Forms.Timer(this.components);
            this.Min_Grid = new System.Windows.Forms.DataGridView();
            this.Gen_Grid = new System.Windows.Forms.DataGridView();
            this.Cln_Grid = new System.Windows.Forms.DataGridView();
            this.btn_Hide = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.Eld_Grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Min_Grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gen_Grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Cln_Grid)).BeginInit();
            this.SuspendLayout();
            // 
            // Eld_Grid
            // 
            this.Eld_Grid.AllowUserToAddRows = false;
            this.Eld_Grid.AllowUserToDeleteRows = false;
            this.Eld_Grid.AllowUserToResizeColumns = false;
            this.Eld_Grid.AllowUserToResizeRows = false;
            this.Eld_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Eld_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Eld_Grid.Location = new System.Drawing.Point(12, 12);
            this.Eld_Grid.MultiSelect = false;
            this.Eld_Grid.Name = "Eld_Grid";
            this.Eld_Grid.ReadOnly = true;
            this.Eld_Grid.RowHeadersWidth = 20;
            this.Eld_Grid.Size = new System.Drawing.Size(474, 175);
            this.Eld_Grid.TabIndex = 0;
            // 
            // Min_Grid
            // 
            this.Min_Grid.AllowUserToAddRows = false;
            this.Min_Grid.AllowUserToDeleteRows = false;
            this.Min_Grid.AllowUserToResizeColumns = false;
            this.Min_Grid.AllowUserToResizeRows = false;
            this.Min_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Min_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Min_Grid.Location = new System.Drawing.Point(12, 193);
            this.Min_Grid.MultiSelect = false;
            this.Min_Grid.Name = "Min_Grid";
            this.Min_Grid.ReadOnly = true;
            this.Min_Grid.RowHeadersWidth = 20;
            this.Min_Grid.Size = new System.Drawing.Size(474, 204);
            this.Min_Grid.TabIndex = 2;
            // 
            // Gen_Grid
            // 
            this.Gen_Grid.AllowUserToAddRows = false;
            this.Gen_Grid.AllowUserToDeleteRows = false;
            this.Gen_Grid.AllowUserToResizeColumns = false;
            this.Gen_Grid.AllowUserToResizeRows = false;
            this.Gen_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Gen_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Gen_Grid.Location = new System.Drawing.Point(12, 403);
            this.Gen_Grid.MultiSelect = false;
            this.Gen_Grid.Name = "Gen_Grid";
            this.Gen_Grid.ReadOnly = true;
            this.Gen_Grid.RowHeadersWidth = 20;
            this.Gen_Grid.Size = new System.Drawing.Size(474, 321);
            this.Gen_Grid.TabIndex = 3;
            // 
            // Cln_Grid
            // 
            this.Cln_Grid.AllowUserToAddRows = false;
            this.Cln_Grid.AllowUserToDeleteRows = false;
            this.Cln_Grid.AllowUserToResizeColumns = false;
            this.Cln_Grid.AllowUserToResizeRows = false;
            this.Cln_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Cln_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Cln_Grid.Location = new System.Drawing.Point(502, 12);
            this.Cln_Grid.MultiSelect = false;
            this.Cln_Grid.Name = "Cln_Grid";
            this.Cln_Grid.ReadOnly = true;
            this.Cln_Grid.RowHeadersWidth = 20;
            this.Cln_Grid.Size = new System.Drawing.Size(145, 180);
            this.Cln_Grid.TabIndex = 4;
            // 
            // btn_Hide
            // 
            this.btn_Hide.Location = new System.Drawing.Point(511, 224);
            this.btn_Hide.Name = "btn_Hide";
            this.btn_Hide.Size = new System.Drawing.Size(75, 23);
            this.btn_Hide.TabIndex = 5;
            this.btn_Hide.Text = "Hide";
            this.btn_Hide.UseVisualStyleBackColor = true;
            // 
            // DB_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(659, 751);
            this.ControlBox = false;
            this.Controls.Add(this.btn_Hide);
            this.Controls.Add(this.Cln_Grid);
            this.Controls.Add(this.Gen_Grid);
            this.Controls.Add(this.Min_Grid);
            this.Controls.Add(this.Eld_Grid);
            this.MaximizeBox = false;
            this.Name = "DB_Form";
            this.Text = "DB_Form";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DB_Form_FormClosing);
            this.Load += new System.EventHandler(this.DB_Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Eld_Grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Min_Grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gen_Grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Cln_Grid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView Eld_Grid;
        private System.Windows.Forms.Timer timer_refresh;
        private System.Windows.Forms.DataGridView Min_Grid;
        private System.Windows.Forms.DataGridView Gen_Grid;
        private System.Windows.Forms.DataGridView Cln_Grid;
        private System.Windows.Forms.Button btn_Hide;
    }
}