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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DB_Form));
            this.Eld_Grid = new System.Windows.Forms.DataGridView();
            this.Min_Grid = new System.Windows.Forms.DataGridView();
            this.Gen_Grid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.Eld_Grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Min_Grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gen_Grid)).BeginInit();
            this.SuspendLayout();
            // 
            // Eld_Grid
            // 
            resources.ApplyResources(this.Eld_Grid, "Eld_Grid");
            this.Eld_Grid.AllowUserToAddRows = false;
            this.Eld_Grid.AllowUserToDeleteRows = false;
            this.Eld_Grid.AllowUserToResizeColumns = false;
            this.Eld_Grid.AllowUserToResizeRows = false;
            this.Eld_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Eld_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Eld_Grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.Eld_Grid.MultiSelect = false;
            this.Eld_Grid.Name = "Eld_Grid";
            this.Eld_Grid.ReadOnly = true;
            // 
            // Min_Grid
            // 
            resources.ApplyResources(this.Min_Grid, "Min_Grid");
            this.Min_Grid.AllowUserToAddRows = false;
            this.Min_Grid.AllowUserToDeleteRows = false;
            this.Min_Grid.AllowUserToResizeColumns = false;
            this.Min_Grid.AllowUserToResizeRows = false;
            this.Min_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Min_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Min_Grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.Min_Grid.MultiSelect = false;
            this.Min_Grid.Name = "Min_Grid";
            this.Min_Grid.ReadOnly = true;
            // 
            // Gen_Grid
            // 
            resources.ApplyResources(this.Gen_Grid, "Gen_Grid");
            this.Gen_Grid.AllowUserToAddRows = false;
            this.Gen_Grid.AllowUserToDeleteRows = false;
            this.Gen_Grid.AllowUserToResizeColumns = false;
            this.Gen_Grid.AllowUserToResizeRows = false;
            this.Gen_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.Gen_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Gen_Grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.Gen_Grid.MultiSelect = false;
            this.Gen_Grid.Name = "Gen_Grid";
            this.Gen_Grid.ReadOnly = true;
            // 
            // DB_Form
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ControlBox = false;
            this.Controls.Add(this.Gen_Grid);
            this.Controls.Add(this.Min_Grid);
            this.Controls.Add(this.Eld_Grid);
            this.MaximizeBox = false;
            this.Name = "DB_Form";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DB_Form_FormClosing);
            this.Load += new System.EventHandler(this.DB_Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Eld_Grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Min_Grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gen_Grid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView Eld_Grid;
        private System.Windows.Forms.DataGridView Min_Grid;
        private System.Windows.Forms.DataGridView Gen_Grid;
    }
}