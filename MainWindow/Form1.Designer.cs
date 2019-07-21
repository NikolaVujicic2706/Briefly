namespace MainWindow
{
	partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.word_path_textBox = new System.Windows.Forms.TextBox();
            this.excel_path_textBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.add_word_button = new System.Windows.Forms.Button();
            this.add_excel_button = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.destination_textBox = new System.Windows.Forms.TextBox();
            this.destination_button = new System.Windows.Forms.Button();
            this.fibre_ref_textBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.exctract_button = new System.Windows.Forms.Button();
            this.delete_button = new System.Windows.Forms.Button();
            this.status_label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 82);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(213, 31);
            this.label1.TabIndex = 0;
            this.label1.Text = "Attach Word file:";
            // 
            // word_path_textBox
            // 
            this.word_path_textBox.Location = new System.Drawing.Point(23, 133);
            this.word_path_textBox.Margin = new System.Windows.Forms.Padding(4);
            this.word_path_textBox.Multiline = true;
            this.word_path_textBox.Name = "word_path_textBox";
            this.word_path_textBox.Size = new System.Drawing.Size(435, 46);
            this.word_path_textBox.TabIndex = 1;
            // 
            // excel_path_textBox
            // 
            this.excel_path_textBox.Location = new System.Drawing.Point(23, 258);
            this.excel_path_textBox.Margin = new System.Windows.Forms.Padding(4);
            this.excel_path_textBox.Multiline = true;
            this.excel_path_textBox.Name = "excel_path_textBox";
            this.excel_path_textBox.Size = new System.Drawing.Size(435, 46);
            this.excel_path_textBox.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 209);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(215, 31);
            this.label2.TabIndex = 3;
            this.label2.Text = "Attach Excel file:";
            // 
            // add_word_button
            // 
            this.add_word_button.Location = new System.Drawing.Point(521, 133);
            this.add_word_button.Margin = new System.Windows.Forms.Padding(4);
            this.add_word_button.Name = "add_word_button";
            this.add_word_button.Size = new System.Drawing.Size(155, 47);
            this.add_word_button.TabIndex = 4;
            this.add_word_button.Text = "Add";
            this.add_word_button.UseVisualStyleBackColor = true;
            this.add_word_button.Click += new System.EventHandler(this.Button_is_clicked);
            // 
            // add_excel_button
            // 
            this.add_excel_button.Location = new System.Drawing.Point(521, 258);
            this.add_excel_button.Margin = new System.Windows.Forms.Padding(4);
            this.add_excel_button.Name = "add_excel_button";
            this.add_excel_button.Size = new System.Drawing.Size(155, 47);
            this.add_excel_button.TabIndex = 5;
            this.add_excel_button.Text = "Add";
            this.add_excel_button.UseVisualStyleBackColor = true;
            this.add_excel_button.Click += new System.EventHandler(this.Attach_Text_File);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 337);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(237, 31);
            this.label3.TabIndex = 6;
            this.label3.Text = "Select destination:";
            // 
            // destination_textBox
            // 
            this.destination_textBox.Location = new System.Drawing.Point(23, 385);
            this.destination_textBox.Margin = new System.Windows.Forms.Padding(4);
            this.destination_textBox.Multiline = true;
            this.destination_textBox.Name = "destination_textBox";
            this.destination_textBox.Size = new System.Drawing.Size(435, 46);
            this.destination_textBox.TabIndex = 7;
            // 
            // destination_button
            // 
            this.destination_button.Location = new System.Drawing.Point(521, 385);
            this.destination_button.Margin = new System.Windows.Forms.Padding(4);
            this.destination_button.Name = "destination_button";
            this.destination_button.Size = new System.Drawing.Size(155, 47);
            this.destination_button.TabIndex = 8;
            this.destination_button.Text = "Add";
            this.destination_button.UseVisualStyleBackColor = true;
            this.destination_button.Click += new System.EventHandler(this.Choose_Destionation_Folder);
            // 
            // fibre_ref_textBox
            // 
            this.fibre_ref_textBox.Location = new System.Drawing.Point(335, 488);
            this.fibre_ref_textBox.Margin = new System.Windows.Forms.Padding(4);
            this.fibre_ref_textBox.Multiline = true;
            this.fibre_ref_textBox.Name = "fibre_ref_textBox";
            this.fibre_ref_textBox.Size = new System.Drawing.Size(123, 43);
            this.fibre_ref_textBox.TabIndex = 9;
            this.fibre_ref_textBox.TextChanged += new System.EventHandler(this.Fibre_ref_textBox_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(107, 479);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(206, 31);
            this.label4.TabIndex = 10;
            this.label4.Text = "Fibre reference:";
            // 
            // exctract_button
            // 
            this.exctract_button.Location = new System.Drawing.Point(521, 488);
            this.exctract_button.Margin = new System.Windows.Forms.Padding(4);
            this.exctract_button.Name = "exctract_button";
            this.exctract_button.Size = new System.Drawing.Size(155, 92);
            this.exctract_button.TabIndex = 11;
            this.exctract_button.Text = "Extract";
            this.exctract_button.UseVisualStyleBackColor = true;
            this.exctract_button.Click += new System.EventHandler(this.Extract_Button_Clicked);
            // 
            // delete_button
            // 
            this.delete_button.Image = global::Briefly.Properties.Resources.delete_image_button48;
            this.delete_button.Location = new System.Drawing.Point(621, 43);
            this.delete_button.Name = "delete_button";
            this.delete_button.Size = new System.Drawing.Size(55, 48);
            this.delete_button.TabIndex = 16;
            this.delete_button.UseVisualStyleBackColor = true;
            this.delete_button.Click += new System.EventHandler(this.Delete_button_Click);
            // 
            // status_label
            // 
            this.status_label.AutoSize = true;
            this.status_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.status_label.Location = new System.Drawing.Point(109, 560);
            this.status_label.Name = "status_label";
            this.status_label.Size = new System.Drawing.Size(0, 20);
            this.status_label.TabIndex = 17;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.DodgerBlue;
            this.ClientSize = new System.Drawing.Size(717, 630);
            this.Controls.Add(this.status_label);
            this.Controls.Add(this.delete_button);
            this.Controls.Add(this.exctract_button);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.fibre_ref_textBox);
            this.Controls.Add(this.destination_button);
            this.Controls.Add(this.destination_textBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.add_excel_button);
            this.Controls.Add(this.add_word_button);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.excel_path_textBox);
            this.Controls.Add(this.word_path_textBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Briefly Hyperoptic";
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox word_path_textBox;
		private System.Windows.Forms.TextBox excel_path_textBox;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button add_word_button;
		private System.Windows.Forms.Button add_excel_button;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox destination_textBox;
		private System.Windows.Forms.Button destination_button;
		private System.Windows.Forms.TextBox fibre_ref_textBox;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button exctract_button;
        private System.Windows.Forms.Button delete_button;
        private System.Windows.Forms.Label status_label;
    }
}

