namespace AccountingOfCourtCases
{
    partial class Регистрация
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
            System.Windows.Forms.Label логинLabel;
            System.Windows.Forms.Label фИОLabel;
            System.Windows.Forms.Label парольLabel;
            System.Windows.Forms.Label emailLabel;
            System.Windows.Forms.Label телефонLabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Регистрация));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.логинTextBox = new System.Windows.Forms.TextBox();
            this.фИОTextBox = new System.Windows.Forms.TextBox();
            this.парольTextBox = new System.Windows.Forms.TextBox();
            this.emailTextBox = new System.Windows.Forms.TextBox();
            this.телефонTextBox = new System.Windows.Forms.TextBox();
            логинLabel = new System.Windows.Forms.Label();
            фИОLabel = new System.Windows.Forms.Label();
            парольLabel = new System.Windows.Forms.Label();
            emailLabel = new System.Windows.Forms.Label();
            телефонLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // логинLabel
            // 
            логинLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            логинLabel.AutoSize = true;
            логинLabel.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            логинLabel.Location = new System.Drawing.Point(45, 154);
            логинLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            логинLabel.Name = "логинLabel";
            логинLabel.Size = new System.Drawing.Size(83, 24);
            логинLabel.TabIndex = 2;
            логинLabel.Text = "Логин:*";
            // 
            // фИОLabel
            // 
            фИОLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            фИОLabel.AutoSize = true;
            фИОLabel.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            фИОLabel.Location = new System.Drawing.Point(45, 207);
            фИОLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            фИОLabel.Name = "фИОLabel";
            фИОLabel.Size = new System.Drawing.Size(64, 24);
            фИОLabel.TabIndex = 4;
            фИОLabel.Text = "ФИО:";
            // 
            // парольLabel
            // 
            парольLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            парольLabel.AutoSize = true;
            парольLabel.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            парольLabel.Location = new System.Drawing.Point(45, 257);
            парольLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            парольLabel.Name = "парольLabel";
            парольLabel.Size = new System.Drawing.Size(99, 24);
            парольLabel.TabIndex = 6;
            парольLabel.Text = "Пароль:*";
            // 
            // emailLabel
            // 
            emailLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            emailLabel.AutoSize = true;
            emailLabel.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            emailLabel.Location = new System.Drawing.Point(45, 314);
            emailLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            emailLabel.Name = "emailLabel";
            emailLabel.Size = new System.Drawing.Size(68, 24);
            emailLabel.TabIndex = 10;
            emailLabel.Text = "Email:";
            // 
            // телефонLabel
            // 
            телефонLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            телефонLabel.AutoSize = true;
            телефонLabel.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            телефонLabel.Location = new System.Drawing.Point(45, 362);
            телефонLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            телефонLabel.Name = "телефонLabel";
            телефонLabel.Size = new System.Drawing.Size(105, 24);
            телефонLabel.TabIndex = 12;
            телефонLabel.Text = "Телефон:";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.DarkGray;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(логинLabel);
            this.panel1.Controls.Add(this.логинTextBox);
            this.panel1.Controls.Add(фИОLabel);
            this.panel1.Controls.Add(this.фИОTextBox);
            this.panel1.Controls.Add(парольLabel);
            this.panel1.Controls.Add(this.парольTextBox);
            this.panel1.Controls.Add(emailLabel);
            this.panel1.Controls.Add(this.emailTextBox);
            this.panel1.Controls.Add(телефонLabel);
            this.panel1.Controls.Add(this.телефонTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(499, 682);
            this.panel1.TabIndex = 2;
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Panel1_MouseMove);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(93, 447);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(312, 25);
            this.label1.TabIndex = 41;
            this.label1.Text = "*Заполните обязательные поля";
            this.label1.Visible = false;
            // 
            // button2
            // 
            this.button2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.button2.Location = new System.Drawing.Point(15, 36);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(100, 28);
            this.button2.TabIndex = 7;
            this.button2.Text = "Назад";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.button1.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(125, 496);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(247, 60);
            this.button1.TabIndex = 6;
            this.button1.Text = "Регистрация";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // логинTextBox
            // 
            this.логинTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.логинTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.логинTextBox.Location = new System.Drawing.Point(229, 153);
            this.логинTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.логинTextBox.Name = "логинTextBox";
            this.логинTextBox.Size = new System.Drawing.Size(224, 30);
            this.логинTextBox.TabIndex = 1;
            // 
            // фИОTextBox
            // 
            this.фИОTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.фИОTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.фИОTextBox.Location = new System.Drawing.Point(229, 205);
            this.фИОTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.фИОTextBox.Name = "фИОTextBox";
            this.фИОTextBox.Size = new System.Drawing.Size(224, 30);
            this.фИОTextBox.TabIndex = 2;
            // 
            // парольTextBox
            // 
            this.парольTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.парольTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.парольTextBox.Location = new System.Drawing.Point(229, 256);
            this.парольTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.парольTextBox.Name = "парольTextBox";
            this.парольTextBox.Size = new System.Drawing.Size(224, 30);
            this.парольTextBox.TabIndex = 3;
            this.парольTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ПарольTextBox_KeyPress);
            // 
            // emailTextBox
            // 
            this.emailTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.emailTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.emailTextBox.Location = new System.Drawing.Point(229, 309);
            this.emailTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.emailTextBox.Name = "emailTextBox";
            this.emailTextBox.Size = new System.Drawing.Size(224, 30);
            this.emailTextBox.TabIndex = 4;
            // 
            // телефонTextBox
            // 
            this.телефонTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.телефонTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.телефонTextBox.Location = new System.Drawing.Point(229, 361);
            this.телефонTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.телефонTextBox.Name = "телефонTextBox";
            this.телефонTextBox.Size = new System.Drawing.Size(224, 30);
            this.телефонTextBox.TabIndex = 5;
            // 
            // Регистрация
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(499, 682);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Регистрация";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Регистрация";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Регистрация_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox логинTextBox;
        private System.Windows.Forms.TextBox фИОTextBox;
        private System.Windows.Forms.TextBox парольTextBox;
        private System.Windows.Forms.TextBox emailTextBox;
        private System.Windows.Forms.TextBox телефонTextBox;
    }
}