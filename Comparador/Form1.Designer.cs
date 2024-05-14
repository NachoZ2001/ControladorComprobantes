namespace Comparador
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            textBoxAfip = new TextBox();
            textBoxHolistor = new TextBox();
            buttonAfip = new Button();
            buttonHolistor = new Button();
            buttonProcesar = new Button();
            pictureBoxLogoEstudio = new PictureBox();
            pictureBoxRuedaCargando = new PictureBox();
            buttonDefinirColumnas = new Button();
            textBoxTolerancia = new TextBox();
            numericUpDownTolerancia = new NumericUpDown();
            comboBoxEsquemas = new ComboBox();
            textBox1 = new TextBox();
            ((System.ComponentModel.ISupportInitialize)pictureBoxLogoEstudio).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBoxRuedaCargando).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownTolerancia).BeginInit();
            SuspendLayout();
            // 
            // textBoxAfip
            // 
            textBoxAfip.BackColor = Color.BlueViolet;
            textBoxAfip.BorderStyle = BorderStyle.FixedSingle;
            textBoxAfip.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxAfip.ForeColor = SystemColors.ButtonFace;
            textBoxAfip.Location = new Point(29, 23);
            textBoxAfip.Name = "textBoxAfip";
            textBoxAfip.ReadOnly = true;
            textBoxAfip.Size = new Size(163, 23);
            textBoxAfip.TabIndex = 0;
            textBoxAfip.Text = "Archivo AFIP";
            textBoxAfip.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxHolistor
            // 
            textBoxHolistor.BackColor = Color.BlueViolet;
            textBoxHolistor.BorderStyle = BorderStyle.FixedSingle;
            textBoxHolistor.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxHolistor.ForeColor = SystemColors.ButtonFace;
            textBoxHolistor.Location = new Point(276, 23);
            textBoxHolistor.Name = "textBoxHolistor";
            textBoxHolistor.ReadOnly = true;
            textBoxHolistor.Size = new Size(163, 23);
            textBoxHolistor.TabIndex = 1;
            textBoxHolistor.Text = "Archivo Contabilidad";
            textBoxHolistor.TextAlign = HorizontalAlignment.Center;
            // 
            // buttonAfip
            // 
            buttonAfip.BackColor = Color.BlueViolet;
            buttonAfip.FlatStyle = FlatStyle.Popup;
            buttonAfip.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonAfip.ForeColor = SystemColors.ButtonFace;
            buttonAfip.Location = new Point(29, 69);
            buttonAfip.Name = "buttonAfip";
            buttonAfip.Size = new Size(183, 23);
            buttonAfip.TabIndex = 2;
            buttonAfip.Text = "Seleccionar archivo AFIP";
            buttonAfip.UseVisualStyleBackColor = false;
            buttonAfip.Click += buttonAfip_Click;
            // 
            // buttonHolistor
            // 
            buttonHolistor.BackColor = Color.BlueViolet;
            buttonHolistor.FlatStyle = FlatStyle.Popup;
            buttonHolistor.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonHolistor.ForeColor = SystemColors.ButtonFace;
            buttonHolistor.Location = new Point(276, 69);
            buttonHolistor.Name = "buttonHolistor";
            buttonHolistor.Size = new Size(209, 23);
            buttonHolistor.TabIndex = 3;
            buttonHolistor.Text = "Seleccionar archivo Contabilidad";
            buttonHolistor.UseVisualStyleBackColor = false;
            buttonHolistor.Click += buttonHolistor_Click;
            // 
            // buttonProcesar
            // 
            buttonProcesar.BackColor = Color.BlueViolet;
            buttonProcesar.FlatStyle = FlatStyle.Popup;
            buttonProcesar.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonProcesar.ForeColor = SystemColors.ButtonFace;
            buttonProcesar.Location = new Point(531, 23);
            buttonProcesar.Name = "buttonProcesar";
            buttonProcesar.Size = new Size(155, 69);
            buttonProcesar.TabIndex = 4;
            buttonProcesar.Text = "Procesar";
            buttonProcesar.UseVisualStyleBackColor = false;
            buttonProcesar.Click += buttonProcesar_Click;
            // 
            // pictureBoxLogoEstudio
            // 
            pictureBoxLogoEstudio.BackColor = Color.Purple;
            pictureBoxLogoEstudio.BackgroundImageLayout = ImageLayout.None;
            pictureBoxLogoEstudio.Image = (Image)resources.GetObject("pictureBoxLogoEstudio.Image");
            pictureBoxLogoEstudio.Location = new Point(29, 268);
            pictureBoxLogoEstudio.Name = "pictureBoxLogoEstudio";
            pictureBoxLogoEstudio.Size = new Size(776, 213);
            pictureBoxLogoEstudio.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBoxLogoEstudio.TabIndex = 7;
            pictureBoxLogoEstudio.TabStop = false;
            // 
            // pictureBoxRuedaCargando
            // 
            pictureBoxRuedaCargando.BackColor = Color.Purple;
            pictureBoxRuedaCargando.Image = (Image)resources.GetObject("pictureBoxRuedaCargando.Image");
            pictureBoxRuedaCargando.Location = new Point(712, 23);
            pictureBoxRuedaCargando.Name = "pictureBoxRuedaCargando";
            pictureBoxRuedaCargando.Size = new Size(93, 69);
            pictureBoxRuedaCargando.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBoxRuedaCargando.TabIndex = 8;
            pictureBoxRuedaCargando.TabStop = false;
            // 
            // buttonDefinirColumnas
            // 
            buttonDefinirColumnas.BackColor = Color.BlueViolet;
            buttonDefinirColumnas.FlatStyle = FlatStyle.Popup;
            buttonDefinirColumnas.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonDefinirColumnas.ForeColor = SystemColors.ButtonFace;
            buttonDefinirColumnas.Location = new Point(276, 111);
            buttonDefinirColumnas.Name = "buttonDefinirColumnas";
            buttonDefinirColumnas.Size = new Size(209, 23);
            buttonDefinirColumnas.TabIndex = 9;
            buttonDefinirColumnas.Text = "Crear esquema";
            buttonDefinirColumnas.UseVisualStyleBackColor = false;
            buttonDefinirColumnas.Click += buttonDefinirColumnas_Click;
            // 
            // textBoxTolerancia
            // 
            textBoxTolerancia.BackColor = Color.BlueViolet;
            textBoxTolerancia.BorderStyle = BorderStyle.FixedSingle;
            textBoxTolerancia.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxTolerancia.ForeColor = SystemColors.ButtonFace;
            textBoxTolerancia.Location = new Point(276, 220);
            textBoxTolerancia.Name = "textBoxTolerancia";
            textBoxTolerancia.ReadOnly = true;
            textBoxTolerancia.Size = new Size(209, 23);
            textBoxTolerancia.TabIndex = 11;
            textBoxTolerancia.Text = "Tolerancia";
            textBoxTolerancia.TextAlign = HorizontalAlignment.Center;
            // 
            // numericUpDownTolerancia
            // 
            numericUpDownTolerancia.BackColor = Color.BlueViolet;
            numericUpDownTolerancia.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownTolerancia.ForeColor = SystemColors.ButtonFace;
            numericUpDownTolerancia.Location = new Point(531, 220);
            numericUpDownTolerancia.Name = "numericUpDownTolerancia";
            numericUpDownTolerancia.Size = new Size(120, 23);
            numericUpDownTolerancia.TabIndex = 12;
            numericUpDownTolerancia.TextAlign = HorizontalAlignment.Center;
            numericUpDownTolerancia.UpDownAlign = LeftRightAlignment.Left;
            // 
            // comboBoxEsquemas
            // 
            comboBoxEsquemas.BackColor = Color.BlueViolet;
            comboBoxEsquemas.FlatStyle = FlatStyle.Popup;
            comboBoxEsquemas.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            comboBoxEsquemas.ForeColor = SystemColors.ButtonFace;
            comboBoxEsquemas.FormattingEnabled = true;
            comboBoxEsquemas.Location = new Point(276, 173);
            comboBoxEsquemas.Name = "comboBoxEsquemas";
            comboBoxEsquemas.RightToLeft = RightToLeft.No;
            comboBoxEsquemas.Size = new Size(209, 23);
            comboBoxEsquemas.TabIndex = 13;
            comboBoxEsquemas.UseWaitCursor = true;
            // 
            // textBox1
            // 
            textBox1.BackColor = Color.Purple;
            textBox1.BorderStyle = BorderStyle.None;
            textBox1.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point);
            textBox1.ForeColor = SystemColors.ButtonFace;
            textBox1.Location = new Point(293, 147);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(192, 20);
            textBox1.TabIndex = 14;
            textBox1.Text = "Seleccionar un esquema";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(832, 493);
            Controls.Add(textBox1);
            Controls.Add(comboBoxEsquemas);
            Controls.Add(numericUpDownTolerancia);
            Controls.Add(textBoxTolerancia);
            Controls.Add(buttonDefinirColumnas);
            Controls.Add(pictureBoxRuedaCargando);
            Controls.Add(pictureBoxLogoEstudio);
            Controls.Add(buttonProcesar);
            Controls.Add(buttonHolistor);
            Controls.Add(buttonAfip);
            Controls.Add(textBoxHolistor);
            Controls.Add(textBoxAfip);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBoxLogoEstudio).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBoxRuedaCargando).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownTolerancia).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBoxAfip;
        private TextBox textBoxHolistor;
        private Button buttonAfip;
        private Button buttonHolistor;
        private Button buttonProcesar;
        private PictureBox pictureBoxLogoEstudio;
        private PictureBox pictureBoxRuedaCargando;
        private Button buttonDefinirColumnas;
        private TextBox textBoxTolerancia;
        private NumericUpDown numericUpDownTolerancia;
        private ComboBox comboBoxEsquemas;
        private TextBox textBox1;
    }
}