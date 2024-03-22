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
            ((System.ComponentModel.ISupportInitialize)pictureBoxLogoEstudio).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBoxRuedaCargando).BeginInit();
            SuspendLayout();
            // 
            // textBoxAfip
            // 
            textBoxAfip.BackColor = Color.BlueViolet;
            textBoxAfip.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxAfip.ForeColor = SystemColors.ButtonFace;
            textBoxAfip.Location = new Point(29, 23);
            textBoxAfip.Name = "textBoxAfip";
            textBoxAfip.Size = new Size(154, 23);
            textBoxAfip.TabIndex = 0;
            textBoxAfip.Text = "Archivo AFIP";
            textBoxAfip.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxHolistor
            // 
            textBoxHolistor.BackColor = Color.BlueViolet;
            textBoxHolistor.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxHolistor.ForeColor = SystemColors.ButtonFace;
            textBoxHolistor.Location = new Point(276, 23);
            textBoxHolistor.Name = "textBoxHolistor";
            textBoxHolistor.Size = new Size(163, 23);
            textBoxHolistor.TabIndex = 1;
            textBoxHolistor.Text = "Archivo Holistor";
            textBoxHolistor.TextAlign = HorizontalAlignment.Center;
            // 
            // buttonAfip
            // 
            buttonAfip.BackColor = Color.BlueViolet;
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
            buttonHolistor.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonHolistor.ForeColor = SystemColors.ButtonFace;
            buttonHolistor.Location = new Point(276, 69);
            buttonHolistor.Name = "buttonHolistor";
            buttonHolistor.Size = new Size(183, 23);
            buttonHolistor.TabIndex = 3;
            buttonHolistor.Text = "Seleccionar archivo Holistor";
            buttonHolistor.UseVisualStyleBackColor = false;
            buttonHolistor.Click += buttonHolistor_Click;
            // 
            // buttonProcesar
            // 
            buttonProcesar.BackColor = Color.BlueViolet;
            buttonProcesar.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonProcesar.ForeColor = SystemColors.ButtonFace;
            buttonProcesar.Location = new Point(491, 23);
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
            pictureBoxLogoEstudio.Location = new Point(29, 117);
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
            pictureBoxRuedaCargando.Location = new Point(689, 23);
            pictureBoxRuedaCargando.Name = "pictureBoxRuedaCargando";
            pictureBoxRuedaCargando.Size = new Size(93, 69);
            pictureBoxRuedaCargando.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBoxRuedaCargando.TabIndex = 8;
            pictureBoxRuedaCargando.TabStop = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(822, 332);
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
    }
}