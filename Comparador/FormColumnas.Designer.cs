namespace Comparador
{
    partial class FormColumnas
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
            numericUpDownCuit = new NumericUpDown();
            textBoxColumna = new TextBox();
            textBoxNumeroColumna = new TextBox();
            textBoxCUIT = new TextBox();
            textBoxPuntoVenta = new TextBox();
            numericUpDownPuntoVenta = new NumericUpDown();
            textBoxNumeroComprobante = new TextBox();
            numericUpDownNumeroComprobante = new NumericUpDown();
            textBoxIVA = new TextBox();
            numericUpDownIVA = new NumericUpDown();
            textBoxTotal = new TextBox();
            numericUpDownTotal = new NumericUpDown();
            buttonGuardar = new Button();
            buttonCancelar = new Button();
            ((System.ComponentModel.ISupportInitialize)numericUpDownCuit).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownPuntoVenta).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownNumeroComprobante).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownIVA).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownTotal).BeginInit();
            SuspendLayout();
            // 
            // numericUpDownCuit
            // 
            numericUpDownCuit.BackColor = Color.BlueViolet;
            numericUpDownCuit.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownCuit.ForeColor = SystemColors.ButtonFace;
            numericUpDownCuit.Location = new Point(164, 50);
            numericUpDownCuit.Minimum = new decimal(new int[] { 100, 0, 0, int.MinValue });
            numericUpDownCuit.Name = "numericUpDownCuit";
            numericUpDownCuit.Size = new Size(120, 23);
            numericUpDownCuit.TabIndex = 0;
            numericUpDownCuit.TextAlign = HorizontalAlignment.Center;
            numericUpDownCuit.Value = new decimal(new int[] { 1, 0, 0, int.MinValue });
            // 
            // textBoxColumna
            // 
            textBoxColumna.BackColor = Color.BlueViolet;
            textBoxColumna.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxColumna.ForeColor = SystemColors.ButtonFace;
            textBoxColumna.Location = new Point(12, 12);
            textBoxColumna.Name = "textBoxColumna";
            textBoxColumna.ReadOnly = true;
            textBoxColumna.Size = new Size(133, 23);
            textBoxColumna.TabIndex = 1;
            textBoxColumna.Text = "Columna";
            textBoxColumna.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxNumeroColumna
            // 
            textBoxNumeroColumna.BackColor = Color.BlueViolet;
            textBoxNumeroColumna.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxNumeroColumna.ForeColor = SystemColors.ButtonFace;
            textBoxNumeroColumna.Location = new Point(164, 12);
            textBoxNumeroColumna.Name = "textBoxNumeroColumna";
            textBoxNumeroColumna.ReadOnly = true;
            textBoxNumeroColumna.Size = new Size(119, 23);
            textBoxNumeroColumna.TabIndex = 2;
            textBoxNumeroColumna.Text = "Índice";
            textBoxNumeroColumna.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxCUIT
            // 
            textBoxCUIT.BackColor = Color.BlueViolet;
            textBoxCUIT.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxCUIT.ForeColor = SystemColors.ButtonFace;
            textBoxCUIT.Location = new Point(12, 50);
            textBoxCUIT.Name = "textBoxCUIT";
            textBoxCUIT.ReadOnly = true;
            textBoxCUIT.Size = new Size(133, 23);
            textBoxCUIT.TabIndex = 3;
            textBoxCUIT.Text = "Cuit";
            textBoxCUIT.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxPuntoVenta
            // 
            textBoxPuntoVenta.BackColor = Color.BlueViolet;
            textBoxPuntoVenta.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxPuntoVenta.ForeColor = SystemColors.ButtonFace;
            textBoxPuntoVenta.Location = new Point(12, 90);
            textBoxPuntoVenta.Name = "textBoxPuntoVenta";
            textBoxPuntoVenta.ReadOnly = true;
            textBoxPuntoVenta.Size = new Size(133, 23);
            textBoxPuntoVenta.TabIndex = 4;
            textBoxPuntoVenta.Text = "Punto venta";
            textBoxPuntoVenta.TextAlign = HorizontalAlignment.Center;
            // 
            // numericUpDownPuntoVenta
            // 
            numericUpDownPuntoVenta.BackColor = Color.BlueViolet;
            numericUpDownPuntoVenta.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownPuntoVenta.ForeColor = SystemColors.ButtonFace;
            numericUpDownPuntoVenta.Location = new Point(164, 90);
            numericUpDownPuntoVenta.Minimum = new decimal(new int[] { 100, 0, 0, int.MinValue });
            numericUpDownPuntoVenta.Name = "numericUpDownPuntoVenta";
            numericUpDownPuntoVenta.Size = new Size(120, 23);
            numericUpDownPuntoVenta.TabIndex = 5;
            numericUpDownPuntoVenta.TextAlign = HorizontalAlignment.Center;
            numericUpDownPuntoVenta.Value = new decimal(new int[] { 1, 0, 0, int.MinValue });
            numericUpDownPuntoVenta.ValueChanged += numericUpDownPuntoVenta_ValueChanged;
            // 
            // textBoxNumeroComprobante
            // 
            textBoxNumeroComprobante.BackColor = Color.BlueViolet;
            textBoxNumeroComprobante.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxNumeroComprobante.ForeColor = SystemColors.ButtonFace;
            textBoxNumeroComprobante.Location = new Point(12, 131);
            textBoxNumeroComprobante.Name = "textBoxNumeroComprobante";
            textBoxNumeroComprobante.ReadOnly = true;
            textBoxNumeroComprobante.Size = new Size(133, 23);
            textBoxNumeroComprobante.TabIndex = 6;
            textBoxNumeroComprobante.Text = "Comprobante";
            textBoxNumeroComprobante.TextAlign = HorizontalAlignment.Center;
            // 
            // numericUpDownNumeroComprobante
            // 
            numericUpDownNumeroComprobante.BackColor = Color.BlueViolet;
            numericUpDownNumeroComprobante.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownNumeroComprobante.ForeColor = SystemColors.ButtonFace;
            numericUpDownNumeroComprobante.Location = new Point(164, 131);
            numericUpDownNumeroComprobante.Minimum = new decimal(new int[] { 100, 0, 0, int.MinValue });
            numericUpDownNumeroComprobante.Name = "numericUpDownNumeroComprobante";
            numericUpDownNumeroComprobante.Size = new Size(120, 23);
            numericUpDownNumeroComprobante.TabIndex = 7;
            numericUpDownNumeroComprobante.TextAlign = HorizontalAlignment.Center;
            numericUpDownNumeroComprobante.Value = new decimal(new int[] { 1, 0, 0, int.MinValue });
            // 
            // textBoxIVA
            // 
            textBoxIVA.BackColor = Color.BlueViolet;
            textBoxIVA.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxIVA.ForeColor = SystemColors.ButtonFace;
            textBoxIVA.Location = new Point(12, 170);
            textBoxIVA.Name = "textBoxIVA";
            textBoxIVA.ReadOnly = true;
            textBoxIVA.Size = new Size(133, 23);
            textBoxIVA.TabIndex = 8;
            textBoxIVA.Text = "IVA";
            textBoxIVA.TextAlign = HorizontalAlignment.Center;
            // 
            // numericUpDownIVA
            // 
            numericUpDownIVA.BackColor = Color.BlueViolet;
            numericUpDownIVA.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownIVA.ForeColor = SystemColors.ButtonFace;
            numericUpDownIVA.Location = new Point(164, 170);
            numericUpDownIVA.Minimum = new decimal(new int[] { 100, 0, 0, int.MinValue });
            numericUpDownIVA.Name = "numericUpDownIVA";
            numericUpDownIVA.Size = new Size(120, 23);
            numericUpDownIVA.TabIndex = 9;
            numericUpDownIVA.TextAlign = HorizontalAlignment.Center;
            numericUpDownIVA.Value = new decimal(new int[] { 1, 0, 0, int.MinValue });
            // 
            // textBoxTotal
            // 
            textBoxTotal.BackColor = Color.BlueViolet;
            textBoxTotal.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxTotal.ForeColor = SystemColors.ButtonFace;
            textBoxTotal.Location = new Point(12, 211);
            textBoxTotal.Name = "textBoxTotal";
            textBoxTotal.ReadOnly = true;
            textBoxTotal.Size = new Size(133, 23);
            textBoxTotal.TabIndex = 10;
            textBoxTotal.Text = "Total";
            textBoxTotal.TextAlign = HorizontalAlignment.Center;
            // 
            // numericUpDownTotal
            // 
            numericUpDownTotal.BackColor = Color.BlueViolet;
            numericUpDownTotal.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            numericUpDownTotal.ForeColor = SystemColors.ButtonFace;
            numericUpDownTotal.Location = new Point(164, 212);
            numericUpDownTotal.Minimum = new decimal(new int[] { 100, 0, 0, int.MinValue });
            numericUpDownTotal.Name = "numericUpDownTotal";
            numericUpDownTotal.Size = new Size(120, 23);
            numericUpDownTotal.TabIndex = 11;
            numericUpDownTotal.TextAlign = HorizontalAlignment.Center;
            numericUpDownTotal.Value = new decimal(new int[] { 1, 0, 0, int.MinValue });
            // 
            // buttonGuardar
            // 
            buttonGuardar.BackColor = Color.BlueViolet;
            buttonGuardar.FlatStyle = FlatStyle.Flat;
            buttonGuardar.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonGuardar.ForeColor = SystemColors.ButtonFace;
            buttonGuardar.Location = new Point(164, 264);
            buttonGuardar.Name = "buttonGuardar";
            buttonGuardar.Size = new Size(75, 23);
            buttonGuardar.TabIndex = 12;
            buttonGuardar.Text = "Guardar";
            buttonGuardar.UseVisualStyleBackColor = false;
            buttonGuardar.Click += buttonGuardar_Click;
            // 
            // buttonCancelar
            // 
            buttonCancelar.BackColor = Color.BlueViolet;
            buttonCancelar.FlatStyle = FlatStyle.Flat;
            buttonCancelar.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            buttonCancelar.ForeColor = SystemColors.ButtonFace;
            buttonCancelar.Location = new Point(70, 264);
            buttonCancelar.Name = "buttonCancelar";
            buttonCancelar.Size = new Size(75, 23);
            buttonCancelar.TabIndex = 13;
            buttonCancelar.Text = "Cancelar";
            buttonCancelar.UseVisualStyleBackColor = false;
            buttonCancelar.Click += buttonCancelar_Click;
            // 
            // FormColumnas
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(303, 305);
            Controls.Add(buttonCancelar);
            Controls.Add(buttonGuardar);
            Controls.Add(numericUpDownTotal);
            Controls.Add(textBoxTotal);
            Controls.Add(numericUpDownIVA);
            Controls.Add(textBoxIVA);
            Controls.Add(numericUpDownNumeroComprobante);
            Controls.Add(textBoxNumeroComprobante);
            Controls.Add(numericUpDownPuntoVenta);
            Controls.Add(textBoxPuntoVenta);
            Controls.Add(textBoxCUIT);
            Controls.Add(textBoxNumeroColumna);
            Controls.Add(textBoxColumna);
            Controls.Add(numericUpDownCuit);
            Name = "FormColumnas";
            Text = "FormColumnas";
            Load += FormColumnas_Load;
            ((System.ComponentModel.ISupportInitialize)numericUpDownCuit).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownPuntoVenta).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownNumeroComprobante).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownIVA).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownTotal).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private NumericUpDown numericUpDownCuit;
        private TextBox textBoxColumna;
        private TextBox textBoxNumeroColumna;
        private TextBox textBoxCUIT;
        private TextBox textBoxPuntoVenta;
        private NumericUpDown numericUpDownPuntoVenta;
        private TextBox textBoxNumeroComprobante;
        private NumericUpDown numericUpDownNumeroComprobante;
        private TextBox textBoxIVA;
        private NumericUpDown numericUpDownIVA;
        private TextBox textBoxTotal;
        private NumericUpDown numericUpDownTotal;
        private Button buttonGuardar;
        private Button buttonCancelar;
    }
}