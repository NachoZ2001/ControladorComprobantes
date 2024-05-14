using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Comparador
{
    public partial class FormColumnas : Form
    {
        public List<int> ColumnasSeleccionadas { get; private set; }

        public FormColumnas()
        {
            InitializeComponent();
        }

        private void buttonGuardar_Click(object sender, EventArgs e)
        {
            // Mostrar el cuadro de diálogo de entrada
            InputDialog inputDialog = new InputDialog();
            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                string nombreEsquema = inputDialog.EnteredText;

                if (!string.IsNullOrEmpty(nombreEsquema))
                {
                    ColumnasSeleccionadas = new List<int>()
                    {
                        (int)numericUpDownCuit.Value,
                        (int)numericUpDownPuntoVenta.Value,
                        (int)numericUpDownNumeroComprobante.Value,
                        (int)numericUpDownIVA.Value,
                        (int)numericUpDownTotal.Value
                    };

                    // Crear un nuevo objeto Esquema con el nombre y las columnas seleccionadas
                    EsquemaColumnas esquema = new EsquemaColumnas(nombreEsquema, (int)numericUpDownCuit.Value, (int)numericUpDownPuntoVenta.Value,
                        (int)numericUpDownNumeroComprobante.Value, (int)numericUpDownIVA.Value, (int)numericUpDownTotal.Value);

                    // Resto del proceso para guardar el esquema...
                    string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

                    try
                    {
                        // Serializar el objeto Esquema a una cadena de texto en formato JSON
                        string esquemaJson = Newtonsoft.Json.JsonConvert.SerializeObject(esquema);

                        // Escribir la cadena de texto en el archivo
                        using (StreamWriter writer = new StreamWriter(filePath, true))
                        {
                            // Escribir una nueva línea antes de agregar el nuevo esquema, si el archivo ya contiene datos
                            if (writer.BaseStream.Length > 0)
                                writer.WriteLine();

                            writer.Write(esquemaJson);
                        }

                        MessageBox.Show("Esquema guardado en el archivo Esquemas.txt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al guardar el esquema: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Por favor, ingrese un nombre para el esquema.");
                }
            }

            this.Close();
        }
        private void buttonCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormColumnas_Load(object sender, EventArgs e)
        {

        }
    }

    public class EsquemaColumnas
    {
        public string Nombre { get; set; }
        public int IndiceColumnaCuit { get; set; }
        public int IndiceColumnaPuntoVenta { get; set; }
        public int IndiceColumnaComprobante { get; set; }
        public int IndiceColumnaIVA { get; set; }
        public int IndiceColumnaTotal { get; set; }

        public EsquemaColumnas() { }

        public EsquemaColumnas(string nombre, int indiceCuit, int indicePuntoVenta, int indiceNumeroComprobante, int indiceIVA, int indiceTotal)
        {
            Nombre = nombre;
            IndiceColumnaCuit = indiceCuit;
            IndiceColumnaPuntoVenta = indicePuntoVenta;
            IndiceColumnaComprobante = indiceNumeroComprobante;
            IndiceColumnaIVA = indiceIVA;
            IndiceColumnaTotal = indiceTotal;
        }
    }

    public class InputDialog : Form
    {
        private TextBox textBox;
        private Button okButton;

        public InputDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.textBox = new TextBox();
            this.okButton = new Button();

            this.SuspendLayout();

            this.textBox.Location = new Point(20, 20);
            this.textBox.Size = new Size(200, 20);
            this.textBox.BackColor = System.Drawing.Color.BlueViolet;
            this.textBox.ForeColor = System.Drawing.Color.White;
            this.textBox.BorderStyle = BorderStyle.FixedSingle;
            this.textBox.TextAlign = HorizontalAlignment.Center;

            this.okButton.Text = "OK";
            this.okButton.FlatStyle = FlatStyle.Popup;
            this.okButton.Location = new Point(20, 50);
            this.okButton.BackColor = System.Drawing.Color.BlueViolet;
            this.okButton.ForeColor = System.Drawing.Color.White;
            this.okButton.TextAlign = ContentAlignment.MiddleCenter;
            this.okButton.Click += OkButton_Click;

            this.Controls.Add(this.textBox);
            this.Controls.Add(this.okButton);

            this.ClientSize = new Size(240, 100);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Ingrese un nombre";
            this.BackColor = System.Drawing.Color.Purple; // Establece el color de fondo del formulario en Purple

            this.ResumeLayout(false);
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        public string EnteredText
        {
            get { return textBox.Text; }
        }
    }
}
