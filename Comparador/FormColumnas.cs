﻿using System;
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

        public string NombreEsquema { get; set; }

        public FormColumnas()
        {
            InitializeComponent();
        }

        private void buttonGuardar_Click(object sender, EventArgs e)
        {
            if (this.NombreEsquema == null)
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
                        (int)numericUpDownTotal.Value,
                        (int)numericUpDownFecha.Value
                    };

                        // Crear un nuevo objeto Esquema con el nombre y las columnas seleccionadas
                        EsquemaColumnas esquema = new EsquemaColumnas(nombreEsquema, (int)numericUpDownCuit.Value, (int)numericUpDownPuntoVenta.Value,
                            (int)numericUpDownNumeroComprobante.Value, (int)numericUpDownIVA.Value, (int)numericUpDownTotal.Value, (int)numericUpDownFecha.Value);

                        // Ruta del archivo Esquemas en el directorio de la aplicación
                        string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

                        try
                        {
                            // Leer los esquemas existentes
                            List<EsquemaColumnas> esquemasExistentes = LeerEsquemasDesdeArchivo(filePath);

                            // Verificar si ya existe un esquema con el mismo nombre
                            var esquemaExistente = esquemasExistentes.Find(e => e.Nombre == nombreEsquema);
                            if (esquemaExistente != null)
                            {
                                // Preguntar al usuario si desea sobrescribir el esquema existente
                                DialogResult result = MessageBox.Show($"Ya existe un esquema con el nombre '{nombreEsquema}'. ¿Desea reemplazarlo?",
                                    "Confirmar sobrescritura", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (result == DialogResult.No)
                                {
                                    MessageBox.Show("El esquema no ha sido guardado.");
                                    return;
                                }
                                else
                                {
                                    // Eliminar el esquema existente de la lista
                                    esquemasExistentes.Remove(esquemaExistente);
                                }
                            }

                            // Agregar el nuevo esquema a la lista
                            esquemasExistentes.Add(esquema);

                            // Serializar la lista de esquemas a una cadena de texto en formato JSON
                            List<string> esquemasJson = new List<string>();
                            foreach (var esq in esquemasExistentes)
                            {
                                esquemasJson.Add(Newtonsoft.Json.JsonConvert.SerializeObject(esq));
                            }

                            // Escribir la cadena de texto en el archivo
                            File.WriteAllLines(filePath, esquemasJson);

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
            else
            {
                // Ruta del archivo Esquemas en el directorio de la aplicación
                string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

                List<EsquemaColumnas> esquemasExistentes = LeerEsquemasDesdeArchivo(filePath);

                // Verificar si ya existe un esquema con el mismo nombre
                var esquemaExistente = esquemasExistentes.Find(e => e.Nombre == NombreEsquema);
                if (esquemaExistente != null)
                {
                    // Crear un nuevo objeto Esquema con el nombre y las columnas seleccionadas
                    EsquemaColumnas esquema = new EsquemaColumnas(NombreEsquema, (int)numericUpDownCuit.Value, (int)numericUpDownPuntoVenta.Value,
                        (int)numericUpDownNumeroComprobante.Value, (int)numericUpDownIVA.Value, (int)numericUpDownTotal.Value, (int)numericUpDownFecha.Value);

                    // Eliminar el esquema existente de la lista
                    esquemasExistentes.Remove(esquemaExistente);

                    // Agregar el nuevo esquema a la lista
                    esquemasExistentes.Add(esquema);

                    // Serializar la lista de esquemas a una cadena de texto en formato JSON
                    List<string> esquemasJson = new List<string>();
                    foreach (var esq in esquemasExistentes)
                    {
                        esquemasJson.Add(Newtonsoft.Json.JsonConvert.SerializeObject(esq));
                    }

                    // Escribir la cadena de texto en el archivo
                    File.WriteAllLines(filePath, esquemasJson);

                    MessageBox.Show("Esquema modificado correctamente");
                }

                this.Close();
            }
        }

        List<EsquemaColumnas> LeerEsquemasDesdeArchivo(string filePath)
        {
            List<EsquemaColumnas> esquemas = new List<EsquemaColumnas>();

            if (File.Exists(filePath))
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    var esquema = Newtonsoft.Json.JsonConvert.DeserializeObject<EsquemaColumnas>(line);
                    esquemas.Add(esquema);
                }
            }

            return esquemas;
        }


        private void buttonCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void cargarDatos(int IndiceColumnaCuit, int IndiceColumnaIVA, int IndiceColumnaTotal, int IndiceColumnaPuntoVenta, int IndiceColumnaComprobante, string nombre, int IndiceColumnaFecha)
        {
            numericUpDownCuit.Value = IndiceColumnaCuit;
            numericUpDownIVA.Value = IndiceColumnaIVA;
            numericUpDownTotal.Value = IndiceColumnaTotal;
            numericUpDownPuntoVenta.Value = IndiceColumnaPuntoVenta;
            numericUpDownNumeroComprobante.Value = IndiceColumnaComprobante;
            this.NombreEsquema = nombre;
            numericUpDownFecha.Value = IndiceColumnaFecha;
        }

        private void FormColumnas_Load(object sender, EventArgs e)
        {

        }

        private void numericUpDownPuntoVenta_ValueChanged(object sender, EventArgs e)
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
        public int IndiceColumnaFecha { get; set; }

        public EsquemaColumnas() { }

        public EsquemaColumnas(string nombre, int indiceCuit, int indicePuntoVenta, int indiceNumeroComprobante, int indiceIVA, int indiceTotal, int indiceFecha)
        {
            Nombre = nombre;
            IndiceColumnaCuit = indiceCuit;
            IndiceColumnaPuntoVenta = indicePuntoVenta;
            IndiceColumnaComprobante = indiceNumeroComprobante;
            IndiceColumnaIVA = indiceIVA;
            IndiceColumnaTotal = indiceTotal;
            IndiceColumnaFecha = indiceFecha;
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