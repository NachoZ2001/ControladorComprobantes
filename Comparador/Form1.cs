using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Comparador
{
    public partial class Form1 : Form
    {

        public List<int> columnas { get; set; }

        public int UltimaColumnaAFIP { get; set; }

        public int UltimaColumnaHolistor { get; set; }

        public int UltimaColumnaContabilidad { get; set; }

        public Form1()
        {
            InitializeComponent();

            columnas = new List<int>();

            UltimaColumnaAFIP = new int();

            UltimaColumnaHolistor = new int();

            // Establecer el estilo del borde y deshabilitar el cambio de tamaño
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Establecer el tamaño mínimo y máximo para evitar el cambio de tamaño
            this.MinimumSize = this.MaximumSize = this.Size;

            buttonEditarEsquema.Visible = false;

            InicializarYMostrarEsquemas();
        }

        private void InicializarYMostrarEsquemas()
        {
            // Borra los elementos existentes en el ComboBox
            comboBoxEsquemas.Items.Clear();

            // Define una lista de esquemas
            List<Esquema> listaEsquemas = new List<Esquema>();

            // Ruta del archivo Esquemas en el directorio de la aplicación
            string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

            // Cargar los esquemas desde el archivo
            CargarEsquemasDesdeArchivo(filePath, listaEsquemas);

            // Agregar los nombres de los esquemas al ComboBox
            foreach (Esquema esquema in listaEsquemas)
            {
                comboBoxEsquemas.Items.Add(esquema.Nombre);
            }

            // Mostrar el primer esquema en el ComboBox si hay al menos uno
            if (comboBoxEsquemas.Items.Count > 0)
            {
                comboBoxEsquemas.SelectedIndex = 0;
            }

            buttonEditarEsquema.Visible = true;
        }

        private void CargarEsquemasDesdeArchivo(string filePath, List<Esquema> listaEsquemas)
        {
            try
            {
                // Leer todas las líneas del archivo
                string[] lines = File.ReadAllLines(filePath);

                foreach (string line in lines)
                {
                    // Ignorar las líneas en blanco o nulas
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    try
                    {
                        // Deserializar cada línea del archivo en un objeto Esquema
                        Esquema esquema = Newtonsoft.Json.JsonConvert.DeserializeObject<Esquema>(line);

                        // Agregar el esquema a la lista
                        listaEsquemas.Add(esquema);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al deserializar la línea '{line}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los esquemas: " + ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pictureBoxRuedaCargando.Visible = false;
        }

        private void buttonAfip_Click(object sender, EventArgs e)
        {
            SeleccionarArchivo(textBoxAfip);
        }

        private void buttonHolistor_Click(object sender, EventArgs e)
        {
            SeleccionarArchivo(textBoxHolistor);
        }

        private void SeleccionarArchivo(TextBox textBox)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls";
                openFileDialog.Title = "Seleccionar el archivo Excel";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Muestra la ruta seleccionada en el TextBox correspondiente
                    textBox.Text = openFileDialog.FileName;
                }
            }
        }

        private async void buttonProcesar_Click(object sender, EventArgs e)
        {
            if (textBoxAfip.Text == "Archivo AFIP")
            {
                MessageBox.Show("Falta la ruta del archivo AFIP", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (textBoxHolistor.Text == "Archivo Contabilidad")
            {
                MessageBox.Show("Falta la ruta del archivo HOLISTOR", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (comboBoxEsquemas.Text == null)
            {
                MessageBox.Show("Falta seleccionar un esquema", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            pictureBoxRuedaCargando.Visible = true;

            // Define una lista de esquemas
            List<Esquema> listaEsquemas = new List<Esquema>();

            // Ruta del archivo Esquemas en el directorio de la aplicación
            string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

            // Cargar los esquemas desde el archivo
            CargarEsquemasDesdeArchivo(filePath, listaEsquemas);

            if (columnas != null)
            {
                this.columnas.Clear();
            }

            if (comboBoxEsquemas.SelectedItem != null)
            {
                foreach (Esquema esquema in listaEsquemas)
                {
                    if (comboBoxEsquemas.SelectedItem.ToString() == esquema.Nombre)
                    {
                        this.columnas.Add(esquema.IndiceCuit);
                        this.columnas.Add(esquema.IndicePuntoVenta);
                        this.columnas.Add(esquema.IndiceNumeroComprobante);
                        this.columnas.Add(esquema.IndiceIVA);
                        this.columnas.Add(esquema.IndiceTotal);
                        this.columnas.Add(esquema.IndiceFecha);
                    }
                }
            }

            if (comboBoxEsquemas.SelectedItem.ToString() == "Holistor")
            {
                // Realizar el proceso de manera asíncrona de comparar con Holistor
                await Task.Run(() => RealizarComparacionHolistor(textBoxAfip.Text, textBoxHolistor.Text));
            }
            else
            {
                // Realizar el proceso de manera asíncrona para comparar
                await Task.Run(() => RealizarComparacion(textBoxAfip.Text, textBoxHolistor.Text));
            }

            pictureBoxRuedaCargando.Visible = false;

            // Muestra un mensaje de éxito
            MessageBox.Show("Proceso completado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static int ExtraerTipoFacturaAFIP(string tipo)
        {
            var diccionarioTiposFacturas = new Dictionary<string, int>
            {
                {"Factura A", 1},
                {"FAC A", 1},
                {"FA A", 1},
                {"Factura B", 2},
                {"FAC B", 2},
                {"FA B", 2},
                {"Factura C", 3},
                {"FAC C", 3},
                {"FA C", 3},
                {"Nota de Crédito A", 4},
                {"NC A", 4},
                {"Nota de Débito A", 5},
                {"ND A", 5},
                {"Recibo A", 1},
                {"Recibo B", 2},
                {"Recibo C", 3}
            };

            // Normalizar cadenas y extraer tipo y letra para comparación con diccionarioAFIP
            var normalizadoAFIP = NormalizarAFIP(tipo);

            return ObtenerValorDeDiccionario(diccionarioTiposFacturas, normalizadoAFIP);
        }

        static int ExtraerTipoFacturaHolistor(string tipo)
        {
            var diccionarioTiposFacturas = new Dictionary<string, int>
            {
                {"Factura A", 1},
                {"FAC A", 1},
                {"FA A", 1},
                {"Factura B", 2},
                {"FAC B", 2},
                {"FA B", 2},
                {"Factura C", 3},
                {"FAC C", 3},
                {"FA C", 3},
                {"Nota de Crédito A", 4},
                {"NC A", 4},
                {"Nota de Débito A", 5},
                {"ND A", 5},
                {"Recibo A", 1},
                {"Recibo B", 2},
                {"Recibo C", 3}
            };

            // Normalizar cadenas y extraer tipo y letra para comparación con diccionarioAFIP
            var normalizadoHolistor = NormalizarHolistor(tipo);

            return ObtenerValorDeDiccionario(diccionarioTiposFacturas, normalizadoHolistor);
        }

        // Función para normalizar cadena de AFIP (extraer tipo y letra después del identificador numérico)
        static string NormalizarAFIP(string input)
        {
            int separatorIndex = input.IndexOf('-');
            if (separatorIndex >= 0)
            {
                return input.Substring(separatorIndex + 1).Trim();
            }
            return input.Trim();
        }

        // Función para normalizar cadena de Excel (extraer tipo y letra al principio)
        static string NormalizarHolistor(string input)
        {
            string[] parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                return parts[0] + " " + parts[1];
            }
            return input.Trim();
        }

        // Función para buscar valor en el diccionario
        static int ObtenerValorDeDiccionario(Dictionary<string, int> dictionary, string key)
        {
            if (dictionary.TryGetValue(key, out int value))
            {
                return value;
            }
            return -1; // Valor por defecto o manejo de error según necesidad
        }

        private void RealizarComparacion(string rutaExcelAfip, string rutaExcelContabilidad)
        {
            int indiceColumnaCUIT = columnas[0];
            int indiceColumnaPuntoVenta = columnas[1];
            int indiceColumnaNumeroComprobante = columnas[2];
            int indiceColumnaIVA = columnas[3];
            int indiceColumnaTotal = columnas[4];
            int indiceColumnaFecha = columnas[5];

            //Armar diccionarios
            var diccionarioContabilidad = ArmarDiccionarioContabilidad(rutaExcelContabilidad, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha);
            var diccionarioAFIP = ArmarDiccionarioAFIP(rutaExcelAfip);

            //Comparar y marcar filas 
            CompararYMarcarFilasContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAfip, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha,(double)numericUpDownTolerancia.Value);

            //Marcar y señalizar en AFIP porque no coincidieron en la comparacion
            MarcarNoSeñalizadosEnRojo(diccionarioContabilidad, diccionarioAFIP, rutaExcelAfip);
        }

        private void RealizarComparacionHolistor(string rutaExcelAfip, string rutaExcelHolistor)
        {
            //Armar ambos diccionarios
            var diccionarioHolistor = ArmarDiccionarioHolistor(rutaExcelHolistor);
            var diccionarioAFIP = ArmarDiccionarioAFIP(rutaExcelAfip);

            //Comparar y marcar filas en base al Excel Holistor
            CompararYMarcarFilasHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAfip, (double)numericUpDownTolerancia.Value);

            //Marcar y señalizar en AFIP porque no coincidieron en la comparacion
            MarcarNoSeñalizadosEnRojo(diccionarioHolistor, diccionarioAFIP, rutaExcelAfip);
        }

        // Función para obtener el índice de una columna específica
        static int ObtenerIndiceColumna(IXLWorksheet worksheet, string nombreColumna)
        {
            int indiceColumna = -1;

            for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
            {
                string valor = worksheet.Cell(1, col).GetString();

                if (valor.Equals(nombreColumna, StringComparison.OrdinalIgnoreCase))
                {
                    indiceColumna = col;
                    break;
                }
            }

            return indiceColumna;
        }

        // Función para procesar correctamente los números de comprobante
        static string ProcesarNumeros(string input)
        {
            // Eliminar caracteres que no sean números
            string numeros = Regex.Replace(input, @"\D", "");

            // Insertar guion en la posición 5, para separar punto de venta de comprobante
            numeros = numeros.Insert(4, "-");

            // Separar punto de venta y número de comprobante con un guion
            string[] partes = numeros.Split('-');
            string puntoDeVenta = partes[0];
            string numeroComprobante = partes[1];

            // Eliminar los ceros no significativos antes del primer número distinto de 0
            puntoDeVenta = EliminarCerosNoSignificativos(puntoDeVenta);
            numeroComprobante = EliminarCerosNoSignificativos(numeroComprobante);

            // Unir punto de venta y número de comprobante sin el guion
            return puntoDeVenta + numeroComprobante;
        }

        static string EliminarCerosNoSignificativos(string input)
        {
            // Encuentra el índice del primer dígito distinto de cero
            int indice = 0;
            while (indice < input.Length && input[indice] == '0')
            {
                indice++;
            }

            // Elimina los ceros no significativos antes del primer dígito distinto de cero
            return input.Substring(indice);
        }

        // Función para limpiar los caracteres que no sean números de los CUIT
        static string LimpiarCUIT(string cuit)
        {
            return Regex.Replace(cuit, @"[^\d]", "");
        }

        // Función para armar el diccionario de Contabilidad --> {CUIT}: (fila, iva, total, comprobante, fecha)
        static Dictionary<string, List<(int, double, double, string, DateTime, int)>> ArmarDiccionarioContabilidad(string rutaExcel, int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string, DateTime, int)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja

                //Arma cuando tengo punto de venta y comprobante en una sola columna
                if (indiceColumnaPuntoVenta == -1 && indiceColumnaNumeroComprobante != -1)
                {
                    int ultimaFila = worksheet.LastRowUsed().RowNumber();

                    for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                    {
                        string numeroComprobante = worksheet.Cell(fila, indiceColumnaNumeroComprobante).GetString();

                        // Procesar el número de comprobante
                        numeroComprobante = ProcesarNumeros(numeroComprobante);

                        // Obtener valor de la columna IVA
                        string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                        string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                        double iva = double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna Total
                        string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaTotal).GetString();
                        string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                        double total = double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna CUIT
                        string cuit = worksheet.Cell(fila, indiceColumnaCUIT).GetString();
                        cuit = LimpiarCUIT(cuit);

                        // Obtener valor de la columna Fecha
                        string stringFechaArchivo1 = worksheet.Cell(fila, indiceColumnaFecha).GetString();
                        DateTime fechaArchivoContabilidad = DateTime.Parse(stringFechaArchivo1);

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string, DateTime, int)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, numeroComprobante, fechaArchivoContabilidad, 1));
                    }
                }
                //Arma cuando tenga punto de venta y comprobante separados en distintas columnas
                else
                {
                    int ultimaFila = worksheet.LastRowUsed().RowNumber();

                    for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                    {
                        string puntoVenta = worksheet.Cell(fila, indiceColumnaPuntoVenta).Value.ToString();
                        string numeroComprobante = worksheet.Cell(fila, indiceColumnaNumeroComprobante).Value.ToString();
                        string comprobanteCompleto = puntoVenta + numeroComprobante;

                        // Obtener valor de la columna IVA
                        string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                        string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                        double iva = 0;
                        if (valorCeldaIVASinComa != "")
                        {
                            iva = double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture);
                        }

                        // Obtener valor de la columna Total
                        string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaTotal).GetString();
                        string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                        double total = double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna CUIT
                        string cuit = worksheet.Cell(fila, indiceColumnaCUIT).GetString();
                        cuit = LimpiarCUIT(cuit);

                        // Obtener valor de la columna Fecha
                        string stringFechaArchivo1 = worksheet.Cell(fila, indiceColumnaFecha).GetString();
                        DateTime fechaArchivoContabilidad = DateTime.Parse(stringFechaArchivo1);

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string, DateTime, int)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, comprobanteCompleto, fechaArchivoContabilidad, 1));
                    }
                }
            }

            return diccionario;
        }

        // Función para armar el diccionario de Holistor --> {CUIT}: (fila, iva, total, comprobante, fecha, tipo comprobante)
        static Dictionary<string, List<(int, double, double, string, DateTime, int)>> ArmarDiccionarioHolistor(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string, DateTime, int)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja
                int indiceColumnaComprobante = ObtenerIndiceColumna(worksheet, "Comprobante");

                if (indiceColumnaComprobante != -1)
                {
                    int ultimaFila = worksheet.LastRowUsed().RowNumber();

                    for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                    {
                        string numeroComprobante = worksheet.Cell(fila, indiceColumnaComprobante).GetString();
                        string Comprobante = numeroComprobante;

                        // Procesar el número de comprobante
                        numeroComprobante = ProcesarNumeros(numeroComprobante);

                        // Obtener valor de la columna IVA
                        int indiceColumnaIVA = ObtenerIndiceColumna(worksheet, "IVA");
                        string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                        string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                        double iva = double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna Total
                        int indiceColumnaTotal = ObtenerIndiceColumna(worksheet, "Total");
                        string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaTotal).GetString();
                        string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                        double total = double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna CUIT
                        int indiceColumnaCuit = ObtenerIndiceColumna(worksheet, "Tipo/Nro.Doc.");
                        string cuit = worksheet.Cell(fila, indiceColumnaCuit).GetString();
                        cuit = LimpiarCUIT(cuit);

                        // Obtener valor de la columna Fecha
                        int indiceFechaArchivoHolistor = ObtenerIndiceColumna(worksheet, "Fecha");
                        string stringFechaArchivo1 = worksheet.Cell(fila, indiceFechaArchivoHolistor).GetString();
                        DateTime fechaArchivoHolistor = DateTime.Parse(stringFechaArchivo1);

                        // Obtener el valor del tipo de comprobante mapeado
                        int tipoComprobante = ExtraerTipoFacturaHolistor(Comprobante);

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string,DateTime, int)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, numeroComprobante, fechaArchivoHolistor, tipoComprobante));
                    }
                }
                else
                {
                    Console.WriteLine("La columna 'Comprobante' no se encontró en el Excel.");
                }
            }

            return diccionario;
        }

        // Función para armar el diccionario de AFIP --> {CUIT}: (fila, iva, total, comprobante, fecha, tipo comprobante)
        static Dictionary<string, List<(int, double, double, string, DateTime, int)>> ArmarDiccionarioAFIP(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string, DateTime, int)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja
                int indiceColumnaPuntoVenta = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobante = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaIVA = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotal = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaCuit = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");
                int indiceColumnaFecha = ObtenerIndiceColumna(worksheet, "Fecha");
                int indiceColumnaTipo = ObtenerIndiceColumna(worksheet, "Tipo");

                if (indiceColumnaPuntoVenta != -1 && indiceColumnaComprobante != -1 && indiceColumnaIVA != -1 && indiceColumnaTotal != -1 && indiceColumnaCuit != -1 && indiceColumnaFecha != -1)
                {
                    int ultimaFila = worksheet.LastRowUsed().RowNumber();

                    for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                    {
                        string puntoVenta = worksheet.Cell(fila, indiceColumnaPuntoVenta).Value.ToString();
                        string numeroComprobante = worksheet.Cell(fila, indiceColumnaComprobante).Value.ToString();
                        string comprobanteCompleto = puntoVenta + numeroComprobante;

                        // Obtener valor de la columna IVA
                        string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                        string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                        double iva = 0;
                        if (valorCeldaIVASinComa != "")
                        {
                            iva = double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture);
                        }

                        // Obtener valor de la columna Total
                        string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaTotal).GetString();
                        string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                        double total = double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture);

                        // Obtener valor de la columna CUIT
                        string cuit = worksheet.Cell(fila, indiceColumnaCuit).GetString();

                        // Obtener el valor de la columna Fecha
                        string stringFechaArchivo1 = worksheet.Cell(fila, indiceColumnaFecha).GetString();
                        DateTime fechaArchivo1 = DateTime.Parse(stringFechaArchivo1);

                        // Obtener el valor mapeado del tipo de comprobante                      
                        int tipoMapeado = ExtraerTipoFacturaAFIP(worksheet.Cell(fila, indiceColumnaTipo).GetString());
                        
                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string, DateTime, int)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, comprobanteCompleto, fechaArchivo1, tipoMapeado));
                    }
                }
                else
                {
                    Console.WriteLine("Alguna de las columnas necesarias no se encontró en el Excel.");
                }
            }

            return diccionario;
        }

        //Comparacion para los archivos de contabilidad que son de Holistor
        private async void CompararYMarcarFilasHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            // Primera comparacion
            CompararFechaImporteComprobanteHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAFIP, tolerancia);

            // Segunda comparacion
            CompararImporteComprobanteHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAFIP, tolerancia);

            // Tercera comparacion
            CompararComprobanteHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAFIP, tolerancia);

            // Cuarta comparacion
            CompararImportesHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAFIP, tolerancia);

            // Marcar en rojo los que no entraron en ningun filtro
            MarcarNoSeñalizadosEnRojoHolistor(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor);
        }

        private async void CompararFechaImporteComprobanteHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoHolistor = workbookHolistor.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                int ultimaColumnaHolistor = worksheetArchivoHolistor.LastColumnUsed().ColumnNumber();
                UltimaColumnaHolistor = ultimaColumnaHolistor;

                int ultimaColumnaAFIP = worksheetArchivoAFIP.LastColumnUsed().ColumnNumber();
                UltimaColumnaAFIP = ultimaColumnaAFIP;

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");
                int indiceColumnaTipoComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo");

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Tipo/Nro.Doc.");
                int indiceColumnaFechaHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Fecha");

                worksheetArchivoHolistor.Cell(1, ultimaColumnaHolistor + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpHolistor in diccionarioHolistor)
                {
                    string claveHolistor = kvpHolistor.Key;
                    var registrosHolistor = kvpHolistor.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveHolistor))
                    {
                        var registrosAFIP = diccionarioAFIP[claveHolistor];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosHolistor = registrosHolistor.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroHolistor in registrosHolistor)
                        {
                            // Señalizar en verde CUIT
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                
                                // Comparamos por comprobante, importes, fecha y tipo de comprobante
                                if ((Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia) && (Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia) && registroHolistor.Item4 == registroAFIP.Item4 && registroHolistor.Item5 == registroAFIP.Item5 && registroHolistor.Item6 == registroAFIP.Item6)
                                {
                                    // Coincide

                                    // Señalizar en verde ambos comprobantes
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Señalizo en verde el IVA
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde el TOTAL
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde la FECHA
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde el TIPO de COMPROBANTE
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    // El tipo de comprobante de Holistor ya lo señalizo arriba cuando marco los comprobantes
                                }

                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroHolistor in registrosHolistor)
                        {
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookHolistor.SaveAs(rutaExcelHolistor);
            }
        }

        private async void CompararImporteComprobanteHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoHolistor = workbookHolistor.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                int ultimaColumnaHolistor = UltimaColumnaHolistor;
                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");
                int indiceColumnaTipoComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo");

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Tipo/Nro.Doc.");
                int indiceColumnaFechaHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Fecha");

                worksheetArchivoHolistor.Cell(1, ultimaColumnaHolistor + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpHolistor in diccionarioHolistor)
                {
                    string claveHolistor = kvpHolistor.Key;
                    var registrosHolistor = kvpHolistor.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveHolistor))
                    {
                        var registrosAFIP = diccionarioAFIP[claveHolistor];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosHolistor = registrosHolistor.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroHolistor in registrosHolistor)
                        {
                            // Señalizar en verde CUIT
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante, tipo de comprobante e importes, ignorando los ya señalizados
                                if ((Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia) && (Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia) && registroHolistor.Item4 == registroAFIP.Item4 && registroHolistor.Item6 == registroAFIP.Item6 &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)  
                                     && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Coincide

                                    // Señalizar en verde ambos comprobantes (en esta se señaliza tambien el tipo de comprobante de HOLISTOR)
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Señalizo en verde el IVA
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde el TOTAL
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde el TIPO de COMPROBANTE de AFIP
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Comparar la FECHA
                                    if (registroHolistor.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroHolistor in registrosHolistor)
                        {
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookHolistor.SaveAs(rutaExcelHolistor);
            }
        }

        private async void CompararComprobanteHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoHolistor = workbookHolistor.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                int ultimaColumnaHolistor = UltimaColumnaHolistor;
                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");
                int indiceColumnaTipoComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo");

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Tipo/Nro.Doc.");
                int indiceColumnaFechaHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Fecha");

                worksheetArchivoHolistor.Cell(1, ultimaColumnaHolistor + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpHolistor in diccionarioHolistor)
                {
                    string claveHolistor = kvpHolistor.Key;
                    var registrosHolistor = kvpHolistor.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveHolistor))
                    {
                        var registrosAFIP = diccionarioAFIP[claveHolistor];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosHolistor = registrosHolistor.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroHolistor in registrosHolistor)
                        {
                            // Señalizar en verde CUIT
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante y tipo del mismo, ignorando los ya señalizados
                                if (registroHolistor.Item4 == registroAFIP.Item4 && registroHolistor.Item6 == registroAFIP.Item6 &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                     && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Coincide

                                    // Señalizar en verde ambos comprobantes (en esta señalizo tambien el tipo de comprobante de HOLISTOR)
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Señalizo el TIPO de COMPROBANTE de AFIP
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Comparar el IVA
                                    if ((Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia))
                                    {
                                        // Señalizo en verde el IVA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el IVA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += "IVA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "IVA esta mal";
                                    }

                                    // Comparar el TOTAL
                                    if ((Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia))
                                    {
                                        // Señalizo en verde el TOTAL
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el TOTAL
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += "TOTAL esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "TOTAL esta mal";
                                    }

                                    // Comparar la FECHA
                                    if (registroHolistor.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroHolistor in registrosHolistor)
                        {
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookHolistor.SaveAs(rutaExcelHolistor);
            }
        }

        private async void CompararImportesHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoHolistor = workbookHolistor.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                int ultimaColumnaHolistor = UltimaColumnaHolistor;
                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");
                int indiceColumnaTipoComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo");

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Tipo/Nro.Doc.");
                int indiceColumnaFechaHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Fecha");

                worksheetArchivoHolistor.Cell(1, ultimaColumnaHolistor + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpHolistor in diccionarioHolistor)
                {
                    string claveHolistor = kvpHolistor.Key;
                    var registrosHolistor = kvpHolistor.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveHolistor))
                    {
                        var registrosAFIP = diccionarioAFIP[claveHolistor];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosHolistor = registrosHolistor.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroHolistor in registrosHolistor)
                        {
                            // Señalizar en verde CUIT
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por importes, ignorando los ya señalizados
                                if ((Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia) && (Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia) &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                     && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Coinciden los importes
                                    // 
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Señalizo en verde el IVA
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
  

                                    // Señalizo en verde el TOTAL
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Comparar la FECHA
                                    if (registroHolistor.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaFechaHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }

                                    // Comparar el COMPROBANTE
                                    if (registroHolistor.Item4 == registroAFIP.Item4)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el COMPROBANTE
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " COMPROBANTE esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " COMPROBANTE esta mal";
                                    }

                                    // Comparar el TIPO de COMPROBANTE
                                    if (registroHolistor.Item6 == registroAFIP.Item6)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el TIPO de COMPROBANTE
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " TIPO COMPROBANTE esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " TIPO COMPROBANTE esta mal";
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroHolistor in registrosHolistor)
                        {
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookHolistor.SaveAs(rutaExcelHolistor);
            }
        }

        private async void MarcarNoSeñalizadosEnRojoHolistor(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelHolistor)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            {
                var worksheet = workbookHolistor.Worksheets.First();
                var defaultColor = XLColor.FromIndex(0); // Color predeterminado de Excel

                int ultimaColumnaHolistor = worksheet.LastColumnUsed().ColumnNumber();

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheet, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheet, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheet, "Tipo/Nro.Doc.");
                int indiceColumnaFechaHolistor = ObtenerIndiceColumna(worksheet, "Fecha");

                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        string cuit = worksheet.Cell(row.RowNumber(), indiceColumnaCuitHolistor).GetString();
                        if (!diccionarioAFIP.ContainsKey(cuit))
                        {
                            worksheet.Cell(row.RowNumber(), indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheet.Cell(row.RowNumber(), UltimaColumnaHolistor + 1).Value = "Este cuit no tiene ningun registro en HOLISTOR";
                        }
                    }
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitHolistor).Style.Fill.BackgroundColor == XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        worksheet.Row(row.RowNumber()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                        worksheet.Cell(row.RowNumber(), UltimaColumnaHolistor + 1).Value = "NO coincide ningun registro";
                    }
                }

                workbookHolistor.Save();
            }
        }

        //Comparacion para los archivos de contabilidad que no son de Holistor
        private async void CompararYMarcarFilasContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha, double tolerancia)
        {
            // Primera comparacion
            CompararFechaImportesComprobanteContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAFIP, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha, tolerancia);

            // Segunda comparacion
            CompararFechaComprobanteContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAFIP, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha, tolerancia);

            // Tercera comparacion
            CompararComprobanteContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAFIP, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha, tolerancia);

            // Cuarta comparacion
            CompararImportesContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAFIP, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, indiceColumnaFecha, tolerancia);

            // Marcar los no señalizados en rojo
            MarcarNoSeñalizadosEnRojoContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaIVA, indiceColumnaTotal, indiceColumnaCUIT, indiceColumnaFecha);
        }

        private async void CompararFechaImportesComprobanteContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha, double tolerancia)
        {
            using (var workbookContabilidad = new XLWorkbook(rutaExcelContabilidad))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoContabilidad = workbookContabilidad.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                //Obtengo indices de las columnas del Excel de AFIP
                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");

                int ultimaColumnaContabilidad = worksheetArchivoContabilidad.LastColumnUsed().ColumnNumber();
                UltimaColumnaContabilidad = ultimaColumnaContabilidad;

                int ultimaColumnaAFIP = worksheetArchivoAFIP.LastColumnUsed().ColumnNumber();
                UltimaColumnaAFIP = ultimaColumnaAFIP;

                worksheetArchivoContabilidad.Cell(1, ultimaColumnaContabilidad + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpContabilidad in diccionarioContabilidad)
                {
                    string claveContabilidad = kvpContabilidad.Key;
                    var registrosContabilidad = kvpContabilidad.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveContabilidad))
                    {
                        var registrosAFIP = diccionarioAFIP[claveContabilidad];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosContabilidad = registrosContabilidad.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            int ban = 0;

                            // Señalizar en verde CUIT
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante, importes y fecha
                                if (registroContabilidad.Item4 == registroAFIP.Item4 && registroContabilidad.Item2 == registroAFIP.Item2 * tipoCambio && registroContabilidad.Item3 == registroAFIP.Item3 && registroContabilidad.Item5 == registroAFIP.Item5)
                                {
                                    // Señalizar en verde ambos comprobantes
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Encontramos el comprobante, asignamos bandera
                                    ban = 1;

                                    //Señalizo expresado en USD
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Marcar en verde el IVA
                                     worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                     worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Marcar en verde el TOTAL
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Marcar en verde la FECHA
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                }
                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookContabilidad.SaveAs(rutaExcelContabilidad);
            }
        }

        private async void CompararFechaComprobanteContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha, double tolerancia)
        {
            using (var workbookContabilidad = new XLWorkbook(rutaExcelContabilidad))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoContabilidad = workbookContabilidad.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                //Obtengo indices de las columnas del Excel de AFIP
                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");

                int ultimaColumnaContabilidad = UltimaColumnaContabilidad;

                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                worksheetArchivoContabilidad.Cell(1, ultimaColumnaContabilidad + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpContabilidad in diccionarioContabilidad)
                {
                    string claveContabilidad = kvpContabilidad.Key;
                    var registrosContabilidad = kvpContabilidad.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveContabilidad))
                    {
                        var registrosAFIP = diccionarioAFIP[claveContabilidad];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosContabilidad = registrosContabilidad.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            int ban = 0;

                            // Señalizar en verde CUIT
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante e importes, ignorando la fecha y los señalizados en verde
                                if (registroContabilidad.Item4 == registroAFIP.Item4 && Math.Abs(Math.Abs(registroContabilidad.Item2) - Math.Abs(registroAFIP.Item2)) <= tolerancia && Math.Abs(Math.Abs(registroContabilidad.Item3) - Math.Abs(registroAFIP.Item3)) <= tolerancia &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && 
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Señalizar en verde ambos comprobantes
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Encontramos el comprobante, asignamos bandera
                                    ban = 1;

                                    //Señalizo expresado en USD
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Marcar en verde el IVA
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Marcar en verde el TOTAL
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Comparar la FECHA
                                    if (registroContabilidad.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookContabilidad.SaveAs(rutaExcelContabilidad);
            }
        }

        private async void CompararComprobanteContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha, double tolerancia)
        {
            using (var workbookContabilidad = new XLWorkbook(rutaExcelContabilidad))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoContabilidad = workbookContabilidad.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                //Obtengo indices de las columnas del Excel de AFIP
                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");

                int ultimaColumnaContabilidad = UltimaColumnaContabilidad;

                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                worksheetArchivoContabilidad.Cell(1, ultimaColumnaContabilidad + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpContabilidad in diccionarioContabilidad)
                {
                    string claveContabilidad = kvpContabilidad.Key;
                    var registrosContabilidad = kvpContabilidad.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveContabilidad))
                    {
                        var registrosAFIP = diccionarioAFIP[claveContabilidad];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosContabilidad = registrosContabilidad.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            int ban = 0;

                            // Señalizar en verde CUIT
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante, ignorando la fecha, importes y los señalizados en verde
                                if (registroContabilidad.Item4 == registroAFIP.Item4 &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Señalizar en verde ambos comprobantes
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Encontramos el comprobante, asignamos bandera
                                    ban = 1;

                                    //Señalizo expresado en USD
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Comparar el IVA
                                    if ((Math.Abs(Math.Abs(registroContabilidad.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia))
                                    {
                                        // Señalizo en verde el IVA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el IVA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += "IVA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "IVA esta mal";
                                    }

                                    // Comparar el TOTAL
                                    if ((Math.Abs(Math.Abs(registroContabilidad.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia))
                                    {
                                        // Señalizo en verde el TOTAL
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el TOTAL
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += "TOTAL esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "TOTAL esta mal";
                                    }

                                    // Comparar la FECHA
                                    if (registroContabilidad.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookContabilidad.SaveAs(rutaExcelContabilidad);
            }
        }

        private async void CompararImportesContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, int indiceColumnaFecha, double tolerancia)
        {
            using (var workbookContabilidad = new XLWorkbook(rutaExcelContabilidad))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoContabilidad = workbookContabilidad.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                //Obtengo indices de las columnas del Excel de AFIP
                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Fecha");

                int ultimaColumnaContabilidad = UltimaColumnaContabilidad;

                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                worksheetArchivoContabilidad.Cell(1, ultimaColumnaContabilidad + 1).Value = "Detalle";
                worksheetArchivoAFIP.Cell(1, ultimaColumnaAFIP + 1).Value = "Detalle";

                // Recorrer el diccionario de Holistor
                foreach (var kvpContabilidad in diccionarioContabilidad)
                {
                    string claveContabilidad = kvpContabilidad.Key;
                    var registrosContabilidad = kvpContabilidad.Value;

                    // Verificar si la clave existe en el diccionario de AFIP                   
                    if (diccionarioAFIP.ContainsKey(claveContabilidad))
                    {
                        var registrosAFIP = diccionarioAFIP[claveContabilidad];

                        // Ordenar los registros por el valor numérico del comprobante
                        registrosAFIP = registrosAFIP.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();
                        registrosContabilidad = registrosContabilidad.OrderByDescending(registro => Convert.ToInt64(registro.Item4)).ToList();

                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            int ban = 0;

                            // Señalizar en verde CUIT
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por importes, ignorando comprobante, fecha y los señalizados en verde
                                if (Math.Abs(Math.Abs(registroContabilidad.Item2) - Math.Abs(registroAFIP.Item2)) <= tolerancia && Math.Abs(Math.Abs(registroContabilidad.Item3) - Math.Abs(registroAFIP.Item3)) <= tolerancia &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {                                   
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Encontramos el comprobante, asignamos bandera
                                    ban = 1;

                                    //Señalizo expresado en USD
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Señalizo en verde el IVA
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    // Señalizo en verde el TOTAL
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);                                   

                                    // Comparar la FECHA
                                    if (registroContabilidad.Item5 == registroAFIP.Item5)
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal la FECHA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaFecha).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaFechaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " FECHA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " FECHA esta mal";
                                    }

                                    // Comparar el COMPROBANTE
                                    if (registroContabilidad.Item4 == registroAFIP.Item4)
                                    {
                                        // Señalizar en verde ambos comprobantes
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " COMPROBANTE esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " COMPROBANTE esta mal";
                                    }
                                    else
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo el en Holistor
                        foreach (var registroContabilidad in registrosContabilidad)
                        {
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaCUIT).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }
                }
                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookContabilidad.SaveAs(rutaExcelContabilidad);
            }
        }

        private async void MarcarNoSeñalizadosEnRojoContabilidad(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelContabilidad, int indiceColumnaPuntoVenta, int indiceColumnaComprobante, int indiceColumnaIVA, int indiceColumnaTotal, int indiceColumnaCuit, int indiceColumnaFecha)
        {
            using (var workbookContabilidad = new XLWorkbook(rutaExcelContabilidad))
            {
                var worksheet = workbookContabilidad.Worksheets.First();
                var defaultColor = XLColor.FromIndex(0); // Color predeterminado de Excel

                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuit).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuit).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        string cuit = worksheet.Cell(row.RowNumber(), indiceColumnaCuit).GetString();
                        if (!diccionarioAFIP.ContainsKey(cuit))
                        {
                            worksheet.Cell(row.RowNumber(), indiceColumnaCuit).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheet.Cell(row.RowNumber(), UltimaColumnaContabilidad + 1).Value = "Este cuit no tiene ningun registro en AFIP";
                        }
                    }                        
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuit).Style.Fill.BackgroundColor == XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFecha).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        worksheet.Row(row.RowNumber()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                        worksheet.Cell(row.RowNumber(), UltimaColumnaContabilidad + 1).Value = "NO coincide ningun registro";
                    }
                }

                workbookContabilidad.Save();
            }
        }

        private async void MarcarNoSeñalizadosEnRojo(Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string, DateTime, int)>> diccionarioAFIP, string rutaExcelAFIP)
        {
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheet = workbookAFIP.Worksheets.First();
                var defaultColor = XLColor.FromIndex(0); // Color predeterminado de Excel

                int ultimaColumnaAFIP = UltimaColumnaAFIP;

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");
                int indiceColumnaFechaAFIP = ObtenerIndiceColumna(worksheet, "Fecha");
                int indiceColumnaTipoComprobanteAFIP = ObtenerIndiceColumna(worksheet, "Tipo");

                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        string cuit = worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).GetString();
                        if (!diccionarioHolistor.ContainsKey(cuit))
                        {
                            worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheet.Cell(row.RowNumber(), ultimaColumnaAFIP + 1).Value = "Este cuit no tiene ningun registro en HOLISTOR";
                        }
                    }
                    if (row.RowNumber() != 1 &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaIVAAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTotalAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor == XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaFechaAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204) &&
                        worksheet.Cell(row.RowNumber(), indiceColumnaTipoComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        worksheet.Row(row.RowNumber()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                        worksheet.Cell(row.RowNumber(), ultimaColumnaAFIP + 1).Value = "NO coincide ningun registro";
                    }
                }

                workbookAFIP.SaveAs(rutaExcelAFIP);
            }
        }

        private void buttonDefinirColumnas_Click(object sender, EventArgs e)
        {
            // Abrir el formulario para definir columnas
            FormColumnas columnasForm = new FormColumnas();
            columnasForm.ShowDialog();

            InicializarYMostrarEsquemas();
        }

        private void buttonEditarEsquema_Click(object sender, EventArgs e)
        {
            // Define una lista de esquemas
            List<Esquema> listaEsquemas = new List<Esquema>();

            // Ruta del archivo Esquemas en el directorio de la aplicación
            string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

            // Cargar los esquemas desde el archivo
            CargarEsquemasDesdeArchivo(filePath, listaEsquemas);

            FormColumnas columnasForm = new FormColumnas();

            foreach (Esquema esquema in listaEsquemas)
            {
                if (comboBoxEsquemas.SelectedItem.ToString() == esquema.Nombre)
                {
                    columnasForm.cargarDatos(esquema.IndiceCuit, esquema.IndiceIVA, esquema.IndiceTotal, esquema.IndicePuntoVenta, esquema.IndiceNumeroComprobante, esquema.Nombre, esquema.IndiceFecha);
                }
            }

            columnasForm.ShowDialog();

            InicializarYMostrarEsquemas();
        }
    }

    // Clase para representar un esquema
    class Esquema
    {
        [JsonProperty("Nombre")]
        public string Nombre { get; set; }

        [JsonProperty("IndiceColumnaCuit")]
        public int IndiceCuit { get; set; }

        [JsonProperty("IndiceColumnaPuntoVenta")]
        public int IndicePuntoVenta { get; set; }

        [JsonProperty("IndiceColumnaComprobante")]
        public int IndiceNumeroComprobante { get; set; }

        [JsonProperty("IndiceColumnaIVA")]
        public int IndiceIVA { get; set; }

        [JsonProperty("IndiceColumnaTotal")]
        public int IndiceTotal { get; set; }

        [JsonProperty("IndiceColumnaFecha")]
        public int IndiceFecha { get; set; }

        public Esquema() { }

        public Esquema(int indiceCuit, int indicePuntoVenta, int indiceNumeroComprobante, int indiceIVA, int indiceTotal, string nombre, int indiceFecha)
        {
            IndiceCuit = indiceCuit;
            IndicePuntoVenta = indicePuntoVenta;
            IndiceNumeroComprobante = indiceNumeroComprobante;
            IndiceIVA = indiceIVA;
            IndiceTotal = indiceTotal;
            Nombre = nombre;
            IndiceFecha = indiceFecha;
        }
    }
}