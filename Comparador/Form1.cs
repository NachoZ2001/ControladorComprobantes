using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Comparador
{
    public partial class Form1 : Form
    {

        public List<int> columnas { get; set; }

        public Form1()
        {
            InitializeComponent();

            columnas = new List<int>();

            // Establecer el estilo del borde y deshabilitar el cambio de tamaño
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Establecer el tamaño mínimo y máximo para evitar el cambio de tamaño
            this.MinimumSize = this.MaximumSize = this.Size;

            comboBoxEsquemas.Text = "Seleccionar un esquema";

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
            pictureBoxRuedaCargando.Visible = true;

            // Define una lista de esquemas
            List<Esquema> listaEsquemas = new List<Esquema>();

            // Ruta del archivo Esquemas en el directorio de la aplicación
            string filePath = Path.Combine(Application.StartupPath, "Esquemas.txt");

            // Cargar los esquemas desde el archivo
            CargarEsquemasDesdeArchivo(filePath, listaEsquemas);

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

        private void RealizarComparacion(string rutaExcelAfip, string rutaExcelContabilidad)
        {
            int indiceColumnaCUIT = columnas[0];
            int indiceColumnaPuntoVenta = columnas[1];
            int indiceColumnaNumeroComprobante = columnas[2];
            int indiceColumnaIVA = columnas[3];
            int indiceColumnaTotal = columnas[4];

            //Armar diccionarios
            var diccionarioContabilidad = ArmarDiccionarioContabilidad(rutaExcelContabilidad, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA);
            var diccionarioAFIP = ArmarDiccionarioAFIP(rutaExcelAfip);

            //Comparar y marcar filas 
            CompararYMarcarFilasContabilidad(diccionarioContabilidad, diccionarioAFIP, rutaExcelContabilidad, rutaExcelAfip, indiceColumnaCUIT, indiceColumnaPuntoVenta, indiceColumnaNumeroComprobante, indiceColumnaTotal, indiceColumnaIVA, (double)numericUpDownTolerancia.Value);

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

        // Función para armar el diccionario de Contabilidad --> {CUIT}: (fila, iva, total, comprobante)
        static Dictionary<string, List<(int, double, double, string)>> ArmarDiccionarioContabilidad(string rutaExcel, int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string)>>();

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

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, numeroComprobante));
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

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, comprobanteCompleto));
                    }
                }
            }

            return diccionario;
        }

        // Función para armar el diccionario de Holistor --> {CUIT}: (fila, iva, total, comprobante)
        static Dictionary<string, List<(int, double, double, string)>> ArmarDiccionarioHolistor(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string)>>();

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

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, numeroComprobante));
                    }
                }
                else
                {
                    Console.WriteLine("La columna 'Comprobante' no se encontró en el Excel.");
                }
            }

            return diccionario;
        }

        // Función para armar el diccionario de AFIP --> {CUIT}: (fila, iva, total, comprobante)
        static Dictionary<string, List<(int, double, double, string)>> ArmarDiccionarioAFIP(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, string)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja
                int indiceColumnaPuntoVenta = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobante = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaIVA = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotal = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaCuit = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");

                if (indiceColumnaPuntoVenta != -1 && indiceColumnaComprobante != -1 && indiceColumnaIVA != -1 && indiceColumnaTotal != -1 && indiceColumnaCuit != -1)
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

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, iva, total, comprobanteCompleto));
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
        static void CompararYMarcarFilasHolistor(Dictionary<string, List<(int, double, double, string)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP, double tolerancia)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheetArchivoHolistor = workbookHolistor.Worksheets.First();
                var worksheetArchivoAFIP = workbookAFIP.Worksheets.First();

                int ultimaColumnaHolistor = worksheetArchivoHolistor.LastColumnUsed().ColumnNumber();
                int ultimaColumnaAFIP = worksheetArchivoAFIP.LastColumnUsed().ColumnNumber();

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheetArchivoAFIP, "Nro. Doc. Emisor");

                int indiceColumnaComprobanteHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Comprobante");
                int indiceColumnaIVAHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "IVA");
                int indiceColumnaTotalHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Total");
                int indiceColumnaCuitHolistor = ObtenerIndiceColumna(worksheetArchivoHolistor, "Tipo/Nro.Doc.");

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
                            int ban = 0;

                            // Señalizar en verde CUIT
                            worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaCuitHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                            foreach (var registroAFIP in registrosAFIP)
                            {
                                int indiceTipoCambio = ObtenerIndiceColumna(worksheetArchivoAFIP, "Tipo Cambio");
                                double tipoCambio = (double)worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceTipoCambio).Value;

                                // Señalizar en verde CUIT
                                worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                // Comparamos por comprobante primero
                                if (registroHolistor.Item4 == registroAFIP.Item4 && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Señalizar en verde ambos comprobantes
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = " ";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value = " ";

                                    // Encontramos el comprobante, asignamos bandera
                                    ban = 1;

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    // Comparar el IVA
                                    if (Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el IVA
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " IVA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " IVA esta mal";
                                    }

                                    // Comparar el TOTAL
                                    if (Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia)
                                    {
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el TOTAL
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += " TOTAL esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " TOTAL esta mal";
                                    }
                                }
                                else if ((Math.Abs(Math.Abs(registroHolistor.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia) && (Math.Abs(Math.Abs(registroHolistor.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia)
                                    && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204)
                                    && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204)
                                    && worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Coinciden los total y los importe pero no el comprobante
                                    ban = 1;

                                    //Señalizo expresado en dolares
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    //Señalizo en verde ambos TOTAL
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaTotalHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    //Señalizo en verde ambos IVA
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaIVAHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    //Señalizo en rojo ambos comprobantes
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, indiceColumnaComprobanteHolistor).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value += "COMPROBANTE esta mal";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "COMPROBANTE esta mal";

                                }
                            }
                            if (ban == 0)
                            {
                                // No se encontro ninguno que coincida, señalizo en rojo todas las columnas en holistor
                                worksheetArchivoHolistor.Row(registroHolistor.Item1).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                worksheetArchivoHolistor.Cell(registroHolistor.Item1, ultimaColumnaHolistor + 1).Value = "NO coincide ningun registro";
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

        //Comparacion para los archivos de contabilidad que no son de Holistor
        static void CompararYMarcarFilasContabilidad(Dictionary<string, List<(int, double, double, string)>> diccionarioContabilidad, Dictionary<string, List<(int, double, double, string)>> diccionarioAFIP, string rutaExcelContabilidad, string rutaExcelAFIP,
            int indiceColumnaCUIT, int indiceColumnaPuntoVenta, int indiceColumnaNumeroComprobante, int indiceColumnaTotal, int indiceColumnaIVA, double tolerancia)
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

                int ultimaColumnaContabilidad = worksheetArchivoContabilidad.LastColumnUsed().ColumnNumber();
                int ultimaColumnaAFIP = worksheetArchivoAFIP.LastColumnUsed().ColumnNumber();

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

                                // Comparamos por comprobante primero
                                if (registroContabilidad.Item4 == registroAFIP.Item4 && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
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
                                    if (Math.Abs(Math.Abs(registroContabilidad.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia)
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el IVA
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " IVA esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " IVA esta mal";
                                    }

                                    // Comparar el TOTAL
                                    if (Math.Abs(Math.Abs(registroContabilidad.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia)
                                    {
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    }
                                    else
                                    {
                                        //Esta mal el TOTAL
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += " TOTAL esta mal";
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += " TOTAL esta mal";
                                    }
                                }
                                else if ((Math.Abs(Math.Abs(registroContabilidad.Item2) - Math.Abs(registroAFIP.Item2 * tipoCambio)) <= tolerancia) && (Math.Abs(Math.Abs(registroContabilidad.Item3) - Math.Abs(registroAFIP.Item3 * tipoCambio)) <= tolerancia)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204)
                                    && worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 204, 255, 204))
                                {
                                    // Coinciden los total y los importe pero no el comprobante
                                    ban = 1;

                                    //Señalizo expresado en USD
                                    if (tipoCambio != 1)
                                    {
                                        int indiceMoneda = ObtenerIndiceColumna(worksheetArchivoAFIP, "Moneda");
                                        string moneda = worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceMoneda).Value.ToString();
                                        worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += $"Expresado en {moneda}";
                                        worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = $"Expresado en {moneda} en AFIP";
                                    }

                                    //Señalizo en verde ambos TOTAL
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaTotal).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaTotalAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    //Señalizo en verde ambos IVA
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaIVA).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaIVAAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 204, 255, 204);

                                    //Señalizo en rojo ambos comprobantes
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaPuntoVenta).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, indiceColumnaNumeroComprobante).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaPuntoVentaAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, indiceColumnaComprobanteAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                    worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value += "COMPROBANTE esta mal";
                                    worksheetArchivoAFIP.Cell(registroAFIP.Item1, ultimaColumnaAFIP + 1).Value += "COMPROBANTE esta mal";

                                }
                            }
                            if (ban == 0)
                            {
                                // No se encontro ninguno que coincida, señalizo en rojo todas las columnas en holistor
                                worksheetArchivoContabilidad.Row(registroContabilidad.Item1).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                                worksheetArchivoContabilidad.Cell(registroContabilidad.Item1, ultimaColumnaContabilidad + 1).Value = "NO coincide ningun registro";
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

        static void MarcarNoSeñalizadosEnRojo(Dictionary<string, List<(int, double, double, string)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, string)>> diccionarioAFIP, string rutaExcelAFIP)
        {
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                var worksheet = workbookAFIP.Worksheets.First();
                var defaultColor = XLColor.FromIndex(0); // Color predeterminado de Excel

                int ultimaColumnaAFIP = worksheet.LastColumnUsed().ColumnNumber();

                int indiceColumnaPuntoVentaAFIP = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobanteAFIP = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaIVAAFIP = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotalAFIP = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaCuitAFIP = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");

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
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor != XLColor.FromArgb(255, 255, 204, 204))
                    {
                        string cuit = worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).GetString();
                        if (!diccionarioHolistor.ContainsKey(cuit))
                        {
                            worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                            worksheet.Cell(row.RowNumber(), ultimaColumnaAFIP).Value = "Este cuit no tiene ningun registro en HOLISTOR";
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
                        worksheet.Cell(row.RowNumber(), indiceColumnaCuitAFIP).Style.Fill.BackgroundColor == XLColor.FromArgb(255, 204, 255, 204))
                    {
                        worksheet.Row(row.RowNumber()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 204, 204);
                        worksheet.Cell(row.RowNumber(), ultimaColumnaAFIP).Value = "NO coincide ningun registro";
                    }
                }

                workbookAFIP.Save();
            }
        }

        private void buttonDefinirColumnas_Click(object sender, EventArgs e)
        {
            // Abrir el formulario para definir columnas
            FormColumnas columnasForm = new FormColumnas();
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

        public Esquema() { }

        public Esquema(int indiceCuit, int indicePuntoVenta, int indiceNumeroComprobante, int indiceIVA, int indiceTotal, string nombre)
        {
            IndiceCuit = indiceCuit;
            IndicePuntoVenta = indicePuntoVenta;
            IndiceNumeroComprobante = indiceNumeroComprobante;
            IndiceIVA = indiceIVA;
            IndiceTotal = indiceTotal;
            Nombre = nombre;
        }
    }
}