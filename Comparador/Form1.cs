using ClosedXML.Excel;
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
        public Form1()
        {
            InitializeComponent();
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

            // Realizar el proceso de manera asíncrona
            await Task.Run(() => RealizarComparacion(textBoxAfip.Text, textBoxHolistor.Text));

            pictureBoxRuedaCargando.Visible = false;

            // Muestra un mensaje de éxito
            MessageBox.Show("Proceso completado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static void RealizarComparacion(string rutaExcelAfip, string rutaExcelHolistor)
        {
            var diccionarioHolistor = ArmarDiccionarioHolistor(rutaExcelHolistor);
            var diccionarioAFIP = ArmarDiccionarioAFIP(rutaExcelAfip);

            CompararYMarcarFilas(diccionarioHolistor, diccionarioAFIP, rutaExcelHolistor, rutaExcelAfip);
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

            // Insertar guion en la posición 5
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
        static string LimpiarCUITHolistor(string cuit)
        {
            return Regex.Replace(cuit, @"[^\d]", "");
        }

        // Función para armar el diccionario de Holistor
        static Dictionary<string, List<(int, double, double, double, string)>> ArmarDiccionarioHolistor(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, double, string)>>();

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

                        // Obtener valor de la columna Neto
                        int indiceColumnaNeto = ObtenerIndiceColumna(worksheet, "Neto");
                        string valorCeldaNeto = worksheet.Cell(fila, indiceColumnaNeto).GetString();
                        string valorCeldaNetoSinComa = valorCeldaNeto.Replace(",", ".");
                        double neto = double.Parse(valorCeldaNetoSinComa, CultureInfo.InvariantCulture);

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
                        cuit = LimpiarCUITHolistor(cuit);

                        // Agregar al diccionario
                        if (!diccionario.ContainsKey(cuit))
                        {
                            diccionario[cuit] = new List<(int, double, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, neto, iva, total, numeroComprobante)); 
                    }
                }
                else
                {
                    Console.WriteLine("La columna 'Comprobante' no se encontró en el Excel.");
                }
            }

            return diccionario;
        }

        // Función para armar el diccionario de AFIP
        static Dictionary<string, List<(int, double, double, double, string)>> ArmarDiccionarioAFIP(string rutaExcel)
        {
            var diccionario = new Dictionary<string, List<(int, double, double, double, string)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja
                int indiceColumnaPuntoVenta = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobante = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaNeto = ObtenerIndiceColumna(worksheet, "Imp. Neto Gravado");
                int indiceColumnaIVA = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaTotal = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaCuit = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");

                if (indiceColumnaPuntoVenta != -1 && indiceColumnaComprobante != -1 && indiceColumnaNeto != -1 && indiceColumnaIVA != -1 && indiceColumnaTotal != -1 && indiceColumnaCuit != -1)
                {
                    int ultimaFila = worksheet.LastRowUsed().RowNumber();

                    for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                    {
                        string puntoVenta = worksheet.Cell(fila, indiceColumnaPuntoVenta).Value.ToString();
                        string numeroComprobante = worksheet.Cell(fila, indiceColumnaComprobante).Value.ToString();
                        string comprobanteCompleto = puntoVenta + numeroComprobante;

                        // Obtener valor de la columna Neto
                        string valorCeldaNeto = worksheet.Cell(fila, indiceColumnaNeto).GetString();
                        string valorCeldaNetoSinComa = valorCeldaNeto.Replace(",", ".");
                        double neto = 0;
                        if (valorCeldaNetoSinComa != "")
                        {
                            neto = double.Parse(valorCeldaNetoSinComa, CultureInfo.InvariantCulture);
                        }

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
                            diccionario[cuit] = new List<(int, double, double, double, string)>();
                        }

                        diccionario[cuit].Add((fila, neto, iva, total, comprobanteCompleto));
                    }
                }
                else
                {
                    Console.WriteLine("Alguna de las columnas necesarias no se encontró en el Excel.");
                }
            }

            return diccionario;
        }

        static void CompararYMarcarFilas(Dictionary<string, List<(int, double, double, double, string)>> diccionarioHolistor, Dictionary<string, List<(int, double, double, double, string)>> diccionarioAFIP, string rutaExcelHolistor, string rutaExcelAFIP)
        {
            using (var workbookHolistor = new XLWorkbook(rutaExcelHolistor))
            using (var workbookAFIP = new XLWorkbook(rutaExcelAFIP))
            {
                // Recorrer el diccionario de Holistor
                foreach (var kvpHolistor in diccionarioHolistor)
                {
                    string claveHolistor = kvpHolistor.Key;
                    var registrosHolistor = kvpHolistor.Value;

                    // Verificar si la clave existe en el diccionario de AFIP
                    if (diccionarioAFIP.ContainsKey(claveHolistor))
                    {
                        var registrosAFIP = diccionarioAFIP[claveHolistor];

                        foreach (var registroHolistor in registrosHolistor)
                        {
                            foreach (var registroAFIP in registrosAFIP)
                            {
                                // Comparar los valores de neto, iva, total y comprobante
                                if (Math.Abs(registroHolistor.Item2 - registroAFIP.Item2) <= 5 &&
                                    Math.Abs(registroHolistor.Item3 - registroAFIP.Item3) <= 5 &&
                                    Math.Abs(registroHolistor.Item4 - registroAFIP.Item4) <= 5 &&
                                    registroHolistor.Item5 == registroAFIP.Item5)
                                {
                                    // Coinciden todos los valores, señalizar en verde ambas filas
                                    MarcarFila(workbookHolistor, registroHolistor.Item1, XLColor.FromArgb(255, 204, 255, 204)); // Tono de verde claro
                                    MarcarFila(workbookAFIP, registroAFIP.Item1, XLColor.FromArgb(255, 204, 255, 204)); // Tono de verde claro
                                }
                                else
                                {
                                    // No coinciden todos los valores, señalizar en rojo ambas filas
                                    MarcarFila(workbookHolistor, registroHolistor.Item1, XLColor.Red);
                                    MarcarFila(workbookAFIP, registroAFIP.Item1, XLColor.Red);
                                }
                            }
                        }
                    }
                    else
                    {
                        // La clave no existe en el diccionario de AFIP, señalizar en rojo en Holistor
                        foreach (var registroHolistor in registrosHolistor)
                        {
                            MarcarFila(workbookHolistor, registroHolistor.Item1, XLColor.Red);
                        }
                    }
                }

                // Recorrer el diccionario de AFIP para marcar las filas que no se han marcado
                foreach (var kvpAFIP in diccionarioAFIP)
                {
                    string claveAFIP = kvpAFIP.Key;
                    var registrosAFIP = kvpAFIP.Value;

                    // Verificar si la clave no se ha marcado en Holistor
                    if (!diccionarioHolistor.ContainsKey(claveAFIP))
                    {
                        // Señalizar en rojo en AFIP
                        foreach (var registroAFIP in registrosAFIP)
                        {
                            MarcarFila(workbookAFIP, registroAFIP.Item1, XLColor.Red);
                        }
                    }
                }

                workbookAFIP.SaveAs(rutaExcelAFIP);
                workbookHolistor.SaveAs(rutaExcelHolistor);
            }
        }


        static void MarcarFila(XLWorkbook workbook, int numeroFila, XLColor color)
        {
            var worksheet = workbook.Worksheet(1); // Supongamos que los datos están en la primera hoja

            // Obtener la fila correspondiente y marcarla con el color especificado
            var fila = worksheet.Row(numeroFila);
            fila.Style.Fill.BackgroundColor = color;
        }
    }
}


