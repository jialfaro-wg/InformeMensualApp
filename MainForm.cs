using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;


namespace Informe_Mensual
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ruta del archivo Excel
            string excelPath = "E:\\Proyectos\\Joaquin\\Web Services\\INFORME\\historial tk y recursostst.xlsx";
            string pdfPath = "InformeMensual.pdf";

            // Mostrar información de tickets y generar el informe PDF
            MostrarInformacionTickets(excelPath, pdfPath);

            Console.WriteLine("Proceso completado. Presiona cualquier tecla para salir.");
            Console.ReadKey();
        }

        // Función principal para mostrar información de los tickets y generar el informe PDF
        public static void MostrarInformacionTickets(string excelPath, string pdfPath)
        {
            // Usar FileStream para leer el archivo Excel
            using (FileStream excel = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                // Crear una instancia de la clase XSSFWorkbook (biblioteca NPOI) para trabajar con el Excel
                IWorkbook workbook = new XSSFWorkbook(excel);

                // Obtener la primera hoja del libro
                ISheet hoja = workbook.GetSheetAt(0);

                // Declarar variables para contar los tickets
                int totalTickets = 0;
                int ticketsEnProceso = 0;
                int ticketsWSIncidente = 0;
                int ticketsWSRequerimiento = 0;

                // Obtener los índices de las columnas necesarias
                int columnaNumeroTicket = ObtenerIndiceColumna(hoja, "N°TK");
                int columnaEstado = ObtenerIndiceColumna(hoja, "Estado");
                int columnaWebService = ObtenerIndiceColumna(hoja, "WebService");
                int columnaTipo = ObtenerIndiceColumna(hoja, "Tipo");

                // Crear el documento PDF
                using (FileStream pdfFile = new FileStream(pdfPath, FileMode.Create))
                {
                    PdfWriter pdfWriter = new PdfWriter(pdfFile);
                    PdfDocument pdfDocument = new PdfDocument(pdfWriter);
                    Document document = new Document(pdfDocument);

                    // Imprimir encabezado del informe en el PDF
                    document.Add(new Paragraph("-----------------------------------------------------"));
                    document.Add(new Paragraph("----------------INFORME MENSUAL TKT------------------"));
                    document.Add(new Paragraph("-----------------------------------------------------"));

                    // Verificar si se encontraron todas las columnas necesarias
                    if (columnaNumeroTicket == -1 || columnaEstado == -1 || columnaWebService == -1 || columnaTipo == -1)
                    {
                        document.Add(new Paragraph("No se encontraron las columnas necesarias en el archivo."));
                        return;
                    }

                    // Diccionario para almacenar la cantidad de tickets por cada WS
                    Dictionary<string, int> ticketsPorWebService = new Dictionary<string, int>();

                    // Iterar sobre las filas del Excel
                    for (int row = 1; row <= hoja.LastRowNum; row++)
                    {
                        IRow currentRow = hoja.GetRow(row);

                        // Verificar si la fila actual no es nula y si la celda en la columna "N°TK" no está vacía
                        if (currentRow != null && currentRow.GetCell(columnaNumeroTicket) != null && !string.IsNullOrEmpty(currentRow.GetCell(columnaNumeroTicket).ToString()))
                        {
                            totalTickets++;

                            // Obtener los valores de las celdas correspondientes
                            string numeroTicket = currentRow.GetCell(columnaNumeroTicket)?.ToString();  // Convertir a string
                            string estado = currentRow.GetCell(columnaEstado)?.ToString(); // Convertir a string

                            // Verificar si el ticket está en proceso o pendiente
                            if (!string.IsNullOrEmpty(estado) && (estado.ToLower() == "en proceso" || estado.ToLower() == "pendiente (en cola)"))
                            {
                                ticketsEnProceso++;
                                document.Add(new Paragraph($"Ticket en proceso - N°TK: {numeroTicket}"));
                            }

                            // Verificar si el ticket tiene un WS asociado
                            bool tieneWebService = !string.IsNullOrEmpty(currentRow.GetCell(columnaWebService)?.StringCellValue);

                            if (tieneWebService)
                            {
                                // Obtener el tipo de Web Service
                                string tipoWebService = currentRow.GetCell(columnaTipo)?.StringCellValue;

                                // Actualizar el conteo de tickets por WS
                                if (!string.IsNullOrEmpty(tipoWebService))
                                {
                                    if (ticketsPorWebService.ContainsKey(tipoWebService))
                                    {
                                        ticketsPorWebService[tipoWebService]++;
                                    }
                                    else
                                    {
                                        ticketsPorWebService[tipoWebService] = 1;
                                    }

                                    // Verificar si el Web Service está asociado a un incidente o requerimiento
                                    if (tipoWebService.ToLower() == "incidente")
                                    {
                                        ticketsWSIncidente++;
                                        document.Add(new Paragraph($"Ticket Web Service (incidente) - N°TK: {numeroTicket}"));
                                    }
                                    else if (tipoWebService.ToLower() == "requerimiento")
                                    {
                                        ticketsWSRequerimiento++;
                                        document.Add(new Paragraph($"Ticket Web Service (requerimiento) - N°TK: {numeroTicket}"));
                                    }
                                }
                            }

                            // Resto de la lógica para procesar los tickets según tus necesidades
                            // ...
                        }
                    }

                    // Agregar resultados al documento PDF
                    document.Add(new Paragraph($"Cantidad total de tickets: {totalTickets}"));
                    document.Add(new Paragraph($"Cantidad de tickets en trámite: {ticketsEnProceso}"));

                    // Mostrar la cantidad de tickets por cada WS en el PDF
                    foreach (var kvp in ticketsPorWebService)
                    {
                        document.Add(new Paragraph($"{kvp.Key} --> {kvp.Value} "));
                    }

                    // Llamada a la función para contar los tickets por cada WS
                    ContarTicketsPorWebService(document, hoja, columnaWebService, columnaTipo);
                }
            }
        }

        // Método para obtener el índice de una columna por su nombre
        public static int ObtenerIndiceColumna(ISheet hoja, string nombreColumna)
        {
            IRow headerRow = hoja.GetRow(0);

            if (headerRow != null)
            {
                for (int i = 0; i < headerRow.LastCellNum; i++)
                {
                    ICell cell = headerRow.GetCell(i);

                    if (cell != null && cell.StringCellValue == nombreColumna)
                    {
                        return i;
                    }
                }
            }

            return -1; // Devolver -1 si la columna no se encuentra
        }

        // Función para contar los tickets por cada WS y su tipo
        public static void ContarTicketsPorWebService(Document document, ISheet hoja, int columnaWebService, int columnaTipo)
        {
            // Lógica para contar tickets por cada Web Service y su tipo
            // ...

            // Agregar resultados al documento PDF
            // ...
        }
    }
}
