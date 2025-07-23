using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ConsultaBD
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CargarDatos();
        }

        private void CargarDatos()
        {
            try
            {
                string connectionString = "Server=192.168.1.22;Database=DispensaWEB;User Id=sa;Password=ofima.sql10;";
                string query = @"SELECT 
                        M.Prefijo,
                        M.Entrega,
                        M.Orden,
                        M.Fecha,
                        M.FechaVto,
                        M.IdMedicamento,
                        M.PLU,
                        M.Nombre AS Nombre,
                        M.NoPendiente,
                        M.OkPendiente,
                        M.QtyOrden,
                        M.QtyEntrega,
                        M.IdLote,
                        M.Entregas,
                        M.Actual,
                        M.Mipres,
                        M.IdMipresSN,
                        M.DiasTto,
                        M.Valor,
                        M.PctajeIVA,
                        M.PctajeDto,
                        M.IdDomOrden,
                        M.typeOrder,
                        M.TipoTecnologia,
                        M.CodigoTecnologia,
                        E.IdConsecutivo AS IdEntrega,
                        E.IdTipoId AS TipoIdentificacion,
                        E.IdPaciente AS Identificacion,
                        E.FhRegistrado AS FechaCreacion,
                        E.Mensaje,
                        E.idusuario,
                        E.IdConsecutivo,
                        E.idbodega
                    FROM 
                        dbo.tbl_MvEntregas_api M
                    INNER JOIN dbo.tbl_Entregas_api E ON M.Entrega = E.Entrega
                    WHERE 
                        E.Procesado = 0;";

                List<Registro> registros = new List<Registro>();

                using (SqlConnection conn = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                    while (reader.Read())
                    {
                        registros.Add(new Registro
                        {
                            Mensaje = reader["Mensaje"] != DBNull.Value ? reader["Mensaje"].ToString() : null,
                            Usuario = reader["idUsuario"] != DBNull.Value ? reader["idUsuario"].ToString() : null,
                            Consecutivo = reader["idConsecutivo"] != DBNull.Value ? reader["idConsecutivo"].ToString() : null,
                            Bodega = reader["idBodega"] != DBNull.Value ? reader["idBodega"].ToString() : null,
                            Prefijo = reader["Prefijo"]?.ToString(),
                            Entrega = reader["Entrega"] != DBNull.Value ? Convert.ToInt32(reader["Entrega"]) : 0,
                            Orden = reader["Orden"]?.ToString(),
                            PLU = reader["PLU"]?.ToString(),
                            Nombre = reader["Nombre"]?.ToString(),
                            Fecha = reader["Fecha"] as DateTime?,
                            FechaVto = reader["FechaVto"] as DateTime?,
                            IdMedicamento = reader["IdMedicamento"] as int?,
                            NoPendiente = reader["NoPendiente"] as bool?,
                            OkPendiente = reader["OkPendiente"] as bool?,
                            QtyOrden = reader["QtyOrden"] as int?,
                            QtyEntrega = reader["QtyEntrega"] as int?,
                            IdLote = reader["IdLote"]?.ToString(),
                            Entregas = reader["Entregas"] as int?,
                            Actual = reader["Actual"] as int?,
                            Mipres = reader["Mipres"] as bool?,
                            IdMipresSN = reader["IdMipresSN"] as int?,
                            DiasTto = reader["DiasTto"] as int?,
                            Valor = reader["Valor"] as decimal?,
                            PctajeIVA = reader["PctajeIVA"] as decimal?,
                            PctajeDto = reader["PctajeDto"] as decimal?,
                            IdDomOrden = reader["IdDomOrden"] as int?,
                            typeOrder = reader["typeOrder"]?.ToString(),
                            TipoTecnologia = reader["TipoTecnologia"]?.ToString(),
                            CodigoTecnologia = reader["CodigoTecnologia"]?.ToString(),
                            TipoIdentificacion = reader["TipoIdentificacion"]?.ToString(),
                            Identificacion = reader["Identificacion"]?.ToString(),
                            FechaCreacion = reader["FechaCreacion"] as DateTime?
                        });
                    }
                    dgDatos.ItemsSource = registros;
                }
            }


                
                txtEstado.Text = "Última actualización: " + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos:\n" + ex.Message);
            }
        }

        private void btnRefrescar_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void btnCerrar_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
                var registros = dgDatos.ItemsSource as List<Registro>;
                if (registros == null || registros.Count == 0)
                {
                    MessageBox.Show("No hay datos para exportar.");
                    return;
                }

                DataTable dt = ConvertToDataTable(registros);

                Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx",
                    FileName = "Exportacion.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(dt, "Datos");
                        wb.SaveAs(sfd.FileName);
                        MessageBox.Show("Datos exportados a Excel exitosamente.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al generar el excel:\n" + ex.Message);
            }
        }

        private DataTable ConvertToDataTable<T>(List<T> items)
        {
            var dt = new DataTable(typeof(T).Name);
            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (var item in items)
            {
                var values = new object[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }
                dt.Rows.Add(values);
            }

            return dt;
        }


        private void btnExportarPDF_Click(object sender, RoutedEventArgs e)
        {
            var registros = dgDatos.ItemsSource as List<Registro>;
            if (registros == null || registros.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "PDF (*.pdf)|*.pdf",
                FileName = "Exportacion.pdf"
            };

            if (sfd.ShowDialog() == true)
            {
                Document doc = new Document(PageSize.A4.Rotate(), 10, 10, 10, 10);
                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                doc.Open();

                // Definir columnas manualmente
                var columnas = new[] {
            "Mensaje", "TipoIdentificación", "Identificación", "FechaCreacion",
            "Prefijo", "Entrega", "Orden", "PLU", "Nombre", "Fecha", "FechaVto",
            "IdMedicamento", "NoPendiente", "OkPendiente", "QtyOrden", "QtyEntrega",
            "IdLote", "Entregas", "Actual", "Mipres", "IdMipresSN", "DiasTto",
            "Valor", "PctajeIVA", "PctajeDto", "IdDomOrden", "typeOrder",
            "TipoTecnologia", "CodigoTecnologia"
        };

                PdfPTable table = new PdfPTable(columnas.Length)
                {
                    WidthPercentage = 100
                };

                // Encabezados
                foreach (var col in columnas)
                {
                    var cell = new PdfPCell(new Phrase(col))
                    {
                        BackgroundColor = new BaseColor(211, 211, 211),
                        HorizontalAlignment = Element.ALIGN_CENTER
                    };
                    table.AddCell(cell);
                }

                // Filas
                foreach (var r in registros)
                {
                    table.AddCell(r.Mensaje);
                    table.AddCell(r.TipoIdentificacion);
                    table.AddCell(r.Identificacion);
                    table.AddCell(r.FechaCreacion?.ToString("g") ?? "");
                    table.AddCell(r.Prefijo);
                    table.AddCell(r.Entrega.ToString());
                    table.AddCell(r.Orden);
                    table.AddCell(r.PLU);
                    table.AddCell(r.Nombre);
                    table.AddCell(r.Fecha?.ToString("d") ?? "");
                    table.AddCell(r.FechaVto?.ToString("d") ?? "");
                    table.AddCell(r.IdMedicamento?.ToString() ?? "");
                    table.AddCell(r.NoPendiente?.ToString() ?? "");
                    table.AddCell(r.OkPendiente?.ToString() ?? "");
                    table.AddCell(r.QtyOrden?.ToString() ?? "");
                    table.AddCell(r.QtyEntrega?.ToString() ?? "");
                    table.AddCell(r.IdLote);
                    table.AddCell(r.Entregas?.ToString() ?? "");
                    table.AddCell(r.Actual?.ToString() ?? "");
                    table.AddCell(r.Mipres?.ToString() ?? "");
                    table.AddCell(r.IdMipresSN?.ToString() ?? "");
                    table.AddCell(r.DiasTto?.ToString() ?? "");
                    table.AddCell(r.Valor?.ToString("N2") ?? "");
                    table.AddCell(r.PctajeIVA?.ToString("N2") ?? "");
                    table.AddCell(r.PctajeDto?.ToString("N2") ?? "");
                    table.AddCell(r.IdDomOrden?.ToString() ?? "");
                    table.AddCell(r.typeOrder);
                    table.AddCell(r.TipoTecnologia);
                    table.AddCell(r.CodigoTecnologia);
                }

                doc.Add(table);
                doc.Close();
                writer.Close();

                MessageBox.Show("Datos exportados a PDF exitosamente.");
        }

    }

}

    public class Registro
    {
        public string Mensaje { get; set; }
        public string Usuario { get; set; }
        public string Consecutivo { get; set; }
        public string Bodega { get; set; }
        public string TipoIdentificacion { get; set; }
        public string Identificacion { get; set; }
        public DateTime? FechaCreacion { get; set; }
        public string Prefijo { get; set; }
        public int Entrega { get; set; }
        public string Orden { get; set; }
        public string PLU { get; set; }
        public string Nombre { get; set; }
        public DateTime? Fecha { get; set; }
        public DateTime? FechaVto { get; set; }
        public int? IdMedicamento { get; set; }
        public bool? NoPendiente { get; set; }
        public bool? OkPendiente { get; set; }
        public int? QtyOrden { get; set; }
        public int? QtyEntrega { get; set; }
        public string IdLote { get; set; }
        public int? Entregas { get; set; }
        public int? Actual { get; set; }
        public bool? Mipres { get; set; }
        public int? IdMipresSN { get; set; }
        public int? DiasTto { get; set; }
        public decimal? Valor { get; set; }
        public decimal? PctajeIVA { get; set; }
        public decimal? PctajeDto { get; set; }
        public int? IdDomOrden { get; set; }
        public string typeOrder { get; set; }
        public string TipoTecnologia { get; set; }
        public string CodigoTecnologia { get; set; }
    }

}
