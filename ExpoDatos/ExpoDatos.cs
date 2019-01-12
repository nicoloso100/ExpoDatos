using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Timers;
using System.Windows.Forms;
using System.IO;
using System.Net.Sockets;
using System.Net;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Security.Cryptography;
using System.Net.Mail;
using System.Diagnostics;
using MySql.Data.MySqlClient;
using System.Threading;

namespace ExpoDatos
{
    public partial class ExpoDatos : Form
    {
        #region Panel

        #region Inicio

        //Directorio
        public static string Direccion = Directory.GetCurrentDirectory();
        public static string Carpeta = @Direccion;
        public static string SubCarpeta = Path.Combine(Carpeta, "Configuracion");
        public static string Archivo = Path.Combine(SubCarpeta, "Configuracion.txt");

        //Muestra en DataGridView
        DataTable Dtable;
        MySqlConnection Conexion;
        MySqlDataAdapter adapter;
        public static string conexion;
        List<string> DatosRow = new List<string>();
        int Celdas = 0;
        int Posicion = 0;

        //Timer
        System.Timers.Timer aTimer= new System.Timers.Timer();
        int TramasNuevas;
        int TramasNuevasAnt;
        int TramasNuevasDesp;
        string query;
        MySqlCommand comando;
        MySqlCommand comando2;
        MySqlDataReader lee;
        MySqlDataReader lee2;

        //Nuevas Tramas
        DataRow row;

        //Contraseña
        public static bool Salir = false;
        public bool Permitido = false;

        //Guarda configRecep
        public static List<string> Config = new List<string>();
        //List<string> Parametros = new List<string>();
        bool SaleEm = false;

        //Progress bar
        delegate void vbar(int v);
        delegate void sbar(string s);


        //licencia
        System.Timers.Timer bTimer;

        //Reportes programados
        System.Timers.Timer cTimer;
        MensajeReporteP MP;

        public ExpoDatos()
        {
            InitializeComponent();
            this.FormClosing += ExpoDatos_FormClosing;
            this.Load += ExpoDatos_Load;
        }

        private void ExpoDatos_Load(object sender, EventArgs e)
        {
            if (YaAbierto() == false)
            {
                IniciaConexion();
                IniciaConRecep();
                Revisa();
            }
            else
            {
                SaleEm = true;
                Application.Exit();
            }
        }

        private void ExpoDatos_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SaleEm == false)
            {
                Salir = false;
                Contraseña cn = new Contraseña();
                cn.ShowDialog();

                if (Salir == false)
                {
                    e.Cancel = true;
                }
            }
        }

        private static bool YaAbierto()
        {
            string currPrsName = Process.GetCurrentProcess().ProcessName;
            Process[] allProcessWithThisName = Process.GetProcessesByName(currPrsName);
            if (allProcessWithThisName.Length > 1)
            {
                MessageBox.Show("El programa ya está abierto!");
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region ConexionSQL

        public void IniciaConexion()
        {
            Config = new List<string>();
            if (File.Exists(Archivo))
            {
                try
                {
                    using (StreamReader lector = new StreamReader(Archivo))
                    {
                        string Line = "";
                        while ((Line = lector.ReadLine()) != null)
                        {
                            Config.Add(Line);
                        }
                        if (Config.Count == 1 || Config.Count == 5)
                        {
                            SeleccionaDB SDB = new SeleccionaDB();
                            SDB.ShowDialog();
                            MP = new MensajeReporteP();
                            MP.label1.Text = "Iniciando el programa, por favor espere...";
                            MP.Show();
                            Application.DoEvents();
                            if(Config.Count == 1)
                            {
                                conexion = Config[0];
                                MessageBox.Show("No se ha establecio una conexión con RecepDatos");
                            }
                            else if (Config.Count == 5)
                            {
                                conexion = Config[0];
                                IP1 = Config[1];
                                P1 = Config[2];
                                IP2 = Config[3];
                                P2 = Config[4];
                                lector.Close();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ocurrió una confusión al leer el archivo de configuración");
                        }
                    }
                    using (MySqlConnection Conexion = new MySqlConnection(conexion))
                    {
                        try
                        {
                            Conexion.Open();
                            label69.BackColor = System.Drawing.Color.FromArgb(146, 208, 80);
                            label69.Text = "SI";
                            CargaDatos();
                        }
                        catch
                        {
                            MessageBox.Show("No se ha podido establecer conexión con SQL, revise los datos de conexión");
                            ConexionSQL sql = new ConexionSQL();
                            sql.ShowDialog();
                            if (ConexionSQL.Exit == false)
                            {
                                IniciaConexion();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer el archivo Configuracion.txt\n\n" + e.ToString());
                }
            }
            else
            {
                MessageBox.Show("No se ha detectado una configuración de SQL");
                ConexionSQL sql = new ConexionSQL();
                sql.ShowDialog();
                if (ConexionSQL.Exit == false)
                {
                    ConexionSQL.Exit = true;
                    label69.BackColor = System.Drawing.Color.FromArgb(146, 208, 80);
                    label69.Text = "SI";
                    CargaDatos();
                }
            }
        }

        #endregion

        #region Carga Datos

        public void CargaDatos()
        {
            try {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    using (adapter = new MySqlDataAdapter("Select * From llamadas_telefonicas", conexion))
                    {
                        Dtable = new DataTable();
                        adapter.Fill(Dtable);
                        dataGridView1.DataSource = Dtable;
                    }
                    Conexion.Close();
                }
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[6].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[9].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = false;
                dataGridView1.Columns[14].Visible = false;
                dataGridView1.Columns[15].Visible = false;
                dataGridView1.Columns[17].Visible = false;
                dataGridView1.Columns[18].Visible = false;
                dataGridView1.Columns[19].Visible = false;
                dataGridView1.Columns[20].Visible = false;
                dataGridView1.Columns[22].Visible = false;
                dataGridView1.Columns[23].Visible = false;
                dataGridView1.Columns[24].Visible = false;
                dataGridView1.Columns[25].Visible = false;
                dataGridView1.Columns[26].Visible = false;
                dataGridView1.Columns[27].Visible = false;
                dataGridView1.Columns[28].Visible = false;
                dataGridView1.Columns[29].Visible = false;
                dataGridView1.Columns[30].Visible = false;
                dataGridView1.Columns[31].Visible = false;
                dataGridView1.Columns[32].Visible = false;
                dataGridView1.Columns[33].Visible = false;
                dataGridView1.Columns[34].Visible = false;
                dataGridView1.Columns[35].Visible = false;
                dataGridView1.Columns[36].Visible = false;
                dataGridView1.Columns[37].Visible = false;
                dataGridView1.Columns[38].Visible = false;
                dataGridView1.Columns[40].Visible = false;
                dataGridView1.Columns[41].Visible = false;
                dataGridView1.Columns[42].Visible = false;
                dataGridView1.Columns[43].Visible = false;
                dataGridView1.Columns[44].Visible = false;
                dataGridView1.Columns[45].Visible = false;
                dataGridView1.Columns[46].Visible = false;
                dataGridView1.Columns[47].Visible = false;
                dataGridView1.Columns[48].Visible = false;
                dataGridView1.Columns[49].Visible = false;
                dataGridView1.Columns[50].Visible = false;
                dataGridView1.Columns[51].Visible = false;

                dataGridView1.Columns[1].HeaderText = "Fecha"; dataGridView1.Columns[1].Width = 65; dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[2].HeaderText = "Hora"; dataGridView1.Columns[2].Width = 65; dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[3].HeaderText = "Extension"; dataGridView1.Columns[3].Width = 65; dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[4].HeaderText = "Troncal"; dataGridView1.Columns[4].Width = 65; dataGridView1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[21].HeaderText = "CL.Ext"; dataGridView1.Columns[21].Width = 55; dataGridView1.Columns[21].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[5].HeaderText = "Numero"; dataGridView1.Columns[5].Width = 160; dataGridView1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[13].HeaderText = "Destino"; dataGridView1.Columns[13].Width = 150; dataGridView1.Columns[13].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[12].HeaderText = "CL.Llam"; dataGridView1.Columns[12].Width = 60; dataGridView1.Columns[12].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[16].HeaderText = "Dur"; dataGridView1.Columns[16].Width = 55; dataGridView1.Columns[16].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[39].HeaderText = "Valor"; dataGridView1.Columns[39].Width = 85; dataGridView1.Columns[39].SortMode = DataGridViewColumnSortMode.NotSortable;

                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204);

                IniciaPrograma();
            }
            catch
            {
                MessageBox.Show("Ocurrió un problema al leer la base de datos! revise la conexión");
                ConexionSQL sql = new ConexionSQL();
                sql.ShowDialog();
                if (ConexionSQL.Exit == false)
                {
                    CargaDatos();
                }

            }
            finally
            {
                MP.Hide();
                MP.label1.Text = "Enviando reporte programado, por favor espere";
            }
        }

        #endregion

        #region Selecciona Trama

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0)
            {

                dataGridView1.Rows[e.RowIndex].Selected = true;
                Celdas = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRow = new List<string>();
                while (Posicion < Celdas)
                {
                    DatosRow.Add(dataGridView1.CurrentRow.Cells[Posicion].Value.ToString());
                    Posicion++;
                }
                label1.Text = DatosRow[40];
                label15.Text = DatosRow[3].Replace(" ", "") + " - " + DatosRow[20].Replace(" ", "");
                label16.Text = DatosRow[21].Replace(" ", "");
                label17.Text = DatosRow[5].Replace(" ", "");
                label18.Text = DatosRow[13].Replace(" ", "");
                label19.Text = DatosRow[4].Replace(" ", "");
                label20.Text = DatosRow[14].Replace(" ", "");
                label21.Text = DatosRow[7].Replace(" ", "");
                label22.Text = DatosRow[15].Replace(" ", "") + " (Seg)" + " - " + DatosRow[16].Replace(" ", "") + " (Min)";
                label23.Text = DatosRow[17].Replace(" ", "") + " (Seg)";
                label24.Text = DatosRow[42].Replace(" ", "");
                label25.Text = DatosRow[41].Replace(" ", "");
                label26.Text = DatosRow[19].Replace(" ", "") + " - " + DatosRow[13].Replace(" ", "");
                label27.Text = DatosRow[18].Replace(" ", "");
                label35.Text = DatosRow[23].Replace(" ", "");
                label36.Text = DatosRow[24].Replace(" ", "");
                label37.Text = DatosRow[25].Replace(" ", "");
                label38.Text = DatosRow[26].Replace(" ", "");
                label39.Text = DatosRow[27].Replace(" ", "");
                label47.Text = DatosRow[28].Replace(" ", "");
                label48.Text = DatosRow[29].Replace(" ", "");
                label49.Text = DatosRow[30].Replace(" ", "");
                label50.Text = DatosRow[31].Replace(" ", "");
                label51.Text = DatosRow[32].Replace(" ", "");
                label53.Text = DatosRow[16].Replace(" ", "");
                label56.Text = DatosRow[33].Replace(" ", "");
                label58.Text = DatosRow[49].Replace(" ", "");
                label59.Text = DatosRow[34].Replace(" ", "") + "%";
                label60.Text = DatosRow[35].Replace(" ", "");
                label65.Text = DatosRow[38].Replace(" ", "");
                label67.Text = DatosRow[39].Replace(" ", "");
                label41.Text = DatosRow[44].Replace(" ", "");
                label42.Text = DatosRow[45].Replace(" ", "");
                label43.Text = DatosRow[46].Replace(" ", "");
                label44.Text = DatosRow[47].Replace(" ", "");
                label45.Text = DatosRow[48].Replace(" ", "");
                label54.Text = DatosRow[49].Replace(" ", "");
                label62.Text = DatosRow[36].Replace(" ", "");
                label64.Text = DatosRow[50].Replace(" ", "") + "%";

                if (DatosRow[51].Equals("-"))
                {
                    label71.Text = "Aceptada";
                    label71.BackColor = System.Drawing.Color.FromArgb(146, 208, 80);
                    label70.Text = "";
                }
                else
                {
                    label71.Text = "Rechazada";
                    label71.BackColor = System.Drawing.Color.FromArgb(180, 68, 40);
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "Select * From errores where idErrores = ?iderr";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?iderr", DatosRow[51]);

                        lee = comando.ExecuteReader();
                        lee.Read();

                        label70.Text = lee["MensajeError"].ToString();

                        Conexion.Close();
                    }
                }

            }

        }

        #endregion

        #region Timer

        public void IniciaTimer()
        {
            aTimer.Interval = 2000;
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;
        }

        public void IniciaPrograma()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "SELECT MAX(idLlamadasTelefonicas) FROM llamadas_telefonicas";
                    comando = new MySqlCommand(query, Conexion);
                    if (comando.ExecuteScalar() != DBNull.Value)
                    {
                        TramasNuevasDesp = Convert.ToInt32(comando.ExecuteScalar());
                        TramasNuevasAnt = TramasNuevasDesp;
                        aTimer = new System.Timers.Timer();
                        IniciaTimer();
                    }
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("No se pudo cargar la información de la base de datos\n\n" + e.ToString());
            }
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            try
            {
                label69.BackColor = System.Drawing.Color.FromArgb(146, 208, 80);
                label69.Invoke(new Action(() => { label69.Text = "SI"; }));
                aTimer.Enabled = false;
                CargaNuevasTramas();
                aTimer.Enabled = true;
            }
            catch
            {
                label69.BackColor = System.Drawing.Color.Red;
                label69.Invoke(new Action(() => { label69.Text = "NO"; }));
            }

        }

        MySqlConnection ConexionPanelP;
        MySqlCommand ComandoPanelP;
        MySqlDataReader LeePanelP;
        public void CargaNuevasTramas()
        {
            try {
                using (ConexionPanelP = new MySqlConnection(conexion))
                {
                    ConexionPanelP.Open();
                    query = "SELECT MAX(idLlamadasTelefonicas) FROM llamadas_telefonicas";
                    ComandoPanelP = new MySqlCommand(query, ConexionPanelP);
                    if (ComandoPanelP.ExecuteScalar() != DBNull.Value)
                    {
                        TramasNuevasAnt = TramasNuevasDesp;
                        TramasNuevasDesp = Convert.ToInt32(ComandoPanelP.ExecuteScalar());

                    }
                    ConexionPanelP.Close();
                }
                TramasNuevas = TramasNuevasDesp - TramasNuevasAnt;
                if (TramasNuevas != 0)
                {
                    TramasNuevasAnt += 1;
                    using (ConexionPanelP = new MySqlConnection(conexion))
                    {
                        while (TramasNuevasAnt <= TramasNuevasDesp)
                        {
                            ConexionPanelP.Open();
                            query = "Select * From llamadas_telefonicas where idLlamadasTelefonicas = ?ID";
                            ComandoPanelP = new MySqlCommand(query, ConexionPanelP);
                            ComandoPanelP.Parameters.AddWithValue("?ID", TramasNuevasAnt);
                            LeePanelP = ComandoPanelP.ExecuteReader();
                            LeePanelP.Read();

                            row = Dtable.NewRow();
                            row["idLlamadasTelefonicas"] = LeePanelP["idLlamadasTelefonicas"].ToString();
                            row["FFechaFinalLlamada"] = LeePanelP["FFechaFinalLlamada"].ToString();
                            row["HHoraFinalLlamada"] = LeePanelP["HHoraFinalLlamada"].ToString();
                            row["EExtension"] = LeePanelP["EExtension"].ToString();
                            row["TTroncal"] = LeePanelP["TTroncal"].ToString();
                            row["NNumeroMarcado"] = LeePanelP["NNumeroMarcado"].ToString();
                            row["DDuracion"] = LeePanelP["DDuracion"].ToString();
                            row["PCodigoPersonal"] = LeePanelP["PCodigoPersonal"].ToString();
                            row["mFechaInicialLlamada"] = LeePanelP["mFechaInicialLlamada"].ToString();
                            row["jHoraInicialLlamada"] = LeePanelP["jHoraInicialLlamada"].ToString();
                            row["lTipoLlamada"] = LeePanelP["lTipoLlamada"].ToString();
                            row["RTraficoInternoExterno"] = LeePanelP["RTraficoInternoExterno"].ToString();
                            row["ClaseLlamada"] = LeePanelP["ClaseLlamada"].ToString();
                            row["Destino"] = LeePanelP["Destino"].ToString();
                            row["CentroDeCosto"] = LeePanelP["CentroDeCosto"].ToString();
                            row["DuracionLlamada"] = LeePanelP["DuracionLlamada"].ToString();
                            row["DuracionLlamadaAproximada"] = LeePanelP["DuracionLlamadaAproximada"].ToString();
                            row["DurMinima"] = LeePanelP["DurMinima"].ToString();
                            row["PlanTarifa"] = LeePanelP["PlanTarifa"].ToString();
                            row["NumeroTarifa"] = LeePanelP["NumeroTarifa"].ToString();
                            row["NumeFolio"] = LeePanelP["NumeFolio"].ToString();
                            row["ClaseExtension"] = LeePanelP["ClaseExtension"].ToString();
                            row["NombreExtension"] = LeePanelP["NombreExtension"].ToString();
                            row["DuracionRango1"] = LeePanelP["DuracionRango1"].ToString();
                            row["DuracionRango2"] = LeePanelP["DuracionRango2"].ToString();
                            row["DuracionRango3"] = LeePanelP["DuracionRango3"].ToString();
                            row["DuracionRango4"] = LeePanelP["DuracionRango4"].ToString();
                            row["DuracionRango5"] = LeePanelP["DuracionRango5"].ToString();
                            row["ValorRango1"] = LeePanelP["ValorRango1"].ToString();
                            row["ValorRango2"] = LeePanelP["ValorRango2"].ToString();
                            row["ValorRango3"] = LeePanelP["ValorRango3"].ToString();
                            row["ValorRango4"] = LeePanelP["ValorRango4"].ToString();
                            row["ValorRango5"] = LeePanelP["ValorRango5"].ToString();
                            row["CargoFijo"] = LeePanelP["CargoFijo"].ToString();
                            row["RecargoServicioPorcentaje"] = LeePanelP["RecargoServicioPorcentaje"].ToString();
                            row["RecargoServicioValor"] = LeePanelP["RecargoServicioValor"].ToString();
                            row["Base_IVA"] = LeePanelP["Base_IVA"].ToString();
                            row["IVA_Incluye_Recago"] = LeePanelP["IVA_Incluye_Recago"].ToString();
                            row["ValorIVA"] = LeePanelP["ValorIVA"].ToString();
                            row["ValorTotal"] = LeePanelP["ValorTotal"].ToString();
                            row["TramaCompleta"] = LeePanelP["TramaCompleta"].ToString();
                            row["Operador"] = LeePanelP["Operador"].ToString();
                            row["DigitosMinimos"] = LeePanelP["DigitosMinimos"].ToString();
                            row["NombreTarifa"] = LeePanelP["NombreTarifa"].ToString();
                            row["Tarifa1"] = LeePanelP["Tarifa1"].ToString();
                            row["Tarifa2"] = LeePanelP["Tarifa2"].ToString();
                            row["Tarifa3"] = LeePanelP["Tarifa3"].ToString();
                            row["Tarifa4"] = LeePanelP["Tarifa4"].ToString();
                            row["Tarifa5"] = LeePanelP["Tarifa5"].ToString();
                            row["ValorLlamadaTarifa"] = LeePanelP["ValorLlamadaTarifa"].ToString();
                            row["PorcentajeIVA"] = LeePanelP["PorcentajeIVA"].ToString();
                            row["Errores"] = LeePanelP["Errores"].ToString();

                            Dtable.Rows.Add(row);
                            dataGridView1.DataSource = Dtable;
                            dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.Black;

                            if (row["Errores"].Equals("-"))
                            {

                                if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("CEL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("DDI")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("TOL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("LOC")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("DDN")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("ITH")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }
                                else if (row["ClaseExtension"].Equals("H") && row["ClaseLlamada"].Equals("SAT")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(169, 208, 142); }

                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("DDN")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("DDI")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("CEL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("TOL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("LOC")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("ITH")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }
                                else if (row["ClaseExtension"].Equals("A") && row["ClaseLlamada"].Equals("SAT")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204); }

                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("DDN")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("DDI")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("CEL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("TOL")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("LOC")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("ITH")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }
                                else if (row["ClaseExtension"].Equals("S") && row["ClaseLlamada"].Equals("SAT")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 255, 93); }

                                else if (row["ClaseLlamada"].Equals("INT")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 153, 204); }
                                else if (row["ClaseLlamada"].Equals("INV")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(191, 191, 191); }
                                else if (row["ClaseLlamada"].Equals("EXC")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 192, 0); }
                                else if (row["ClaseLlamada"].Equals("ENT")) { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(221, 235, 247); }

                                else { dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(244, 176, 132); }

                            }
                            else
                            {
                                dataGridView1.Rows[dataGridView1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(191, 191, 191);
                            }

                            LeePanelP.Close();
                            TramasNuevasAnt++;
                            dataGridView1.Invoke(new Action(() => { dataGridView1.Refresh(); }));
                            ConexionPanelP.Close();
                        }
                        dataGridView1.Invoke(new Action(() => { dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1; }));
                    }
                }
            }
            catch
            {
                label69.BackColor = System.Drawing.Color.Red;
                label69.Invoke(new Action(() => { label69.Text = "NO"; }));
            }
        }

        #endregion

        #endregion

        #region Planes

        DataTable Dtable2;
        int Pos = 0;
        int CantidadCeldas;
        int CurrentRow = 0;
        int Tabpage = 0;
        List<string> DatosRow2 = new List<string>();
        List<string> AuxDatosRow2 = new List<string>();
        List<string> SalvaPL = new List<string>();
        bool Iguales;

        #region PasoDeComponentes

        private void tabControl2_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (DatosIgualesPlanes() == false)
            {
                if (GuardarCambios() == true)
                {
                    GuardaCambiosPlanes();
                    PasaComponentes();
                    Tabpage = e.TabPageIndex;
                    dataGridView2_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, 0));
                }
                else
                {
                    PasaComponentes();
                    Tabpage = e.TabPageIndex;
                    dataGridView2_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, 0));
                }
            }
            else
            {
                PasaComponentes();
                Tabpage = e.TabPageIndex;
                dataGridView2_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, 0));
            }
        }

        public void PasaComponentes()
        {
            DatosRow2 = new List<string>();
            if (tabControl2.SelectedIndex == 0)
            {
                CargaTablas("0");
                tabPage8.Controls.Add(dataGridView2);
                tabPage8.Controls.Add(groupBox3);
                tabPage8.Controls.Add(groupBox4);
                tabPage8.Controls.Add(tableLayoutPanel1);
            }
            else if (tabControl2.SelectedIndex == 1)
            {
                CargaTablas("1");
                tabPage9.Controls.Add(dataGridView2);
                tabPage9.Controls.Add(groupBox3);
                tabPage9.Controls.Add(groupBox4);
                tabPage9.Controls.Add(tableLayoutPanel1);
            }
            else if (tabControl2.SelectedIndex == 2)
            {
                CargaTablas("2");
                tabPage10.Controls.Add(dataGridView2);
                tabPage10.Controls.Add(groupBox3);
                tabPage10.Controls.Add(groupBox4);
                tabPage10.Controls.Add(tableLayoutPanel1);
            }
        }

        #endregion

        #region CargaTablas

        public void CargaTablas(string Tabla)
        {
            dataGridView2.DataSource = null;
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    if (Tabla.Equals("0")) { query = "Select * From plan_tarifario_001"; }
                    else if (Tabla.Equals("1")) { query = "Select * From plan_tarifario_002"; }
                    else if (Tabla.Equals("2")) { query = "Select * From plan_tarifario_003"; }

                    using (adapter = new MySqlDataAdapter(query, Conexion))
                    {
                        Dtable2 = new DataTable();
                        adapter.Fill(Dtable2);
                        dataGridView2.DataSource = Dtable2;
                        adapter.Dispose();
                    }

                    Conexion.Close();
                }
                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.RowHeadersVisible = false;
                dataGridView2.EnableHeadersVisualStyles = false;
                dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(198, 224, 180);

                Pos = 0;
                while (Pos < dataGridView2.ColumnCount)
                {
                    if (Pos != 0)
                    {
                        dataGridView2.Columns[Pos].Visible = false;
                    }
                    else
                    {
                        dataGridView2.Columns[Pos].HeaderText = "Tarifa";
                        dataGridView2.Columns[Pos].Width = 75;
                        dataGridView2.Columns[Pos].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[Pos].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 242, 204);
                    }
                    Pos++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!, más información: \n\n" + e.ToString());
            }
        }

        #endregion

        #region SeleccionaCell

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (dataGridView2.RowCount > 0)
                {
                    if (DatosRow2.Count > 0)
                    {
                        if (DatosIgualesPlanes() == false)
                        {
                            if (GuardarCambios() == true)
                            {
                                GuardaCambiosPlanes();
                                CargaDatos(e);
                            }
                            else
                            {
                                CargaDatos(e);
                            }
                        }
                        else
                        {
                            CargaDatos(e);
                        }
                    }
                    else
                    {
                        CargaDatos(e);
                    }
                }
                else
                {
                    MessageBox.Show("No hay datos en la tabla!");
                }
            }
        }
        public void CargaDatos(DataGridViewCellEventArgs e)
        {
            if (e == null) { dataGridView2.Rows[0].Selected = true; CurrentRow = 0; } else { dataGridView2.Rows[e.RowIndex].Selected = true; CurrentRow = e.RowIndex; }
            CantidadCeldas = dataGridView2.ColumnCount;
            Posicion = 0;
            DatosRow2 = new List<string>();
            while (Posicion < CantidadCeldas)
            {
                DatosRow2.Add(dataGridView2.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            EscribeDatos();
        }

        public void EscribeDatos()
        {
            label76.Text = DatosRow2[1];
            textBox21.Text = DatosRow2[2];

            textBox22.Text = DatosRow2[3];
            textBox23.Text = DatosRow2[4];
            textBox24.Text = DatosRow2[25];
            textBox25.Text = DatosRow2[26];
            comboBox11.Text = DatosRow2[27];

            //1
            label121.Text = DatosRow2[2];
            textBox1.Text = DatosRow2[5];
            textBox2.Text = DatosRow2[10];

            //2
            label122.Text = (DatosRow2[5].Split(':')[1]); if (Convert.ToInt32(label122.Text) + 1 >= 10) { label122.Text = DatosRow2[5].Split(':')[0] + ":" + (Convert.ToInt32(label122.Text) + 1).ToString(); } else { label122.Text = DatosRow2[5].Split(':')[0] + ":0" + (Convert.ToInt32(label122.Text) + 1); }
            textBox3.Text = DatosRow2[6];
            textBox4.Text = DatosRow2[11];
            textBox5.Text = DatosRow2[12];

            //3
            label123.Text = (DatosRow2[6].Split(':')[1]); if (Convert.ToInt32(label123.Text) + 1 >= 10) { label123.Text = DatosRow2[6].Split(':')[0] + ":" + (Convert.ToInt32(label123.Text) + 1).ToString(); } else { label123.Text = DatosRow2[6].Split(':')[0] + ":0" + (Convert.ToInt32(label123.Text) + 1); }
            textBox6.Text = DatosRow2[7];
            textBox7.Text = DatosRow2[13];
            textBox8.Text = DatosRow2[14];
            textBox9.Text = DatosRow2[15];

            //4
            label124.Text = (DatosRow2[7].Split(':')[1]); if (Convert.ToInt32(label124.Text) + 1 >= 10) { label124.Text = DatosRow2[7].Split(':')[0] + ":" + (Convert.ToInt32(label124.Text) + 1).ToString(); } else { label124.Text = DatosRow2[7].Split(':')[0] + ":0" + (Convert.ToInt32(label124.Text) + 1); }
            textBox10.Text = DatosRow2[8];
            textBox11.Text = DatosRow2[16];
            textBox12.Text = DatosRow2[17];
            textBox13.Text = DatosRow2[18];
            textBox14.Text = DatosRow2[19];

            //5
            label125.Text = (DatosRow2[8].Split(':')[1]); if (Convert.ToInt32(label125.Text) + 1 >= 10) { label125.Text = DatosRow2[8].Split(':')[0] + ":" + (Convert.ToInt32(label125.Text) + 1).ToString(); } else { label125.Text = DatosRow2[8].Split(':')[0] + ":0" + (Convert.ToInt32(label125.Text) + 1); }
            textBox15.Text = DatosRow2[9];
            textBox16.Text = DatosRow2[20];
            textBox17.Text = DatosRow2[21];
            textBox18.Text = DatosRow2[22];
            textBox19.Text = DatosRow2[23];
            textBox20.Text = DatosRow2[24];

            //6
            label120.Text = (DatosRow2[9].Split(':')[1]); if (Convert.ToInt32(label120.Text) + 1 >= 10) { label120.Text = DatosRow2[9].Split(':')[0] + ":" + (Convert.ToInt32(label120.Text) + 1).ToString(); } else { label120.Text = DatosRow2[9].Split(':')[0] + ":0" + (Convert.ToInt32(label120.Text) + 1); }
            label119.Text = "o Más";
            label118.Text = DatosRow2[20];
            label117.Text = DatosRow2[21];
            label116.Text = DatosRow2[22];
            label115.Text = DatosRow2[23];
            label114.Text = DatosRow2[24];
        }

        #endregion

        #region Botones

        private void button2_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Al salir se perderá la configuración no guardada, ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                MessageBox.Show("La configuración no se guardará");
                tabControl1.SelectTab(0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DatosIgualesPlanes() == false)
            {
                GuardaCambiosPlanes();
                dataGridView2.ClearSelection();
                dataGridView2_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, CurrentRow));
            }
            else
            {
                MessageBox.Show("No se han detectado cambios");
            }

        }

        #endregion

        #region Guarda Cambios

        public bool GuardarCambios()
        {
            DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("No se han guardado los cambios");
                return (false);
            }
            else
            {
                return (true);
            }
        }

        public void GuardaCambiosPlanes()
        {
            if (DatosCorrectosPlanes() == true)
            {
                ActualizaTabla();
                if (Tabpage == 0) { CargaTablas("0"); }
                else if (Tabpage == 1) { CargaTablas("1"); }
                else if (Tabpage == 2) { CargaTablas("2"); }
                dataGridView2.ClearSelection();
            }
            else
            {
                MessageBox.Show("Uno o más datos no se han ingresado en el formato correcto!");
            }
        }

        public void CargaDatosNuevosPlanes()
        {
            DatosRow2 = new List<string>();
            DatosRow2.Add(dataGridView2.Rows[CurrentRow].Cells[0].Value.ToString());
            DatosRow2.Add(label76.Text);
            DatosRow2.Add(textBox21.Text);
            DatosRow2.Add(textBox22.Text);
            DatosRow2.Add(textBox23.Text);
            DatosRow2.Add(textBox1.Text);
            DatosRow2.Add(textBox3.Text);
            DatosRow2.Add(textBox6.Text);
            DatosRow2.Add(textBox10.Text);
            DatosRow2.Add(textBox15.Text);
            DatosRow2.Add(textBox2.Text);
            DatosRow2.Add(textBox4.Text);
            DatosRow2.Add(textBox5.Text);
            DatosRow2.Add(textBox7.Text);
            DatosRow2.Add(textBox8.Text);
            DatosRow2.Add(textBox9.Text);
            DatosRow2.Add(textBox11.Text);
            DatosRow2.Add(textBox12.Text);
            DatosRow2.Add(textBox13.Text);
            DatosRow2.Add(textBox14.Text);
            DatosRow2.Add(textBox16.Text);
            DatosRow2.Add(textBox17.Text);
            DatosRow2.Add(textBox18.Text);
            DatosRow2.Add(textBox19.Text);
            DatosRow2.Add(textBox20.Text);
            DatosRow2.Add(textBox24.Text);
            DatosRow2.Add(textBox25.Text);
            DatosRow2.Add(comboBox11.Text);
        }


        #endregion

        #region Datos Iguales y Correctos

        public bool DatosIgualesPlanes()
        {
            AuxDatosRow2 = DatosRow2;
            CargaDatosNuevosPlanes();

            Iguales = true;
            if (Tabpage == 0) { query = "Select * From plan_tarifario_001 where Codi_Tarifa = ?CT"; }
            else if (Tabpage == 1) { query = "Select * From plan_tarifario_002 where Codi_Tarifa = ?CT"; }
            else if (Tabpage == 2) { query = "Select * From plan_tarifario_003 where Codi_Tarifa = ?CT"; }

            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", dataGridView2.Rows[CurrentRow].Cells[0].Value.ToString());

                    lee = comando.ExecuteReader();
                    lee.Read();

                    CantidadCeldas = dataGridView2.ColumnCount;
                    Posicion = 0;

                    while (Posicion < CantidadCeldas)
                    {
                        if (!dataGridView2.Rows[CurrentRow].Cells[Posicion].Value.ToString().Equals(DatosRow2[Posicion]))
                        {
                            Iguales = false;
                        }
                        Posicion++;
                    }

                    lee.Close();
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Problema al conectarse con la base de datos o buscar el código de tarifa seleccionado, más información: \n\n" + e.ToString());
            }

            if (Iguales == true)
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        public bool DatosCorrectosPlanes()
        {
            try
            {
                if ((textBox1.Text.Split(':')[0].Length == 2 && textBox1.Text.Split(':')[1].Length == 2) &&
                    (textBox3.Text.Split(':')[0].Length == 2 && textBox3.Text.Split(':')[1].Length == 2) &&
                    (textBox6.Text.Split(':')[0].Length == 2 && textBox6.Text.Split(':')[1].Length == 2) &&
                    (textBox10.Text.Split(':')[0].Length == 2 && textBox10.Text.Split(':')[1].Length == 2) &&
                    (textBox15.Text.Split(':')[0].Length == 2 && textBox15.Text.Split(':')[1].Length == 2) &&
                    (textBox21.Text.Split(':')[0].Length == 2 && textBox21.Text.Split(':')[1].Length == 2) &&
                    (textBox22.Text.Split(':')[0].Length == 2 && textBox22.Text.Split(':')[1].Length == 2) &&
                    (textBox2.Text.Split('.')[1].Length == 2) && (textBox4.Text.Split('.')[1].Length == 2) && (textBox5.Text.Split('.')[1].Length == 2) &&
                    (textBox7.Text.Split('.')[1].Length == 2) && (textBox8.Text.Split('.')[1].Length == 2) && (textBox9.Text.Split('.')[1].Length == 2) &&
                    (textBox11.Text.Split('.')[1].Length == 2) && (textBox12.Text.Split('.')[1].Length == 2) && (textBox13.Text.Split('.')[1].Length == 2) &&
                    (textBox14.Text.Split('.')[1].Length == 2) && (textBox16.Text.Split('.')[1].Length == 2) && (textBox23.Text.Split('.')[1].Length == 2) &&
                    (textBox17.Text.Split('.')[1].Length == 2) && (textBox18.Text.Split('.')[1].Length == 2) && (textBox19.Text.Split('.')[1].Length == 2) &&
                    (textBox20.Text.Split('.')[1].Length == 2) && (textBox24.Text.Split('.')[1].Length == 2) && (textBox25.Text.Split('.')[1].Length == 2) &&
                    (comboBox11.Text.Length == 1))
                {
                    return (true);
                }
                else
                {
                    return (false);
                }
            }
            catch
            {
                return (false);
            }
        }


        #endregion

        #region Actualiza Tabla


        public void ActualizaTabla()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRow2.Count; i++)
                    {
                        SalvaPL.Add(DatosRow2[i]);
                    }

                    if (Tabpage == 0) { query = "delete from plan_tarifario_001 where Codi_Tarifa =?CT"; }
                    else if (Tabpage == 1) { query = "delete from plan_tarifario_002 where Codi_Tarifa =?CT"; }
                    else if (Tabpage == 2) { query = "delete from plan_tarifario_003 where Codi_Tarifa =?CT"; }
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRow2[0]);
                    comando.ExecuteNonQuery();

                    if (Tabpage == 0) { query = @"insert into plan_tarifario_001 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }
                    else if (Tabpage == 1) { query = @"insert into plan_tarifario_002 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }
                    else if (Tabpage == 2) { query = @"insert into plan_tarifario_003 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }

                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", Convert.ToInt32(DatosRow2[0]));
                    comando.Parameters.AddWithValue("?NT", DatosRow2[1]);
                    comando.Parameters.AddWithValue("?DM", DatosRow2[2]);
                    comando.Parameters.AddWithValue("?IT", DatosRow2[3]);
                    comando.Parameters.AddWithValue("?CF", DatosRow2[4]);
                    comando.Parameters.AddWithValue("?TR1", DatosRow2[5]);
                    comando.Parameters.AddWithValue("?TR2", DatosRow2[6]);
                    comando.Parameters.AddWithValue("?TR3", DatosRow2[7]);
                    comando.Parameters.AddWithValue("?TR4", DatosRow2[8]);
                    comando.Parameters.AddWithValue("?TR5", DatosRow2[9]);
                    comando.Parameters.AddWithValue("?VT11", DatosRow2[10]);
                    comando.Parameters.AddWithValue("?VT21", DatosRow2[11]);
                    comando.Parameters.AddWithValue("?VT22", DatosRow2[12]);
                    comando.Parameters.AddWithValue("?VT31", DatosRow2[13]);
                    comando.Parameters.AddWithValue("?VT32", DatosRow2[14]);
                    comando.Parameters.AddWithValue("?VT33", DatosRow2[15]);
                    comando.Parameters.AddWithValue("?VT41", DatosRow2[16]);
                    comando.Parameters.AddWithValue("?VT42", DatosRow2[17]);
                    comando.Parameters.AddWithValue("?VT43", DatosRow2[18]);
                    comando.Parameters.AddWithValue("?VT44", DatosRow2[19]);
                    comando.Parameters.AddWithValue("?VT51", DatosRow2[20]);
                    comando.Parameters.AddWithValue("?VT52", DatosRow2[21]);
                    comando.Parameters.AddWithValue("?VT53", DatosRow2[22]);
                    comando.Parameters.AddWithValue("?VT54", DatosRow2[23]);
                    comando.Parameters.AddWithValue("?VT55", DatosRow2[24]);
                    comando.Parameters.AddWithValue("?RSP", DatosRow2[25]);
                    comando.Parameters.AddWithValue("?PIVA", DatosRow2[26]);
                    comando.Parameters.AddWithValue("?IVAIR", DatosRow2[27]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }
                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        if (Tabpage == 0) { query = @"insert into plan_tarifario_001 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }
                        else if (Tabpage == 1) { query = @"insert into plan_tarifario_002 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }
                        else if (Tabpage == 2) { query = @"insert into plan_tarifario_003 (Codi_Tarifa, Nomb_Tarifa, Dura_Minima, Inte_Cobro, Cargo_Fijo, Rango_Tar_1, Rango_Tar_2,
                            Rango_Tar_3, Rango_Tar_4, Rango_Tar_5, Rango_Tar_1_Vr_1, Rango_Tar_2_Vr_1, Rango_Tar_2_Vr_2, Rango_Tar_3_Vr_1, Rango_Tar_3_Vr_2,
                            Rango_Tar_3_Vr_3, Rango_Tar_4_Vr_1, Rango_Tar_4_Vr_2, Rango_Tar_4_Vr_3, Rango_Tar_4_Vr_4, Rango_Tar_5_Vr_1, Rango_Tar_5_Vr_2,
                            Rango_Tar_5_Vr_3, Rango_Tar_5_Vr_4, Rango_Tar_5_Vr_5, Reca_Servicio_Porcentual, Porc_IVA, IVA_Incluye_Recago) 
                            values (?CT, ?NT, ?DM, ?IT, ?CF, ?TR1, ?TR2, ?TR3, ?TR4, ?TR5, ?VT11, ?VT21, ?VT22, ?VT31, ?VT32, ?VT33, ?VT41, ?VT42, ?VT43,
                            ?VT44, ?VT51, ?VT52, ?VT53, ?VT54, ?VT55, ?RSP, ?PIVA, ?IVAIR)"; }

                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", Convert.ToInt32(SalvaPL[0]));
                        comando.Parameters.AddWithValue("?NT", SalvaPL[1]);
                        comando.Parameters.AddWithValue("?DM", SalvaPL[2]);
                        comando.Parameters.AddWithValue("?IT", SalvaPL[3]);
                        comando.Parameters.AddWithValue("?CF", SalvaPL[4]);
                        comando.Parameters.AddWithValue("?TR1", SalvaPL[5]);
                        comando.Parameters.AddWithValue("?TR2", SalvaPL[6]);
                        comando.Parameters.AddWithValue("?TR3", SalvaPL[7]);
                        comando.Parameters.AddWithValue("?TR4", SalvaPL[8]);
                        comando.Parameters.AddWithValue("?TR5", SalvaPL[9]);
                        comando.Parameters.AddWithValue("?VT11", SalvaPL[10]);
                        comando.Parameters.AddWithValue("?VT21", SalvaPL[11]);
                        comando.Parameters.AddWithValue("?VT22", SalvaPL[12]);
                        comando.Parameters.AddWithValue("?VT31", SalvaPL[13]);
                        comando.Parameters.AddWithValue("?VT32", SalvaPL[14]);
                        comando.Parameters.AddWithValue("?VT33", SalvaPL[15]);
                        comando.Parameters.AddWithValue("?VT41", SalvaPL[16]);
                        comando.Parameters.AddWithValue("?VT42", SalvaPL[17]);
                        comando.Parameters.AddWithValue("?VT43", SalvaPL[18]);
                        comando.Parameters.AddWithValue("?VT44", SalvaPL[19]);
                        comando.Parameters.AddWithValue("?VT51", SalvaPL[20]);
                        comando.Parameters.AddWithValue("?VT52", SalvaPL[21]);
                        comando.Parameters.AddWithValue("?VT53", SalvaPL[22]);
                        comando.Parameters.AddWithValue("?VT54", SalvaPL[23]);
                        comando.Parameters.AddWithValue("?VT55", SalvaPL[24]);
                        comando.Parameters.AddWithValue("?RSP", SalvaPL[25]);
                        comando.Parameters.AddWithValue("?PIVA", SalvaPL[26]);
                        comando.Parameters.AddWithValue("?IVAIR", SalvaPL[27]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                        MessageBox.Show("Los cambios se han revertido");
                    }
                }
                catch
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está correctamente conectado a la base de datos?");
                }
            }
        }

        #endregion

        #endregion

        #region Indicativos

        DataTable Dtable3;
        int CantidadCeldas2;
        int CeldaAnt = 0;
        int TabPage2 = 0;
        bool Agrega;
        List<string> DatosRow3 = new List<string>();
        List<string> ClaseLlamada = new List<string>();
        List<int> Tarifa = new List<int>();
        List<string> Digitos_Minimos = new List<string>();
        List<string> SalvaIN = new List<string>();

        #region CargaTabla

        public void CargaCombo(string Indic)
        {
            if (comboBox1.Items.Count == 0 && comboBox2.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from clase_llamadas";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            ClaseLlamada.Add(lee["Clase_Llamada"].ToString());
                        }
                        lee.Close();
                        query = "select * from plan_tarifario_001";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            Tarifa.Add(Convert.ToInt32(lee["Codi_Tarifa"]));
                        }
                        lee.Close();

                    }

                    Conexion.Close();
                    Tarifa.Sort();
                    ClaseLlamada.Sort();
                    Digitos_Minimos.Sort();

                    for (int i = 0; i < ClaseLlamada.Count; i++)
                    {
                        comboBox1.Items.Add(ClaseLlamada[i]);
                    }
                    for (int i = 0; i < Tarifa.Count; i++)
                    {
                        comboBox2.Items.Add(Tarifa[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        public void CargaInd(string Tabla)
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    if (Tabla.Equals("0")) { query = "Select * From indicativosbg"; }
                    else if (Tabla.Equals("1")) { query = "Select * From indicativoscf"; }
                    else if (Tabla.Equals("2")) { query = "Select * From indicativosip"; }
                    else if (Tabla.Equals("3")) { query = "Select * From indicativosit"; }

                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        Dtable3 = new DataTable();
                        adapter.Fill(Dtable3);
                        dataGridView3.DataSource = Dtable3;
                    }
                    dataGridView3.EnableHeadersVisualStyles = false;
                    dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView3.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dataGridView3.Columns[0].Width = 150;
                    dataGridView3.Columns[1].Width = 250;
                    dataGridView3.Columns[2].Width = 132;
                    dataGridView3.Columns[3].Width = 132;
                    dataGridView3.Columns[4].Width = 132;

                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        #endregion

        #region PasoComponentes

        private void tabControl3_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (DatosIgualesIndic() == true)
            {
                PasaComp(e);
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    e.Cancel = true;
                    MessageBox.Show("Los cambios se descartaron");
                    CargaFila(new DataGridViewCellEventArgs(0, CeldaAnt));
                }
                else
                {
                    GuardaCambiosIndic();
                    PasaComp(e);
                }
            }
        }

        public void PasaComp(TabControlCancelEventArgs e)
        {
            if (tabControl3.SelectedIndex == 0)
            {
                CargaInd("0");
                TabPage2 = e.TabPageIndex;
                tabPage11.Controls.Add(dataGridView3);
                tabPage11.Controls.Add(groupBox5);
                tabPage11.Controls.Add(button3);
                tabPage11.Controls.Add(button4);
                tabPage11.Controls.Add(button47);
                tabPage11.Controls.Add(button48);
            }
            else if (tabControl3.SelectedIndex == 1)
            {
                CargaInd("1");
                TabPage2 = e.TabPageIndex;
                tabPage12.Controls.Add(dataGridView3);
                tabPage12.Controls.Add(groupBox5);
                tabPage12.Controls.Add(button3);
                tabPage12.Controls.Add(button4);
                tabPage12.Controls.Add(button47);
                tabPage12.Controls.Add(button48);
            }
            else if (tabControl3.SelectedIndex == 2)
            {
                CargaInd("2");
                Tabpage = e.TabPageIndex;
                tabPage13.Controls.Add(dataGridView3);
                tabPage13.Controls.Add(groupBox5);
                tabPage13.Controls.Add(button3);
                tabPage13.Controls.Add(button4);
                tabPage13.Controls.Add(button47);
                tabPage13.Controls.Add(button48);
            }
            else if (tabControl3.SelectedIndex == 3)
            {
                CargaInd("3");
                TabPage2 = e.TabPageIndex;
                tabPage14.Controls.Add(dataGridView3);
                tabPage14.Controls.Add(groupBox5);
                tabPage14.Controls.Add(button3);
                tabPage14.Controls.Add(button4);
                tabPage14.Controls.Add(button47);
                tabPage14.Controls.Add(button48);
            }
            dataGridView3.ClearSelection();
            dataGridView3_CellClick(dataGridView3, new DataGridViewCellEventArgs(0, 0));
        }

        #endregion

        #region SeleccionaCelda

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (dataGridView3.RowCount > 0)
                {
                    if (DatosRow3.Count > 0)
                    {
                        if (DatosIgualesIndic() == true)
                        {
                            CargaFila(e);
                        }
                        else
                        {
                            DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.No)
                            {
                                dataGridView3.Rows[CeldaAnt].Selected = true;
                                textBox27.Text = DatosRow3[0];
                                textBox28.Text = DatosRow3[1];
                                comboBox1.Text = DatosRow3[2];
                                comboBox2.Text = DatosRow3[3];
                                textBox26.Text = DatosRow3[4];
                                dataGridView3_CellClick(null, new DataGridViewCellEventArgs(0, CeldaAnt));
                                MessageBox.Show("Los cambios se descartaron");
                            }
                            else
                            {
                                GuardaCambiosIndic();
                                if (TabPage2 == 0) { CargaInd("0"); }
                                else if (TabPage2 == 1) { CargaInd("1"); }
                                else if (TabPage2 == 2) { CargaInd("2"); }
                                else if (TabPage2 == 3) { CargaInd("3"); }
                                CargaFila(e);
                            }
                        }
                    }
                    else
                    {
                        CargaFila(e);
                    }
                }
                else
                {
                    MessageBox.Show("No hay datos en la tabla!");
                }
            }
        }

        public void CargaFila(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView3.ClearSelection();
                dataGridView3.Rows[e.RowIndex].Selected = true;
                CeldaAnt = e.RowIndex;
                CantidadCeldas2 = dataGridView3.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRow3 = new List<string>();
                while (Posicion < CantidadCeldas2)
                {
                    DatosRow3.Add(dataGridView3.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }
                textBox27.Text = DatosRow3[0];
                textBox28.Text = DatosRow3[1];
                comboBox1.Text = DatosRow3[2];
                comboBox2.Text = DatosRow3[3];
                textBox26.Text = DatosRow3[4];
            }
            catch
            {
                dataGridView3.Rows[0].Selected = true;
            }
        }


        #endregion

        #region Botones

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Al salir se perderá la configuración no guardada, ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                MessageBox.Show("La configuración no se guardará");
                tabControl1.SelectTab(0);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (DatosIgualesIndic() == false)
            {
                GuardaCambiosIndic();
                if (TabPage2 == 0) { CargaInd("0"); }
                else if (TabPage2 == 1) { CargaInd("1"); }
                else if (TabPage2 == 2) { CargaInd("2"); }
                else if (TabPage2 == 3) { CargaInd("3"); }
                dataGridView3.ClearSelection();
                dataGridView3_CellClick(dataGridView3, new DataGridViewCellEventArgs(0, CeldaAnt));
            }
            else
            {
                MessageBox.Show("No se han detectado cambios");
            }
        }

        #endregion

        #region Datos Iguales

        public bool DatosIgualesIndic()
        {
            if (textBox27.Text == DatosRow3[0] && textBox28.Text == DatosRow3[1] && comboBox1.Text == DatosRow3[2]
                && comboBox2.Text == DatosRow3[3] && textBox26.Text == DatosRow3[4])
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }


        #endregion

        #region GuardaCambios

        public void GuardaCambiosIndic()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRow3.Count; i++)
                    {
                        SalvaIN.Add(DatosRow3[i]);
                    }
                    if (TabPage2 == 0) { query = "delete from indicativosbg where Indicativo =?CT"; }
                    else if (TabPage2 == 1) { query = "delete from indicativoscf where Indicativo =?CT"; }
                    else if (TabPage2 == 2) { query = "delete from indicativosip where Indicativo =?CT"; }
                    else if (TabPage2 == 3) { query = "delete from indicativosit where Indicativo =?CT"; }
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRow3[0]);
                    comando.ExecuteNonQuery();
                    CargaDatosRow();
                    if (TabPage2 == 0) { query = @"insert into indicativosbg (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                    else if (TabPage2 == 1) { query = @"insert into indicativoscf (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                    else if (TabPage2 == 2) { query = @"insert into indicativosip (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                    else if (TabPage2 == 3) { query = @"insert into indicativosit (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }

                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?Ind", DatosRow3[0]);
                    comando.Parameters.AddWithValue("?Des", DatosRow3[1]);
                    comando.Parameters.AddWithValue("?Cla", DatosRow3[2]);
                    comando.Parameters.AddWithValue("?Tar", DatosRow3[3]);
                    comando.Parameters.AddWithValue("?Dig", DatosRow3[4]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }
                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        if (TabPage2 == 0) { query = @"insert into indicativosbg (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                        else if (TabPage2 == 1) { query = @"insert into indicativoscf (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                        else if (TabPage2 == 2) { query = @"insert into indicativosip (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }
                        else if (TabPage2 == 3) { query = @"insert into indicativosit (Indicativo, Destino, Clase_Llamada, Tarifa, Digitos_Minimos) 
                            values (?Ind, ?Des, ?Cla, ?Tar, ?Dig)"; }

                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?Ind", SalvaIN[0]);
                        comando.Parameters.AddWithValue("?Des", SalvaIN[1]);
                        comando.Parameters.AddWithValue("?Cla", SalvaIN[2]);
                        comando.Parameters.AddWithValue("?Tar", SalvaIN[3]);
                        comando.Parameters.AddWithValue("?Dig", SalvaIN[4]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                }
                catch (Exception s)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está correctamente conectado a la base de datos?" + s.ToString());
                }
            }
        }

        public void CargaDatosRow()
        {
            DatosRow3[0] = textBox27.Text;
            DatosRow3[1] = textBox28.Text;
            DatosRow3[2] = comboBox1.Text;
            DatosRow3[3] = comboBox2.Text;
            DatosRow3[4] = textBox26.Text;
        }


        #endregion

        #region Crea y borra

        private void button47_Click(object sender, EventArgs e)
        {
            DataRow row = Dtable3.NewRow();
            Dtable3.Rows.Add(row);
            dataGridView3.DataSource = Dtable3;
            dataGridView3.Invoke(new Action(() => { dataGridView3.FirstDisplayedScrollingRowIndex = dataGridView3.RowCount - 1; }));
            dataGridView3_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView3.RowCount - 1));
        }

        private void button48_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el indicativo?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        if (TabPage2 == 0) { query = "delete from indicativosbg where Indicativo = ?CT"; }
                        else if (TabPage2 == 1) { query = "delete from indicativoscf where Indicativo = ?CT"; }
                        else if (TabPage2 == 2) { query = "delete from indicativosip where Indicativo = ?CT"; }
                        else if (TabPage2 == 3) { query = "delete from indicativosit where Indicativo = ?CT"; }
                        
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRow3[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    if (TabPage2 == 0) { CargaInd("0"); }
                    else if (TabPage2 == 1) { CargaInd("1"); }
                    else if (TabPage2 == 2) { CargaInd("2"); }
                    else if (TabPage2 == 3) { CargaInd("3"); }
                    MessageBox.Show("el indicativo se ha borrado exitosamente");
                    dataGridView3_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            dataGridView3.SelectedRows[0].Cells[0].Value = textBox27.Text;
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            dataGridView3.SelectedRows[0].Cells[1].Value = textBox28.Text;
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            dataGridView3.SelectedRows[0].Cells[2].Value = comboBox1.Text;
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            dataGridView3.SelectedRows[0].Cells[3].Value = comboBox2.Text;
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            dataGridView3.SelectedRows[0].Cells[4].Value = textBox26.Text;
        }

        #endregion

        #endregion

        #region Tablas

        #region tablas

        private void tabControl4_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabControl4.SelectedIndex != 0 && DatosRowEx.Count > 0) { dataGridView4_CellClick(dataGridView4, new DataGridViewCellEventArgs(0, CeldaEx)); }
            if (tabControl4.SelectedIndex != 1 && DatosRowCC.Count > 0) { dataGridView5_CellClick(dataGridView5, new DataGridViewCellEventArgs(0, CeldaCC)); }
            if (tabControl4.SelectedIndex != 2 && DatosRowTR.Count > 0) { dataGridView6_CellClick(dataGridView6, new DataGridViewCellEventArgs(0, CeldaTR)); }
            if (tabControl4.SelectedIndex != 3 && DatosRowCP.Count > 0) { dataGridView7_CellClick(dataGridView7, new DataGridViewCellEventArgs(0, CeldaCP)); }
            if (tabControl4.SelectedIndex != 4 && DatosRowCE.Count > 0) { dataGridView8_CellClick(dataGridView8, new DataGridViewCellEventArgs(0, CeldaCE)); }
            if (tabControl4.SelectedIndex != 5 && DatosRowCL.Count > 0) { dataGridView9_CellClick(dataGridView9, new DataGridViewCellEventArgs(0, CeldaCL)); }
            if (tabControl4.SelectedIndex != 6 && DatosRowOP.Count > 0) { dataGridView10_CellClick(dataGridView10, new DataGridViewCellEventArgs(0, CeldaOP)); }

            if (e.TabPageIndex == 0) { tabPage15.Controls.Add(button5); CargaTablasEx(); dataGridView4_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 1) { tabPage16.Controls.Add(button5); CargaTablaCC(); dataGridView5_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 2) { tabPage17.Controls.Add(button5); CargaTablaTR(); dataGridView6_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 3) { tabPage18.Controls.Add(button5); CargaTablaCP(); dataGridView7_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 4) { tabPage19.Controls.Add(button5); CargaTablaCE(); dataGridView8_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 5) { tabPage20.Controls.Add(button5); CargaTablaCL(); dataGridView9_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            if (e.TabPageIndex == 6) { tabPage21.Controls.Add(button5); CargaTablaOP(); dataGridView10_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Al salir se decartarán los cambios no guardados, desea continuar?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                tabControl1.SelectTab(0);
            }
        }

        #endregion

        #region Extensiones

        DataTable DtableEx;
        List<string> DatosRowEx = new List<string>();
        List<string> Cl_Extension = new List<string>();
        List<string> Cod_Centro = new List<string>();
        List<string> Env_Diario = new List<string>();
        List<string> SalvaEX = new List<string>();
        int CeldaEx = 0;
        int CantidadCeldasEx;

        public void CargaTablasEx()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From extensiones";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableEx = new DataTable();
                        adapter.Fill(DtableEx);
                        dataGridView4.DataSource = DtableEx;
                    }
                    dataGridView4.EnableHeadersVisualStyles = false;
                    dataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView4.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dataGridView4.Columns[0].HeaderText = "Número";
                    dataGridView4.Columns[1].HeaderText = "Folio";
                    dataGridView4.Columns[2].HeaderText = "Ubicación";
                    dataGridView4.Columns[3].HeaderText = "Responsable";
                    dataGridView4.Columns[4].HeaderText = "CL Extensión";
                    dataGridView4.Columns[5].HeaderText = "Cod Centro";
                    dataGridView4.Columns[6].HeaderText = "E-Mail";
                    dataGridView4.Columns[7].HeaderText = "Envío diario";

                    dataGridView4.Columns[0].Width = 100;
                    dataGridView4.Columns[1].Width = 100;
                    dataGridView4.Columns[2].Width = 200;
                    dataGridView4.Columns[3].Width = 160;
                    dataGridView4.Columns[4].Width = 100;
                    dataGridView4.Columns[5].Width = 120;
                    dataGridView4.Columns[6].Width = 180;
                    dataGridView4.Columns[7].Width = 101;

                    CargaComboxEx();

                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        public void CargaComboxEx()
        {
            if (comboBox4.Items.Count == 0 && comboBox5.Items.Count == 0 && comboBox6.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from clase_extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            Cl_Extension.Add(lee["Clas_Extension"].ToString());
                        }
                        lee.Close();
                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;
                        while (lee.Read())
                        {
                            Agrega = true;
                            if (Cod_Centro.Count > 0)
                            {
                                for (int i = 0; i < Cod_Centro.Count; i++)
                                {
                                    if (Cod_Centro[i].Equals(lee["Codi_Centro"].ToString()))
                                    {
                                        Agrega = false;
                                    }
                                }
                            }
                            else
                            {
                                Agrega = true;
                            }

                            if (Agrega == true)
                            {
                                Cod_Centro.Add(lee["Codi_Centro"].ToString());
                            }
                        }
                        Env_Diario.Add("S");
                        Env_Diario.Add("N");
                    }
                    lee.Close();
                    Conexion.Close();
                    Cl_Extension.Sort();
                    Cod_Centro.Sort();

                    for (int i = 0; i < Cl_Extension.Count; i++)
                    {
                        comboBox4.Items.Add(Cl_Extension[i]);
                    }
                    for (int i = 0; i < Cod_Centro.Count; i++)
                    {
                        comboBox5.Items.Add(Cod_Centro[i]);
                    }
                    for (int i = 0; i < Env_Diario.Count; i++)
                    {
                        comboBox6.Items.Add(Env_Diario[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowEx.Count > 0)
                {
                    if (DatosIgualesEx() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox29.Text = DatosRowEx[0];
                            textBox30.Text = DatosRowEx[1];
                            textBox31.Text = DatosRowEx[2];
                            textBox32.Text = DatosRowEx[3];
                            comboBox4.Text = DatosRowEx[4];
                            comboBox5.Text = DatosRowEx[5];
                            textBox33.Text = DatosRowEx[6];
                            comboBox6.Text = DatosRowEx[7];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView4_CellClick(null, new DataGridViewCellEventArgs(0, CeldaEx));
                        }
                        else
                        {
                            GuardaDatosEx();
                            CargaTablasEx();
                            CargaCellEx(e);
                        }
                    }
                    else
                    {
                        CargaCellEx(e);
                    }
                }
                else
                {
                    CargaCellEx(e);
                }
            }
        }

        public void CargaCellEx(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView4.ClearSelection();
                dataGridView4.Rows[e.RowIndex].Selected = true;
                CeldaEx = e.RowIndex;

                CantidadCeldasEx = dataGridView4.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowEx = new List<string>();

                while (Posicion < CantidadCeldasEx)
                {
                    DatosRowEx.Add(dataGridView4.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox29.Text = DatosRowEx[0];
                textBox30.Text = DatosRowEx[1];
                textBox31.Text = DatosRowEx[2];
                textBox32.Text = DatosRowEx[3];
                comboBox4.Text = DatosRowEx[4];
                comboBox5.Text = DatosRowEx[5];
                textBox33.Text = DatosRowEx[6];
                comboBox6.Text = DatosRowEx[7];
            }
            catch
            {
                dataGridView4.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesEx()
        {
            if (DatosRowEx[0].Equals(textBox29.Text) && DatosRowEx[1].Equals(textBox30.Text) && DatosRowEx[2].Equals(textBox31.Text) &&
                DatosRowEx[3].Equals(textBox32.Text) && DatosRowEx[4].Equals(comboBox4.Text) && DatosRowEx[5].Equals(comboBox5.Text) &&
                DatosRowEx[6].Equals(textBox33.Text) && DatosRowEx[7].Equals(comboBox6.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (DatosIgualesEx() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosEx();
                CargaTablasEx();
                dataGridView4_CellClick(dataGridView4, new DataGridViewCellEventArgs(0, CeldaEx));
            }
        }

        public void GuardaDatosEx()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    for (int i = 0; i < DatosRowEx.Count; i++)
                    {
                        SalvaEX.Add(DatosRowEx[i]);
                    }
                    query = "delete from extensiones where Nume_Extension =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowEx[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowEx();
                    query = @"insert into extensiones (Nume_Extension, Nume_Folio, Nomb_Extension, Resp_Extension, Clas_Extension, Codi_Centro, Corr_Extension, Envi_Diario_Extension) 
                            values (?NumEx, ?NumFol, ?NombEx, ?RespEx, ?ClasEx, ?CodCem, ?CorrEx, ?Envi)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?NumEx", DatosRowEx[0]);
                    comando.Parameters.AddWithValue("?NumFol", DatosRowEx[1]);
                    comando.Parameters.AddWithValue("?NombEx", DatosRowEx[2]);
                    comando.Parameters.AddWithValue("?RespEx", DatosRowEx[3]);
                    comando.Parameters.AddWithValue("?ClasEx", DatosRowEx[4]);
                    comando.Parameters.AddWithValue("?CodCem", DatosRowEx[5]);
                    comando.Parameters.AddWithValue("?CorrEx", DatosRowEx[6]);
                    comando.Parameters.AddWithValue("?Envi", DatosRowEx[7]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into extensiones (Nume_Extension, Nume_Folio, Nomb_Extension, Resp_Extension, Clas_Extension, Codi_Centro, Corr_Extension, Envi_Diario_Extension) 
                            values (?NumEx, ?NumFol, ?NombEx, ?RespEx, ?ClasEx, ?CodCem, ?CorrEx, ?Envi)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?NumEx", SalvaEX[0]);
                        comando.Parameters.AddWithValue("?NumFol", SalvaEX[1]);
                        comando.Parameters.AddWithValue("?NombEx", SalvaEX[2]);
                        comando.Parameters.AddWithValue("?RespEx", SalvaEX[3]);
                        comando.Parameters.AddWithValue("?ClasEx", SalvaEX[4]);
                        comando.Parameters.AddWithValue("?CodCem", SalvaEX[5]);
                        comando.Parameters.AddWithValue("?CorrEx", SalvaEX[6]);
                        comando.Parameters.AddWithValue("?Envi", SalvaEX[7]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowEx()
        {
            DatosRowEx[0] = textBox29.Text;
            DatosRowEx[1] = textBox30.Text;
            DatosRowEx[2] = textBox31.Text;
            DatosRowEx[3] = textBox32.Text;
            DatosRowEx[4] = comboBox4.Text;
            DatosRowEx[5] = comboBox5.Text;
            DatosRowEx[6] = textBox33.Text;
            DatosRowEx[7] = comboBox6.Text;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            DataRow row = DtableEx.NewRow();
            DtableEx.Rows.Add(row);
            dataGridView4.DataSource = DtableEx;
            dataGridView4.Invoke(new Action(() => { dataGridView4.FirstDisplayedScrollingRowIndex = dataGridView4.RowCount - 1; }));
            dataGridView4_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView4.RowCount - 1));
        }

        private void button29_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar la extensión?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from extensiones where Nume_Extension = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowEx[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablasEx();
                    MessageBox.Show("La extensión se ha borrado exitosamente");
                    dataGridView4_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[0].Value = textBox29.Text;
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[1].Value = textBox30.Text;
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[2].Value = textBox31.Text;
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[3].Value = textBox32.Text;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[4].Value = comboBox4.Text;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[5].Value = comboBox5.Text;
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[6].Value = textBox33.Text;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView4.SelectedRows[0].Cells[7].Value = comboBox6.Text;
        }

        #endregion

        #region Centros de Costo

        DataTable DtableCC;
        List<string> DatosRowCC = new List<string>();
        int CeldaCC = 0;
        int CantidadCeldasCC;
        List<string> SalvaCC = new List<string>();

        public void CargaTablaCC()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From centros_costo";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableCC = new DataTable();
                        adapter.Fill(DtableCC);
                        dataGridView5.DataSource = DtableCC;
                    }
                    dataGridView5.EnableHeadersVisualStyles = false;
                    dataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView5.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);


                    dataGridView5.Columns[0].HeaderText = "Cod Centro de Costo";
                    dataGridView5.Columns[1].HeaderText = "Nombre Centro de Costo";
                    dataGridView5.Columns[2].HeaderText = "E-Mail";

                    dataGridView5.Columns[0].Width = 170;
                    dataGridView5.Columns[1].Width = 273;
                    dataGridView5.Columns[2].Width = 273;


                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowCC.Count > 0)
                {
                    if (DatosIgualesCC() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox34.Text = DatosRowCC[0];
                            textBox35.Text = DatosRowCC[1];
                            textBox36.Text = DatosRowCC[2];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView5_CellClick(null, new DataGridViewCellEventArgs(0, CeldaCC));
                        }
                        else
                        {
                            GuardaDatosCC();
                            CargaTablaCC();
                            CargaCellCC(e);
                        }
                    }
                    else
                    {
                        CargaCellCC(e);
                    }
                }
                else
                {
                    CargaCellCC(e);
                }
            }
        }


        public void CargaCellCC(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView5.ClearSelection();
                dataGridView5.Rows[e.RowIndex].Selected = true;
                CeldaCC = e.RowIndex;

                CantidadCeldasCC = dataGridView5.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowCC = new List<string>();

                while (Posicion < CantidadCeldasCC)
                {
                    DatosRowCC.Add(dataGridView5.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox34.Text = DatosRowCC[0];
                textBox35.Text = DatosRowCC[1];
                textBox36.Text = DatosRowCC[2];
            }
            catch
            {
                dataGridView5.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesCC()
        {
            if (DatosRowCC[0].Equals(textBox34.Text) && DatosRowCC[1].Equals(textBox35.Text) && DatosRowCC[2].Equals(textBox36.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (DatosIgualesCC() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosCC();
                CargaTablaCC();
                dataGridView5_CellClick(dataGridView5, new DataGridViewCellEventArgs(0, CeldaCC));
            }
        }

        public void GuardaDatosCC()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowCC.Count; i++)
                    {
                        SalvaCC.Add(DatosRowCC[i]);
                    }

                    query = "delete from centros_costo where Codi_Centro =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowCC[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowCC();
                    query = @"insert into centros_costo (Codi_Centro, Nomb_Centro, Corr_Centro) 
                            values (?CC, ?NC, ?CRC)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CC", DatosRowCC[0]);
                    comando.Parameters.AddWithValue("?NC", DatosRowCC[1]);
                    comando.Parameters.AddWithValue("?CRC", DatosRowCC[2]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into centros_costo (Codi_Centro, Nomb_Centro, Corr_Centro) 
                            values (?CC, ?NC, ?CRC)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CC", SalvaCC[0]);
                        comando.Parameters.AddWithValue("?NC", SalvaCC[1]);
                        comando.Parameters.AddWithValue("?CRC", SalvaCC[2]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("Ha ocurrido un error al revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowCC()
        {
            DatosRowCC[0] = textBox34.Text;
            DatosRowCC[1] = textBox35.Text;
            DatosRowCC[2] = textBox36.Text;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            DataRow row = DtableCC.NewRow();
            DtableCC.Rows.Add(row);
            dataGridView5.DataSource = DtableCC;
            dataGridView5.Invoke(new Action(() => { dataGridView5.FirstDisplayedScrollingRowIndex = dataGridView5.RowCount - 1; }));
            dataGridView5_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView5.RowCount - 1));
        }

        private void button31_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el centro de costo?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from centros_costo where Codi_Centro = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowCC[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaCC();
                    MessageBox.Show("El centro de costo se ha borrado exitosamente");
                    dataGridView5_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            dataGridView5.SelectedRows[0].Cells[0].Value = textBox34.Text;
        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            dataGridView5.SelectedRows[0].Cells[1].Value = textBox35.Text;
        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {
            dataGridView5.SelectedRows[0].Cells[2].Value = textBox36.Text;
        }

        #endregion

        #region Troncales

        DataTable DtableTR;
        List<string> DatosRowTR = new List<string>();
        List<string> OperadorCom = new List<string>();
        List<string> SalvaTR = new List<string>();
        int CeldaTR = 0;
        int CantidadCeldasTR;

        public void CargaTablaTR()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From troncales";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableTR = new DataTable();
                        adapter.Fill(DtableTR);
                        dataGridView6.DataSource = DtableTR;
                    }
                    dataGridView6.EnableHeadersVisualStyles = false;
                    dataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView6.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);


                    dataGridView6.Columns[0].HeaderText = "Troncal";
                    dataGridView6.Columns[1].HeaderText = "Número roncal";
                    dataGridView6.Columns[2].HeaderText = "Con Acceso";
                    dataGridView6.Columns[3].HeaderText = "Operador";

                    dataGridView6.Columns[0].Width = 76;
                    dataGridView6.Columns[1].Width = 273;
                    dataGridView6.Columns[2].Width = 273;
                    dataGridView6.Columns[3].Width = 76;

                    CargaComboTR();
                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        public void CargaComboTR()
        {
            if (comboBox7.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from operadores";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;

                        while (lee.Read())
                        {
                            OperadorCom.Add(lee["Operador"].ToString());
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    OperadorCom.Sort();

                    for (int i = 0; i < OperadorCom.Count; i++)
                    {
                        comboBox7.Items.Add(OperadorCom[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowTR.Count > 0)
                {
                    if (DatosIgualesTR() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox37.Text = DatosRowTR[0];
                            textBox38.Text = DatosRowTR[1];
                            textBox39.Text = DatosRowTR[2];
                            comboBox7.Text = DatosRowTR[3];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView6_CellClick(null, new DataGridViewCellEventArgs(0, CeldaTR));
                        }
                        else
                        {
                            GuardaDatosTR();
                            CargaTablaTR();
                            CargaCellTR(e);
                        }
                    }
                    else
                    {
                        CargaCellTR(e);
                    }
                }
                else
                {
                    CargaCellTR(e);
                }
            }
        }


        public void CargaCellTR(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView6.ClearSelection();
                dataGridView6.Rows[e.RowIndex].Selected = true;
                CeldaTR = e.RowIndex;

                CantidadCeldasTR = dataGridView6.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowTR = new List<string>();

                while (Posicion < CantidadCeldasTR)
                {
                    DatosRowTR.Add(dataGridView6.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox37.Text = DatosRowTR[0];
                textBox38.Text = DatosRowTR[1];
                textBox39.Text = DatosRowTR[2];
                comboBox7.Text = DatosRowTR[3];
            }
            catch
            {
                dataGridView6.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesTR()
        {
            if (DatosRowTR[0].Equals(textBox37.Text) && DatosRowTR[1].Equals(textBox38.Text) && DatosRowTR[2].Equals(textBox39.Text) && DatosRowTR[3].Equals(comboBox7.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (DatosIgualesTR() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosTR();
                CargaTablaTR();
                dataGridView6_CellClick(dataGridView6, new DataGridViewCellEventArgs(0, CeldaTR));
            }
        }

        public void GuardaDatosTR()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowTR.Count; i++)
                    {
                        SalvaTR.Add(DatosRowTR[i]);
                    }

                    query = "delete from troncales where Line_Troncal =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowTR[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowTR();
                    query = @"insert into troncales (Line_Troncal, Nume_Troncal_Operador, Nume_Acceso_Troncal, Operador) 
                            values (?LT, ?NT, ?NAT, ?OP)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?LT", DatosRowTR[0]);
                    comando.Parameters.AddWithValue("?NT", DatosRowTR[1]);
                    comando.Parameters.AddWithValue("?NAT", DatosRowTR[2]);
                    comando.Parameters.AddWithValue("?OP", DatosRowTR[3]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into troncales (Line_Troncal, Nume_Troncal_Operador, Nume_Acceso_Troncal, Operador) 
                            values (?LT, ?NT, ?NAT, ?OP)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?LT", Convert.ToInt32(SalvaTR[0]));
                        comando.Parameters.AddWithValue("?NT", SalvaTR[1]);
                        comando.Parameters.AddWithValue("?NAT", SalvaTR[2]);
                        comando.Parameters.AddWithValue("?OP", SalvaTR[3]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                        MessageBox.Show("Los cambios se han revertido");
                    }
                }
                catch
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está correctamente conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowTR()
        {
            DatosRowTR[0] = textBox37.Text;
            DatosRowTR[1] = textBox38.Text;
            DatosRowTR[2] = textBox39.Text;
            DatosRowTR[3] = comboBox7.Text;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            DataRow row = DtableTR.NewRow();
            DtableTR.Rows.Add(row);
            dataGridView6.DataSource = DtableTR;
            dataGridView6.Invoke(new Action(() => { dataGridView6.FirstDisplayedScrollingRowIndex = dataGridView6.RowCount - 1; }));
            dataGridView6_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView6.RowCount - 1));
        }

        private void button33_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar la troncal?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from troncales where Line_Troncal = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowTR[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaTR();
                    MessageBox.Show("La troncal se ha borrado exitosamente");
                    dataGridView6_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {
            dataGridView6.SelectedRows[0].Cells[0].Value = textBox37.Text;
        }

        private void textBox38_TextChanged(object sender, EventArgs e)
        {
            dataGridView6.SelectedRows[0].Cells[1].Value = textBox38.Text;
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            dataGridView6.SelectedRows[0].Cells[2].Value = textBox39.Text;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView6.SelectedRows[0].Cells[3].Value = comboBox7.Text;
        }

        #endregion

        #region Codigos Personales

        DataTable DtableCP;
        List<string> DatosRowCP = new List<string>();
        List<string> Nume_Extension = new List<string>();
        List<string> SalvaCP = new List<string>();
        int CeldaCP = 0;
        int CantidadCeldasCP;

        public void CargaTablaCP()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From codigos_personales";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableCP = new DataTable();
                        adapter.Fill(DtableCP);
                        dataGridView7.DataSource = DtableCP;
                    }
                    dataGridView7.EnableHeadersVisualStyles = false;
                    dataGridView7.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView7.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);


                    dataGridView7.Columns[0].HeaderText = "Código";
                    dataGridView7.Columns[1].HeaderText = "Asignado a";
                    dataGridView7.Columns[2].HeaderText = "Extensión";
                    dataGridView7.Columns[3].HeaderText = "E-Mail";

                    dataGridView7.Columns[0].Width = 100;
                    dataGridView7.Columns[1].Width = 250;
                    dataGridView7.Columns[2].Width = 100;
                    dataGridView7.Columns[3].Width = 248;

                    CargaComboCP();
                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        public void CargaComboCP()
        {
            if (comboBox8.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;

                        while (lee.Read())
                        {
                            Nume_Extension.Add(lee["Nume_Extension"].ToString());
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    Nume_Extension.Sort();

                    for (int i = 0; i < Nume_Extension.Count; i++)
                    {
                        comboBox8.Items.Add(Nume_Extension[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView7_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowCP.Count > 0)
                {
                    if (DatosIgualesCP() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox40.Text = DatosRowCP[0];
                            textBox41.Text = DatosRowCP[1];
                            comboBox8.Text = DatosRowCP[2];
                            textBox67.Text = DatosRowCP[3];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView7_CellClick(null, new DataGridViewCellEventArgs(0, CeldaCP));
                        }
                        else
                        {
                            GuardaDatosCP();
                            CargaTablaCP();
                            CargaCellCP(e);
                        }
                    }
                    else
                    {
                        CargaCellCP(e);
                    }
                }
                else
                {
                    CargaCellCP(e);
                }
            }
        }


        public void CargaCellCP(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView7.ClearSelection();
                dataGridView7.Rows[e.RowIndex].Selected = true;
                CeldaCP = e.RowIndex;

                CantidadCeldasCP = dataGridView7.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowCP = new List<string>();

                while (Posicion < CantidadCeldasCP)
                {
                    DatosRowCP.Add(dataGridView7.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox40.Text = DatosRowCP[0];
                textBox41.Text = DatosRowCP[1];
                comboBox8.Text = DatosRowCP[2];
                textBox67.Text = DatosRowCP[3];
            }
            catch
            {
                dataGridView7.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesCP()
        {
            if (DatosRowCP[0].Equals(textBox40.Text) && DatosRowCP[1].Equals(textBox41.Text) && DatosRowCP[2].Equals(comboBox8.Text) && DatosRowCP[3].Equals(textBox67.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (DatosIgualesCP() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosCP();
                CargaTablaCP();
                dataGridView7_CellClick(dataGridView7, new DataGridViewCellEventArgs(0, CeldaCP));
            }
        }

        public void GuardaDatosCP()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowCP.Count; i++)
                    {
                        SalvaCP.Add(DatosRowCP[i]);
                    }

                    query = "delete from codigos_personales where Codi_Personal =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowCP[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowCP();
                    query = @"insert into codigos_personales (Codi_Personal, Nomb_Cod_Personal, Nume_Extension, Corr_Codper) 
                            values (?CP, ?NCP, ?NE, ?CC)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CP", DatosRowCP[0]);
                    comando.Parameters.AddWithValue("?NCP", DatosRowCP[1]);
                    comando.Parameters.AddWithValue("?NE", DatosRowCP[2]);
                    comando.Parameters.AddWithValue("?CC", DatosRowCP[3]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into codigos_personales (Codi_Personal, Nomb_Cod_Personal, Nume_Extension, Corr_Codper) 
                            values (?CP, ?NCP, ?NE, ?CC)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CP", SalvaCP[0]);
                        comando.Parameters.AddWithValue("?NCP", SalvaCP[1]);
                        comando.Parameters.AddWithValue("?NE", SalvaCP[2]);
                        comando.Parameters.AddWithValue("?CC", SalvaCP[3]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("Ha ocurrido un error al revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowCP()
        {
            DatosRowCP[0] = textBox40.Text;
            DatosRowCP[1] = textBox41.Text;
            DatosRowCP[2] = comboBox8.Text;
            DatosRowCP[3] = textBox67.Text;
        }

        private void button34_Click(object sender, EventArgs e)
        {
            DataRow row = DtableCP.NewRow();
            DtableCP.Rows.Add(row);
            dataGridView7.DataSource = DtableCP;
            dataGridView7.Invoke(new Action(() => { dataGridView7.FirstDisplayedScrollingRowIndex = dataGridView7.RowCount - 1; }));
            dataGridView7_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView7.RowCount - 1));
        }

        private void button35_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea el código personal?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from codigos_personales where Codi_Personal = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowCP[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaCP();
                    MessageBox.Show("El código personal se ha borrado exitosamente");
                    dataGridView7_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox40_TextChanged(object sender, EventArgs e)
        {
            dataGridView7.SelectedRows[0].Cells[0].Value = textBox40.Text;
        }

        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            dataGridView7.SelectedRows[0].Cells[1].Value = textBox41.Text;
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView7.SelectedRows[0].Cells[2].Value = comboBox8.Text;
        }

        private void textBox67_TextChanged(object sender, EventArgs e)
        {
            dataGridView7.SelectedRows[0].Cells[3].Value = textBox67.Text;
        }

        #endregion

        #region Clase Extensioes

        DataTable DtableCE;
        List<string> DatosRowCE = new List<string>();
        List<string> Plan_Tarifario = new List<string>();
        List<string> SalvaCE = new List<string>();
        int CeldaCE = 0;
        int CantidadCeldasCE;

        public void CargaTablaCE()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From clase_extensiones";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableCE = new DataTable();
                        adapter.Fill(DtableCE);
                        dataGridView8.DataSource = DtableCE;
                    }
                    dataGridView8.EnableHeadersVisualStyles = false;
                    dataGridView8.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView8.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);



                    dataGridView8.Columns[0].HeaderText = "Clase Ext";
                    dataGridView8.Columns[1].HeaderText = "Detalle clase de extensión";
                    dataGridView8.Columns[2].HeaderText = "Plan tarifa";
                    dataGridView8.Columns[3].HeaderText = "Envío interface 1";
                    dataGridView8.Columns[4].HeaderText = "Envío interface 2";
                    dataGridView8.Columns[5].HeaderText = "Envío interface 3";

                    dataGridView8.Columns[0].Width = 76;
                    dataGridView8.Columns[1].Width = 273;
                    dataGridView8.Columns[2].Width = 76;
                    dataGridView8.Columns[3].Width = 91;
                    dataGridView8.Columns[4].Width = 91;
                    dataGridView8.Columns[5].Width = 91;

                    CargaComboCE();
                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        public void CargaComboCE()
        {
            if (comboBox9.Items.Count == 0)
            {
                try
                {
                    Plan_Tarifario.Add("001");
                    Plan_Tarifario.Add("002");
                    Plan_Tarifario.Add("003");
                    Plan_Tarifario.Sort();

                    for (int i = 0; i < Plan_Tarifario.Count; i++)
                    {
                        comboBox9.Items.Add(Plan_Tarifario[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView8_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowCE.Count > 0)
                {
                    if (DatosIgualesCE() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox42.Text = DatosRowCE[0];
                            textBox43.Text = DatosRowCE[1];
                            comboBox9.Text = DatosRowCE[2];
                            comboBox12.Text = DatosRowCE[3];
                            comboBox13.Text = DatosRowCE[4];
                            comboBox14.Text = DatosRowCE[5];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView8_CellClick(null, new DataGridViewCellEventArgs(0, CeldaCE));
                        }
                        else
                        {
                            GuardaDatosCE();
                            CargaTablaCE();
                            CargaCellCE(e);
                        }
                    }
                    else
                    {
                        CargaCellCE(e);
                    }
                }
                else
                {
                    CargaCellCE(e);
                }
            }
        }


        public void CargaCellCE(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView8.ClearSelection();
                dataGridView8.Rows[e.RowIndex].Selected = true;
                CeldaCE = e.RowIndex;

                CantidadCeldasCE = dataGridView8.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowCE = new List<string>();

                while (Posicion < CantidadCeldasCE)
                {
                    DatosRowCE.Add(dataGridView8.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox42.Text = DatosRowCE[0];
                textBox43.Text = DatosRowCE[1];
                comboBox9.Text = DatosRowCE[2];
                comboBox12.Text = DatosRowCE[3];
                comboBox13.Text = DatosRowCE[4];
                comboBox14.Text = DatosRowCE[5];
            }
            catch
            {
                dataGridView8.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesCE()
        {
            if (DatosRowCE[0].Equals(textBox42.Text) && DatosRowCE[1].Equals(textBox43.Text) && DatosRowCE[2].Equals(comboBox9.Text) && DatosRowCE[3].Equals(comboBox12.Text) && DatosRowCE[4].Equals(comboBox13.Text) && DatosRowCE[5].Equals(comboBox14.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (DatosIgualesCE() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosCE();
                CargaTablaCE();
                dataGridView8_CellClick(dataGridView8, new DataGridViewCellEventArgs(0, CeldaCE));
            }
        }

        public void GuardaDatosCE()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowCE.Count; i++)
                    {
                        SalvaCE.Add(DatosRowCE[i]);
                    }

                    query = "delete from clase_extensiones where Clas_Extension =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowCE[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowCE();
                    query = @"insert into clase_extensiones (Clas_Extension, Deta_Clase_Extension, Plan_Tarifario, Envi_Interface_01, Envi_Interface_02, Envi_Interface_03) 
                            values (?CE, ?DCE, ?PT, ?EI1, ?EI2, ?EI3)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CE", DatosRowCE[0]);
                    comando.Parameters.AddWithValue("?DCE", DatosRowCE[1]);
                    comando.Parameters.AddWithValue("?PT", DatosRowCE[2]);
                    comando.Parameters.AddWithValue("?EI1", DatosRowCE[3]);
                    comando.Parameters.AddWithValue("?EI2", DatosRowCE[4]);
                    comando.Parameters.AddWithValue("?EI3", DatosRowCE[5]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into clase_extensiones (Clas_Extension, Deta_Clase_Extension, Plan_Tarifario, Envi_Interface_01, Envi_Interface_02, Envi_Interface_03) 
                            values (?CE, ?DCE, ?PT, ?EI1, ?EI2, ?EI3)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CE", SalvaCE[0]);
                        comando.Parameters.AddWithValue("?DCE", SalvaCE[1]);
                        comando.Parameters.AddWithValue("?PT", SalvaCE[2]);
                        comando.Parameters.AddWithValue("?EI1", SalvaCE[3]);
                        comando.Parameters.AddWithValue("?EI2", SalvaCE[4]);
                        comando.Parameters.AddWithValue("?EI3", SalvaCE[5]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("Ha ocurrido un error al revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowCE()
        {
            DatosRowCE[0] = textBox42.Text;
            DatosRowCE[1] = textBox43.Text;
            DatosRowCE[2] = comboBox9.Text;
            DatosRowCE[3] = comboBox12.Text;
            DatosRowCE[4] = comboBox13.Text;
            DatosRowCE[5] = comboBox14.Text;
        }

        private void button36_Click(object sender, EventArgs e)
        {
            DataRow row = DtableCE.NewRow();
            DtableCE.Rows.Add(row);
            dataGridView8.DataSource = DtableCE;
            dataGridView8.Invoke(new Action(() => { dataGridView8.FirstDisplayedScrollingRowIndex = dataGridView8.RowCount - 1; }));
            dataGridView8_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView8.RowCount - 1));
        }

        private void button37_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea la clase de extensión?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from clase_extensiones where Clas_Extension = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowCE[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaCE();
                    MessageBox.Show("La calse de extensión se ha borrado exitosamente");
                    dataGridView8_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox42_TextChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[0].Value = textBox42.Text;
        }

        private void textBox43_TextChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[1].Value = textBox43.Text;
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[2].Value = comboBox9.Text;
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[3].Value = comboBox12.Text;
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[4].Value = comboBox13.Text;
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView8.SelectedRows[0].Cells[5].Value = comboBox14.Text;
        }


        #endregion

        #region Clase de llamadas

        DataTable DtableCL;
        List<string> DatosRowCL = new List<string>();
        List<int> tarifa = new List<int>();
        List<string> SalvaCL = new List<string>();
        int CeldaCL = 0;
        int CantidadCeldasCL;

        public void CargaTablaCL()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From clase_llamadas";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableCL = new DataTable();
                        adapter.Fill(DtableCL);
                        dataGridView9.DataSource = DtableCL;
                    }
                    dataGridView9.EnableHeadersVisualStyles = false;
                    dataGridView9.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView9.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);



                    dataGridView9.Columns[0].HeaderText = "Clase";
                    dataGridView9.Columns[1].HeaderText = "Detalle calse de llamada";
                    dataGridView9.Columns[2].HeaderText = "Procesa S/N";
                    dataGridView9.Columns[3].HeaderText = "Tarifa Nº";

                    dataGridView9.Columns[0].Width = 122;
                    dataGridView9.Columns[1].Width = 364;
                    dataGridView9.Columns[2].Width = 122;
                    dataGridView9.Columns[3].Width = 91;

                    CargaComboCL();
                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        public void CargaComboCL()
        {
            if (comboBox10.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from plan_tarifario_001";

                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;

                        while (lee.Read())
                        {
                            tarifa.Add(Convert.ToInt32(lee["Codi_Tarifa"]));
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    tarifa.Sort();

                    for (int i = 0; i < tarifa.Count; i++)
                    {
                        comboBox10.Items.Add(tarifa[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView9_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowCL.Count > 0)
                {
                    if (DatosIgualesCL() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox47.Text = DatosRowCL[0];
                            textBox48.Text = DatosRowCL[1];
                            comboBox15.Text = DatosRowCL[2];
                            comboBox10.Text = DatosRowCL[3];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView9_CellClick(null, new DataGridViewCellEventArgs(0, CeldaCL));
                        }
                        else
                        {
                            GuardaDatosCL();
                            CargaTablaCL();
                            CargaCellCL(e);
                        }
                    }
                    else
                    {
                        CargaCellCL(e);
                    }
                }
                else
                {
                    CargaCellCL(e);
                }
            }
        }


        public void CargaCellCL(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView9.ClearSelection();
                dataGridView9.Rows[e.RowIndex].Selected = true;
                CeldaCL = e.RowIndex;

                CantidadCeldasCL = dataGridView9.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowCL = new List<string>();

                while (Posicion < CantidadCeldasCL)
                {
                    DatosRowCL.Add(dataGridView9.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox47.Text = DatosRowCL[0];
                textBox48.Text = DatosRowCL[1];
                comboBox15.Text = DatosRowCL[2];
                comboBox10.Text = DatosRowCL[3];
            }
            catch
            {
                dataGridView9.Rows[0].Selected = true;
            }

        }

        public bool DatosIgualesCL()
        {
            if (DatosRowCL[0].Equals(textBox47.Text) && DatosRowCL[1].Equals(textBox48.Text) && DatosRowCL[3].Equals(comboBox10.Text) && DatosRowCL[2].Equals(comboBox15.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (DatosIgualesCL() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosCL();
                CargaTablaCL();
                dataGridView9_CellClick(dataGridView9, new DataGridViewCellEventArgs(0, CeldaCL));
            }
        }

        public void GuardaDatosCL()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowCL.Count; i++)
                    {
                        SalvaCL.Add(DatosRowCL[i]);
                    }

                    query = "delete from clase_llamadas where Clase_Llamada =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowCL[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowCL();
                    query = @"insert into clase_llamadas (Clase_Llamada, Nombre_Clase_Llamada, tarificarSN, tarifa) 
                            values (?CL, ?NCL, ?TSN, ?T)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CL", DatosRowCL[0]);
                    comando.Parameters.AddWithValue("?NCL", DatosRowCL[1]);
                    comando.Parameters.AddWithValue("?TSN", DatosRowCL[2]);
                    comando.Parameters.AddWithValue("?T", DatosRowCL[3]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into clase_llamadas (Clase_Llamada, Nombre_Clase_Llamada, tarificarSN, tarifa) 
                            values (?CL, ?NCL, ?TSN, ?T)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CL", SalvaCL[0]);
                        comando.Parameters.AddWithValue("?NCL", SalvaCL[1]);
                        comando.Parameters.AddWithValue("?TSN", SalvaCL[2]);
                        comando.Parameters.AddWithValue("?T", SalvaCL[3]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("Ha ocurrido un error al revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowCL()
        {
            DatosRowCL[0] = textBox47.Text;
            DatosRowCL[1] = textBox48.Text;
            DatosRowCL[2] = comboBox15.Text;
            DatosRowCL[3] = comboBox10.Text;
        }

        private void button38_Click(object sender, EventArgs e)
        {
            DataRow row = DtableCL.NewRow();
            DtableCL.Rows.Add(row);
            dataGridView9.DataSource = DtableCL;
            dataGridView9.Invoke(new Action(() => { dataGridView9.FirstDisplayedScrollingRowIndex = dataGridView9.RowCount - 1; }));
            dataGridView9_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView9.RowCount - 1));
        }

        private void button39_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar la calse de llamada?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from clase_llamadas where Clase_Llamada = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowCL[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaCL();
                    MessageBox.Show("La calse de llamada se ha borrado exitosamente");
                    dataGridView9_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox47_TextChanged(object sender, EventArgs e)
        {
            dataGridView9.SelectedRows[0].Cells[0].Value = textBox47.Text;
        }

        private void textBox48_TextChanged(object sender, EventArgs e)
        {
            dataGridView9.SelectedRows[0].Cells[1].Value = textBox48.Text;
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView9.SelectedRows[0].Cells[2].Value = comboBox15.Text;
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView9.SelectedRows[0].Cells[3].Value = comboBox10.Text;
        }

        #endregion

        #region Operadores

        DataTable DtableOP;
        List<string> DatosRowOP = new List<string>();
        List<string> SalvaOP = new List<string>();
        int CeldaOP = 0;
        int CantidadCeldasOP;

        public void CargaTablaOP()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From operadores";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableOP = new DataTable();
                        adapter.Fill(DtableOP);
                        dataGridView10.DataSource = DtableOP;
                    }
                    dataGridView10.EnableHeadersVisualStyles = false;
                    dataGridView10.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView10.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);



                    dataGridView10.Columns[0].HeaderText = "Operador";
                    dataGridView10.Columns[1].HeaderText = "Detalle nombre del operador";

                    dataGridView10.Columns[0].Width = 129;
                    dataGridView10.Columns[1].Width = 571;

                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!");
            }
        }

        private void dataGridView10_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowOP.Count > 0)
                {
                    if (DatosIgualesOP() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            textBox50.Text = DatosRowOP[0];
                            textBox51.Text = DatosRowOP[1];
                            MessageBox.Show("Los cambios se descartaron");
                            dataGridView10_CellClick(null, new DataGridViewCellEventArgs(0, CeldaOP));
                        }
                        else
                        {
                            GuardaDatosOP();
                            CargaTablaOP();
                            CargaCellOP(e);
                        }
                    }
                    else
                    {
                        CargaCellOP(e);
                    }
                }
                else
                {
                    CargaCellOP(e);
                }
            }
        }


        public void CargaCellOP(DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView10.ClearSelection();
                dataGridView10.Rows[e.RowIndex].Selected = true;
                CeldaOP = e.RowIndex;

                CantidadCeldasOP = dataGridView10.GetCellCount(DataGridViewElementStates.Selected);
                Posicion = 0;
                DatosRowOP = new List<string>();

                while (Posicion < CantidadCeldasOP)
                {
                    DatosRowOP.Add(dataGridView10.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                    Posicion++;
                }

                textBox50.Text = DatosRowOP[0];
                textBox51.Text = DatosRowOP[1];
            }
            catch
            {
                dataGridView10.Rows[0].Selected = true;
            }
        }

        public bool DatosIgualesOP()
        {
            if (DatosRowOP[0].Equals(textBox50.Text) && DatosRowOP[1].Equals(textBox51.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (DatosIgualesOP() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosOP();
                CargaTablaOP();
                dataGridView10_CellClick(dataGridView10, new DataGridViewCellEventArgs(0, CeldaOP));
            }
        }

        public void GuardaDatosOP()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();

                    for (int i = 0; i < DatosRowOP.Count; i++)
                    {
                        SalvaOP.Add(DatosRowOP[i]);
                    }

                    query = "delete from operadores where Operador =?CT";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?CT", DatosRowOP[0]);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowOP();
                    query = @"insert into operadores (Operador, Nombre_Operador) 
                            values (?OP, ?NOP)";
                    comando = new MySqlCommand(query, Conexion);
                    comando.Parameters.AddWithValue("?OP", DatosRowOP[0]);
                    comando.Parameters.AddWithValue("?NOP", DatosRowOP[1]);

                    comando.ExecuteNonQuery();
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = @"insert into operadores (Operador, Nombre_Operador) 
                            values (?OP, ?NOP)";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?OP", SalvaOP[0]);
                        comando.Parameters.AddWithValue("?NOP", SalvaOP[1]);

                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }

                    MessageBox.Show("Los cambios se han revertido");
                }
                catch
                {
                    MessageBox.Show("Ha ocurrido un error al revertir los cambios, ¿Está conectado a la base de datos?");
                }
            }
        }

        public void CartgaDatosRowOP()
        {
            DatosRowOP[0] = textBox50.Text;
            DatosRowOP[1] = textBox51.Text;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            DataRow row = DtableOP.NewRow();
            DtableOP.Rows.Add(row);
            dataGridView10.DataSource = DtableOP;
            dataGridView10.Invoke(new Action(() => { dataGridView10.FirstDisplayedScrollingRowIndex = dataGridView10.RowCount - 1; }));
            dataGridView10_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView10.RowCount - 1));
        }

        private void button41_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el operador?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from operadores where Operador = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowOP[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaOP();
                    MessageBox.Show("El operador se ha borrado exitosamente");
                    dataGridView10_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            dataGridView10.SelectedRows[0].Cells[0].Value = textBox50.Text;
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            dataGridView10.SelectedRows[0].Cells[1].Value = textBox51.Text;
        }

        #endregion


        #endregion

        #region Parametros

        #region Paso


        private void tabControl5_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabControl5.SelectedIndex != 0 && DatosRowTF.Count > 0) { dataGridView11_CellClick(dataGridView11, new DataGridViewCellEventArgs(0, CeldaTF)); }
            else if (tabControl5.SelectedIndex != 1 && DatosRowF.Count > 0) { dataGridView12_CellClick(dataGridView12, new DataGridViewCellEventArgs(0, CeldaF)); }
            else if (tabControl5.SelectedIndex != 2 && DatosRowNI.Count > 0) { dataGridView13_CellClick(dataGridView13, new DataGridViewCellEventArgs(0, CeldaNI)); }
            else if (tabControl5.SelectedIndex != 3 && DatosRowINS.Count > 0) { dataGridView17_CellClick(dataGridView17, new DataGridViewCellEventArgs(0, CeldaINS)); }

            if (e.TabPageIndex == 0) { tabPage22.Controls.Add(button19); CargaTablaTF(); dataGridView11_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            else if (e.TabPageIndex == 1) { tabPage23.Controls.Add(button19); CargaTablaF(); dataGridView12_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            else if (e.TabPageIndex == 2) { tabPage24.Controls.Add(button19); CargaTablaNI(); dataGridView13_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            else if (e.TabPageIndex == 3) { tabPage25.Controls.Add(button19); CargaTablaINS(); dataGridView17_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
        }

        private void button19_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Al salir se decartarán los cambios no guardados, desea continuar?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                tabControl1.SelectTab(0);
            }
        }

        #endregion

        #region Tipos de formato

        DataTable DtableTF;
        List<string> DatosRowTF = new List<string>();
        int CeldaTF = 0;
        int CantidadCeldasTF;
        DataTable SalvaTF;

        public void CargaTablaTF()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From tipos_formato";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableTF = new DataTable();
                        adapter.Fill(DtableTF);
                        dataGridView11.DataSource = DtableTF;
                    }
                    dataGridView11.EnableHeadersVisualStyles = false;
                    dataGridView11.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView11.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView11.Columns[0].HeaderText = "Código";
                    dataGridView11.Columns[1].HeaderText = "Detalle del tipo de formato";

                    dataGridView11.Columns[0].Width = 109;
                    dataGridView11.Columns[1].Width = 420;

                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        private void dataGridView11_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowTF.Count > 0)
                {
                    if (DatosIgualesTF() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView11.Rows[CeldaTF].Selected = true;
                            textBox44.Text = DatosRowTF[0];
                            textBox45.Text = DatosRowTF[1];
                            dataGridView11_CellClick(null, new DataGridViewCellEventArgs(0, CeldaTF));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosTF();
                            CargaTablaTF();
                            CargaCellTF(e);
                        }
                    }
                    else
                    {
                        CargaCellTF(e);
                    }
                }
                else
                {
                    CargaCellTF(e);
                }
            }
        }

        public void CargaCellTF(DataGridViewCellEventArgs e)
        {
            dataGridView11.ClearSelection();
            dataGridView11.Rows[e.RowIndex].Selected = true;
            CeldaTF = e.RowIndex;

            CantidadCeldasTF = dataGridView11.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowTF = new List<string>();

            while (Posicion < CantidadCeldasTF)
            {
                DatosRowTF.Add(dataGridView11.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            textBox44.Text = DatosRowTF[0];
            textBox45.Text = DatosRowTF[1];
        }

        public bool DatosIgualesTF()
        {
            if (DatosRowTF[0].Equals(textBox44.Text) && DatosRowTF[1].Equals(textBox45.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (DatosIgualesTF() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosTF();
                CargaTablaTF();
                dataGridView11_CellClick(dataGridView11, new DataGridViewCellEventArgs(0, CeldaTF));
            }
        }

        public void GuardaDatosTF()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From tipos_formato";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaTF = new DataTable();
                        adapter.Fill(SalvaTF);
                    }
                    query = "delete From tipos_formato";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowTF();
                    comando = new MySqlCommand("insert into tipos_formato values (?Codi_Tipo, ?Nomb_Tipo)", Conexion);
                    foreach (DataGridViewRow row in dataGridView11.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?Codi_Tipo", Convert.ToString(row.Cells["Codi_Tipo"].Value));
                        comando.Parameters.AddWithValue("?Nomb_Tipo", Convert.ToString(row.Cells["Nomb_Tipo"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From tipos_formato";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView11.DataSource = SalvaTF;
                        comando = new MySqlCommand("insert into tipos_formato values (?Codi_Tipo, ?Nomb_Tipo)", Conexion);
                        foreach (DataGridViewRow row in dataGridView11.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?Codi_Tipo", Convert.ToString(row.Cells["Codi_Tipo"].Value));
                            comando.Parameters.AddWithValue("?Nomb_Tipo", Convert.ToString(row.Cells["Nomb_Tipo"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView11_CellClick(dataGridView11, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowTF()
        {
            DatosRowTF[0] = textBox44.Text;
            DatosRowTF[1] = textBox45.Text;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataRow row = DtableTF.NewRow();
            DtableTF.Rows.Add(row);
            dataGridView11.DataSource = DtableTF;
            dataGridView11.Invoke(new Action(() => { dataGridView11.FirstDisplayedScrollingRowIndex = dataGridView11.RowCount - 1; }));
            dataGridView11_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView11.RowCount - 1));
        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {
            dataGridView11.SelectedRows[0].Cells[0].Value = textBox44.Text;
        }

        private void textBox45_TextChanged(object sender, EventArgs e)
        {
            dataGridView11.SelectedRows[0].Cells[1].Value = textBox45.Text;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el formato?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from tipos_formato where Codi_Tipo = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowTF[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaTF();
                    MessageBox.Show("El formato se ha borrado exitosamente");
                    dataGridView11_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        #endregion

        #region Formatos

        DataTable DtableF;
        List<string> DatosRowF = new List<string>();
        List<string> TipoDeFormato = new List<string>();
        int CeldaF = 0;
        int CantidadCeldasF;
        DataTable SalvaF;

        public void CargaTablaF()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From formatos";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableF = new DataTable();
                        adapter.Fill(DtableF);
                        dataGridView12.DataSource = DtableF;
                    }
                    dataGridView12.EnableHeadersVisualStyles = false;
                    dataGridView12.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView12.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView12.Columns[0].HeaderText = "Código";
                    dataGridView12.Columns[1].HeaderText = "Tipo de formato";
                    dataGridView12.Columns[2].HeaderText = "String del formato";
                    dataGridView12.Columns[3].HeaderText = "Detalle del formato";

                    dataGridView12.Columns[0].Width = 113;
                    dataGridView12.Columns[1].Width = 133;
                    dataGridView12.Columns[2].Width = 133;
                    dataGridView12.Columns[3].Width = 133;

                    CargaComboF();
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        public void CargaComboF()
        {
            if (comboBox3.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from tipos_formato";

                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;

                        while (lee.Read())
                        {
                            TipoDeFormato.Add(lee["Codi_Tipo"].ToString());
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    TipoDeFormato.Sort();

                    for (int i = 0; i < TipoDeFormato.Count; i++)
                    {
                        comboBox3.Items.Add(TipoDeFormato[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView12_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowF.Count > 0)
                {
                    if (DatosIgualesF() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView12.Rows[CeldaF].Selected = true;
                            textBox46.Text = DatosRowF[0];
                            comboBox3.Text = DatosRowF[1];
                            textBox49.Text = DatosRowF[2];
                            textBox52.Text = DatosRowF[3];
                            dataGridView12_CellClick(null, new DataGridViewCellEventArgs(0, CeldaF));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosF();
                            CargaTablaF();
                            CargaCellF(e);
                        }
                    }
                    else
                    {
                        CargaCellF(e);
                    }
                }
                else
                {
                    CargaCellF(e);
                }
            }
        }

        public void CargaCellF(DataGridViewCellEventArgs e)
        {
            dataGridView12.ClearSelection();
            dataGridView12.Rows[e.RowIndex].Selected = true;
            CeldaF = e.RowIndex;

            CantidadCeldasF = dataGridView12.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowF = new List<string>();

            while (Posicion < CantidadCeldasF)
            {
                DatosRowF.Add(dataGridView12.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            textBox46.Text = DatosRowF[0];
            comboBox3.Text = DatosRowF[1];
            textBox49.Text = DatosRowF[2];
            textBox52.Text = DatosRowF[3];
        }

        public bool DatosIgualesF()
        {
            if (DatosRowF[0].Equals(textBox46.Text) && DatosRowF[1].Equals(comboBox3.Text) && DatosRowF[2].Equals(textBox49.Text) && DatosRowF[3].Equals(textBox52.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (DatosIgualesF() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosF();
                CargaTablaF();
                dataGridView12_CellClick(dataGridView12, new DataGridViewCellEventArgs(0, CeldaF));
            }
        }

        public void GuardaDatosF()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From formatos";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaF = new DataTable();
                        adapter.Fill(SalvaF);
                    }
                    query = "delete From formatos";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowF();
                    comando = new MySqlCommand("insert into formatos values (?CF, ?NF, ?SF, ?DF)", Conexion);
                    foreach (DataGridViewRow row in dataGridView12.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?CF", Convert.ToString(row.Cells["Codi_Formato"].Value));
                        comando.Parameters.AddWithValue("?NF", Convert.ToString(row.Cells["Tipo_Formato"].Value));
                        comando.Parameters.AddWithValue("?SF", Convert.ToString(row.Cells["Strg_Formato"].Value));
                        comando.Parameters.AddWithValue("?DF", Convert.ToString(row.Cells["Deta_Formato"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From formatos";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView12.DataSource = SalvaF;
                        comando = new MySqlCommand("insert into formatos values (?CF, ?NF, ?SF, ?DF)", Conexion);
                        foreach (DataGridViewRow row in dataGridView12.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?CF", Convert.ToString(row.Cells["Codi_Formato"].Value));
                            comando.Parameters.AddWithValue("?NF", Convert.ToString(row.Cells["Tipo_Formato"].Value));
                            comando.Parameters.AddWithValue("?SF", Convert.ToString(row.Cells["Strg_Formato"].Value));
                            comando.Parameters.AddWithValue("?DF", Convert.ToString(row.Cells["Deta_Formato"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView12_CellClick(dataGridView12, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowF()
        {
            DatosRowF[0] = textBox46.Text;
            DatosRowF[1] = comboBox3.Text;
            DatosRowF[2] = textBox49.Text;
            DatosRowF[3] = textBox52.Text;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            DataRow row = DtableF.NewRow();
            DtableF.Rows.Add(row);
            dataGridView12.DataSource = DtableF;
            dataGridView12.Invoke(new Action(() => { dataGridView12.FirstDisplayedScrollingRowIndex = dataGridView12.RowCount - 1; }));
            dataGridView12_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView12.RowCount - 1));
        }

        private void button16_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el formato?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from formatos where Codi_Formato = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowF[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaF();
                    dataGridView12_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                    MessageBox.Show("El formato se ha borrado exitosamente");
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox46_TextChanged(object sender, EventArgs e)
        {
            dataGridView12.SelectedRows[0].Cells[0].Value = textBox46.Text;
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            dataGridView12.SelectedRows[0].Cells[1].Value = comboBox3.Text;
        }

        private void textBox49_TextChanged(object sender, EventArgs e)
        {
            dataGridView12.SelectedRows[0].Cells[2].Value = textBox49.Text;
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            dataGridView12.SelectedRows[0].Cells[3].Value = textBox52.Text;
        }

        #endregion

        #region Numeros importantes

        DataTable DtableNI;
        List<string> DatosRowNI = new List<string>();
        List<string> ClaseDeLlamadaNI = new List<string>();
        List<string> PlanTarifarioNI = new List<string>();
        List<int> CodigoTarifaNI = new List<int>();
        int CeldaNI = 0;
        int CantidadCeldasNI;
        DataTable SalvaNI;

        public void CargaTablaNI()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From numeros_importantes";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableNI = new DataTable();
                        adapter.Fill(DtableNI);
                        dataGridView13.DataSource = DtableNI;
                    }
                    dataGridView13.EnableHeadersVisualStyles = false;
                    dataGridView13.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView13.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView13.Columns[0].HeaderText = "Número marcado";
                    dataGridView13.Columns[1].HeaderText = "Nombre del destino";
                    dataGridView13.Columns[2].HeaderText = "CL.Llamada";
                    dataGridView13.Columns[3].HeaderText = "Plan tarifa";
                    dataGridView13.Columns[4].HeaderText = "Cod tarifa";

                    dataGridView13.Columns[0].Width = 113;
                    dataGridView13.Columns[1].Width = 133;
                    dataGridView13.Columns[2].Width = 89;
                    dataGridView13.Columns[3].Width = 89;
                    dataGridView13.Columns[4].Width = 89;

                    CargaComboNI();
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        public void CargaComboNI()
        {
            if (comboBox16.Items.Count == 0 && comboBox17.Items.Count == 0 && comboBox18.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from clase_llamadas";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;
                        while (lee.Read())
                        {
                            ClaseDeLlamadaNI.Add(lee["Clase_Llamada"].ToString());
                        }
                        lee.Close();
                        PlanTarifarioNI.Add("001");
                        PlanTarifarioNI.Add("002");
                        PlanTarifarioNI.Add("003");

                        query = "select * from plan_tarifario_001";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;
                        while (lee.Read())
                        {
                            CodigoTarifaNI.Add(Convert.ToInt32(lee["Codi_Tarifa"]));
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    ClaseDeLlamadaNI.Sort();
                    PlanTarifarioNI.Sort();
                    CodigoTarifaNI.Sort();

                    for (int i = 0; i < ClaseDeLlamadaNI.Count; i++)
                    {
                        comboBox16.Items.Add(ClaseDeLlamadaNI[i]);
                    }
                    for (int i = 0; i < PlanTarifarioNI.Count; i++)
                    {
                        comboBox17.Items.Add(PlanTarifarioNI[i]);
                    }
                    for (int i = 0; i < CodigoTarifaNI.Count; i++)
                    {
                        comboBox18.Items.Add(CodigoTarifaNI[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }

        private void dataGridView13_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowNI.Count > 0)
                {
                    if (DatosIgualesNI() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView13.Rows[CeldaNI].Selected = true;
                            textBox55.Text = DatosRowNI[0];
                            textBox54.Text = DatosRowNI[1];
                            comboBox16.Text = DatosRowNI[2];
                            comboBox17.Text = DatosRowNI[3];
                            comboBox18.Text = DatosRowNI[4];
                            dataGridView13_CellClick(null, new DataGridViewCellEventArgs(0, CeldaNI));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosNI();
                            CargaTablaNI();
                            CargaCellNI(e);
                        }
                    }
                    else
                    {
                        CargaCellNI(e);
                    }
                }
                else
                {
                    CargaCellNI(e);
                }
            }
        }

        public void CargaCellNI(DataGridViewCellEventArgs e)
        {
            dataGridView13.ClearSelection();
            dataGridView13.Rows[e.RowIndex].Selected = true;
            CeldaNI = e.RowIndex;

            CantidadCeldasNI = dataGridView13.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowNI = new List<string>();

            while (Posicion < CantidadCeldasNI)
            {
                DatosRowNI.Add(dataGridView13.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            textBox55.Text = DatosRowNI[0];
            textBox54.Text = DatosRowNI[1];
            comboBox16.Text = DatosRowNI[2];
            comboBox17.Text = DatosRowNI[3];
            comboBox18.Text = DatosRowNI[4];
        }

        public bool DatosIgualesNI()
        {
            if (DatosRowNI[0].Equals(textBox55.Text) && DatosRowNI[1].Equals(textBox54.Text) && DatosRowNI[2].Equals(comboBox16.Text) && DatosRowNI[3].Equals(comboBox17.Text) && DatosRowNI[4].Equals(comboBox18.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (DatosIgualesNI() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosNI();
                CargaTablaNI();
                dataGridView13_CellClick(dataGridView13, new DataGridViewCellEventArgs(0, CeldaNI));
            }
        }

        public void GuardaDatosNI()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From numeros_importantes";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaNI = new DataTable();
                        adapter.Fill(SalvaNI);
                    }
                    query = "delete From numeros_importantes";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowNI();
                    comando = new MySqlCommand("insert into numeros_importantes values (?NM, ?ND, ?CL, ?PT, ?CT)", Conexion);
                    foreach (DataGridViewRow row in dataGridView13.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?NM", Convert.ToString(row.Cells["Nume_Marcado"].Value));
                        comando.Parameters.AddWithValue("?ND", Convert.ToString(row.Cells["Nomb_Destino"].Value));
                        comando.Parameters.AddWithValue("?CL", Convert.ToString(row.Cells["Clas_Llamada"].Value));
                        comando.Parameters.AddWithValue("?PT", Convert.ToString(row.Cells["Plan_Tarifario"].Value));
                        comando.Parameters.AddWithValue("?CT", Convert.ToString(row.Cells["Codi_Tarifa"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From numeros_importantes";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView13.DataSource = SalvaNI;
                        comando = new MySqlCommand("insert into numeros_importantes values (?NM, ?ND, ?CL, ?PT, ?CT)", Conexion);
                        foreach (DataGridViewRow row in dataGridView13.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?NM", Convert.ToString(row.Cells["Nume_Marcado"].Value));
                            comando.Parameters.AddWithValue("?ND", Convert.ToString(row.Cells["Nomb_Destino"].Value));
                            comando.Parameters.AddWithValue("?CL", Convert.ToString(row.Cells["Clas_Llamada"].Value));
                            comando.Parameters.AddWithValue("?PT", Convert.ToString(row.Cells["Plan_Tarifario"].Value));
                            comando.Parameters.AddWithValue("?CT", Convert.ToString(row.Cells["Codi_Tarifa"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView13_CellClick(dataGridView13, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowNI()
        {
            DatosRowNI[0] = textBox55.Text;
            DatosRowNI[1] = textBox54.Text;
            DatosRowNI[2] = comboBox16.Text;
            DatosRowNI[3] = comboBox17.Text;
            DatosRowNI[4] = comboBox18.Text;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            DataRow row = DtableNI.NewRow();
            DtableNI.Rows.Add(row);
            dataGridView13.DataSource = DtableNI;
            dataGridView13.Invoke(new Action(() => { dataGridView13.FirstDisplayedScrollingRowIndex = dataGridView13.RowCount - 1; }));
            dataGridView13_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView13.RowCount - 1));
        }

        private void button20_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el el número importante?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from numeros_importantes where Nume_Marcado = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowNI[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaNI();
                    dataGridView13_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                    MessageBox.Show("El formato se ha borrado exitosamente");
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            dataGridView13.SelectedRows[0].Cells[0].Value = textBox55.Text;
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            dataGridView13.SelectedRows[0].Cells[1].Value = textBox54.Text;
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView13.SelectedRows[0].Cells[2].Value = comboBox16.Text;
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView13.SelectedRows[0].Cells[3].Value = comboBox17.Text;
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView13.SelectedRows[0].Cells[4].Value = comboBox18.Text;
        }

        #endregion

        #region Instalación

        DataTable DtableINS;
        List<string> DatosRowINS = new List<string>();
        List<string> ClaseDeLlamadaINS = new List<string>();
        List<string> PlanTarifarioINS = new List<string>();
        List<int> CodigoTarifaINS = new List<int>();
        int CeldaINS = 0;
        int CantidadCeldasINS;
        DataTable SalvaINS;

        public void CargaTablaINS()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From parametros";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableINS = new DataTable();
                        adapter.Fill(DtableINS);
                        dataGridView17.DataSource = DtableINS;
                    }
                    dataGridView17.EnableHeadersVisualStyles = false;
                    dataGridView17.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView17.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView17.Columns[0].HeaderText = "parametro";
                    dataGridView17.Columns[1].HeaderText = "seleccion";
                    dataGridView17.Columns[2].HeaderText = "Detalle";

                    dataGridView17.Columns[0].Width = 213;
                    dataGridView17.Columns[1].Width = 213;
                    dataGridView17.Columns[2].Width = 380;
                    
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        private void dataGridView17_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowINS.Count > 0)
                {
                    if (DatosIgualesINS() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView17.Rows[CeldaINS].Selected = true;
                            textBox64.Text = DatosRowINS[0];
                            textBox65.Text = DatosRowINS[1];
                            textBox66.Text = DatosRowINS[2];
                            dataGridView17_CellClick(null, new DataGridViewCellEventArgs(0, CeldaINS));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosINS();
                            CargaTablaINS();
                            CargaCellINS(e);
                        }
                    }
                    else
                    {
                        CargaCellINS(e);
                    }
                }
                else
                {
                    CargaCellINS(e);
                }
            }
        }
        
        public void CargaCellINS(DataGridViewCellEventArgs e)
        {
            dataGridView17.ClearSelection();
            dataGridView17.Rows[e.RowIndex].Selected = true;
            CeldaINS = e.RowIndex;

            CantidadCeldasINS = dataGridView17.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowINS = new List<string>();

            while (Posicion < CantidadCeldasINS)
            {
                DatosRowINS.Add(dataGridView17.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }
            textBox64.Text = DatosRowINS[0];
            textBox65.Text = DatosRowINS[1];
            textBox66.Text = DatosRowINS[2];
        }

        public bool DatosIgualesINS()
        {
            if (DatosRowINS[0].Equals(textBox64.Text) && DatosRowINS[1].Equals(textBox65.Text) && DatosRowINS[2].Equals(textBox66.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button51_Click(object sender, EventArgs e)
        {
            if (DatosIgualesINS() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosINS();
                CargaTablaINS();
                dataGridView17_CellClick(dataGridView17, new DataGridViewCellEventArgs(0, CeldaINS));
            }
        }

        public void GuardaDatosINS()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From parametros";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaINS = new DataTable();
                        adapter.Fill(SalvaINS);
                    }
                    query = "delete From parametros";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowINS();
                    comando = new MySqlCommand("insert into parametros values (?NM, ?ND, ?CL)", Conexion);
                    foreach (DataGridViewRow row in dataGridView17.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?NM", Convert.ToString(row.Cells["parametro"].Value));
                        comando.Parameters.AddWithValue("?ND", Convert.ToString(row.Cells["seleccion"].Value));
                        comando.Parameters.AddWithValue("?CL", Convert.ToString(row.Cells["Detalle"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From parametros";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView17.DataSource = SalvaINS;
                        comando = new MySqlCommand("insert into parametros values (?NM, ?ND, ?CL)", Conexion);
                        foreach (DataGridViewRow row in dataGridView17.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?NM", Convert.ToString(row.Cells["parametro"].Value));
                            comando.Parameters.AddWithValue("?ND", Convert.ToString(row.Cells["seleccion"].Value));
                            comando.Parameters.AddWithValue("?CL", Convert.ToString(row.Cells["Detalle"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView17_CellClick(dataGridView17, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowINS()
        {
            DatosRowINS[0] = textBox64.Text;
            DatosRowINS[1] = textBox65.Text;
            DatosRowINS[2] = textBox66.Text;
        }

        private void textBox64_TextChanged(object sender, EventArgs e)
        {
            dataGridView17.SelectedRows[0].Cells[0].Value = textBox64.Text;
        }

        private void textBox65_TextChanged(object sender, EventArgs e)
        {
            dataGridView17.SelectedRows[0].Cells[1].Value = textBox65.Text;
        }

        private void textBox66_TextChanged(object sender, EventArgs e)
        {
            dataGridView17.SelectedRows[0].Cells[2].Value = textBox66.Text;
        }

        #endregion

        #endregion

        #region Seguridad

        #region Paso

        private void tabControl6_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabControl6.SelectedIndex != 0 && DatosRowP.Count > 0) { dataGridView15_CellClick(dataGridView14, new DataGridViewCellEventArgs(0, CeldaP)); }
            else if (tabControl6.SelectedIndex != 1 && DatosRowU.Count > 0) { dataGridView16_CellClick(dataGridView16, new DataGridViewCellEventArgs(0, CeldaU)); }

            if (e.TabPageIndex == 0) { tabPage25.Controls.Add(button23); CargaTablaP(); dataGridView15_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
            else if (e.TabPageIndex == 1) { tabPage26.Controls.Add(button23); CargaTablaU(); dataGridView16_CellClick(null, new DataGridViewCellEventArgs(0, 0)); }
        }


        private void button23_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Al salir se decartarán los cambios no guardados, desea continuar?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                tabControl1.SelectTab(0);
            }
        }

        #endregion

        #region Perfiles

        DataTable DtableP;
        List<string> DatosRowP = new List<string>();
        int CeldaP = 0;
        int CantidadCeldasP;
        DataTable SalvaP;

        DataTable DtableP2;
        List<string> DatosRowP2 = new List<string>();
        int CeldaP2 = 0;
        int CantidadCeldasP2;
        string TablaP = "";
        DataTable SalvaP2;

        public void CargaTablaP()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From perfiles";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableP = new DataTable();
                        adapter.Fill(DtableP);
                        dataGridView14.DataSource = DtableP;
                    }
                    dataGridView14.EnableHeadersVisualStyles = false;
                    dataGridView14.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView14.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView14.Columns[0].HeaderText = "Código";
                    dataGridView14.Columns[1].HeaderText = "Detalle del perfil";

                    dataGridView14.Columns[0].Width = 100;
                    dataGridView14.Columns[1].Width = 321;

                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        private void dataGridView14_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowP.Count > 0)
                {
                    if (DatosIgualesP() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView14.Rows[CeldaP].Selected = true;
                            textBox53.Text = DatosRowP[0];
                            dataGridView14_CellClick(null, new DataGridViewCellEventArgs(0, CeldaP));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosP1();
                            CargaTablaP();
                            CargaCellP(e);
                        }
                    }
                    else
                    {
                        CargaCellP(e);
                    }
                }
                else
                {
                    CargaCellP(e);
                }
                CargaGrid2(dataGridView14.Rows[e.RowIndex].Cells[1].Value.ToString());
                CeldaP2 = 0;
                dataGridView15_CellClick(null, new DataGridViewCellEventArgs(0, 0));
            }
        }

        private void dataGridView15_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (DatosRowP2.Count > 0)
            {
                if (DatosIgualesP2() == false)
                {
                    GuardaDatosP1();
                    CargaGrid2(TablaP);
                    CargaCellP2(e);
                }
                else
                {
                    CargaCellP2(e);
                }
            }
            else
            {
                CargaCellP2(e);
            }

        }


        public void CargaGrid2(string Opcion)
        {
            label170.Text = "Acceso a perfil: " + Opcion;
            TablaP = Opcion;
            Opcion = "perfil" + Opcion.ToLower().Split(' ')[0];
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From " + Opcion;
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableP2 = new DataTable();
                        adapter.Fill(DtableP2);
                        dataGridView15.DataSource = DtableP2;
                    }
                    dataGridView15.EnableHeadersVisualStyles = false;
                    dataGridView15.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView15.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView15.Columns[0].HeaderText = "Código";
                    dataGridView15.Columns[1].HeaderText = "Accesos del perfil";
                    dataGridView15.Columns[2].HeaderText = "SI/NO";

                    dataGridView15.Columns[0].Width = 50;
                    dataGridView15.Columns[1].Width = 321;
                    dataGridView15.Columns[2].Width = 50;

                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        public void CargaCellP(DataGridViewCellEventArgs e)
        {
            dataGridView14.ClearSelection();
            dataGridView14.Rows[e.RowIndex].Selected = true;
            CeldaP = e.RowIndex;

            CantidadCeldasP = dataGridView14.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowP = new List<string>();

            while (Posicion < CantidadCeldasP)
            {
                DatosRowP.Add(dataGridView14.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            textBox53.Text = DatosRowP[0];

        }

        public void CargaCellP2(DataGridViewCellEventArgs e)
        {
            dataGridView15.ClearSelection();
            dataGridView15.Rows[e.RowIndex].Selected = true;
            CeldaP2 = e.RowIndex;

            CantidadCeldasP2 = dataGridView15.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowP2 = new List<string>();

            while (Posicion < CantidadCeldasP2)
            {
                DatosRowP2.Add(dataGridView15.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }
            comboBox19.Text = DatosRowP2[2];

        }

        public bool DatosIgualesP()
        {
            if (DatosRowP[0].Equals(textBox53.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        public bool DatosIgualesP2()
        {
            if (DatosRowP2[2].Equals(comboBox19.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (DatosIgualesP() == true && DatosIgualesP2() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosP1();
                CargaTablaP();
                CargaGrid2(TablaP);
                dataGridView14_CellClick(dataGridView14, new DataGridViewCellEventArgs(0, CeldaP));
            }
        }

        public void GuardaDatosP1()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From perfiles";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaP = new DataTable();
                        adapter.Fill(SalvaP);
                    }
                    query = "delete From perfiles";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowP();
                    comando = new MySqlCommand("insert into perfiles values (?perfil, ?DetaPerfil)", Conexion);
                    foreach (DataGridViewRow row in dataGridView14.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?perfil", Convert.ToString(row.Cells["perfil"].Value));
                        comando.Parameters.AddWithValue("?DetaPerfil", Convert.ToString(row.Cells["DetaPerfil"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }
                GuardaDatosP2();
                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From perfiles";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView14.DataSource = SalvaP;
                        comando = new MySqlCommand("insert into perfiles values (?perfil, ?DetaPerfil)", Conexion);
                        foreach (DataGridViewRow row in dataGridView14.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?perfil", Convert.ToString(row.Cells["perfil"].Value));
                            comando.Parameters.AddWithValue("?DetaPerfil", Convert.ToString(row.Cells["DetaPerfil"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView14_CellClick(dataGridView14, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void GuardaDatosP2()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From perfil" + TablaP;
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaP2 = new DataTable();
                        adapter.Fill(SalvaP2);
                    }
                    query = "delete From perfil" + TablaP;
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowP2();
                    comando = new MySqlCommand("insert into perfil" + TablaP + " values (?codigo, ?AccesosPerfil, ?SiNo)", Conexion);
                    foreach (DataGridViewRow row in dataGridView15.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?codigo", Convert.ToString(row.Cells["codigo"].Value));
                        comando.Parameters.AddWithValue("?AccesosPerfil", Convert.ToString(row.Cells["AccesosPerfil"].Value));
                        comando.Parameters.AddWithValue("?SiNo", Convert.ToString(row.Cells["SiNo"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From perfil" + TablaP;
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView15.DataSource = SalvaP2;
                        comando = new MySqlCommand("insert into perfil" + TablaP + " values (?codigo, ?AccesosPerfil, ?SiNo)", Conexion);
                        foreach (DataGridViewRow row in dataGridView15.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?codigo", Convert.ToString(row.Cells["codigo"].Value));
                            comando.Parameters.AddWithValue("?AccesosPerfil", Convert.ToString(row.Cells["AccesosPerfil"].Value));
                            comando.Parameters.AddWithValue("?SiNo", Convert.ToString(row.Cells["SiNo"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView15_CellClick(dataGridView15, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowP()
        {
            DatosRowP[0] = textBox53.Text;
        }

        public void CartgaDatosRowP2()
        {
            DatosRowP2[0] = comboBox19.Text;
            DatosRowP2[1] = comboBox19.Text;
            DatosRowP2[2] = comboBox19.Text;
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            dataGridView14.SelectedRows[0].Cells[0].Value = textBox53.Text;
        }

        private void comboBox19_TextChanged(object sender, EventArgs e)
        {
            dataGridView15.SelectedRows[0].Cells[2].Value = comboBox19.Text;
        }

        #endregion

        #region Usuarios

        DataTable DtableU;
        List<string> DatosRowU = new List<string>();
        List<string> perfil = new List<string>();
        int CeldaU = 0;
        int CantidadCeldasU;
        DataTable SalvaU;

        public void CargaTablaU()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From usuarios";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        DtableU = new DataTable();
                        adapter.Fill(DtableU);
                        dataGridView16.DataSource = DtableU;
                    }
                    dataGridView16.EnableHeadersVisualStyles = false;
                    dataGridView16.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView16.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 217, 102);

                    dataGridView16.Columns[0].HeaderText = "ID";
                    dataGridView16.Columns[1].HeaderText = "Nombre de usuario";
                    dataGridView16.Columns[2].HeaderText = "Perfil";
                    dataGridView16.Columns[3].HeaderText = "E-Mail";
                    dataGridView16.Columns[4].HeaderText = "Activo";
                    dataGridView16.Columns[5].HeaderText = "Ingreso";

                    dataGridView16.Columns[0].Width = 70;
                    dataGridView16.Columns[1].Width = 210;
                    dataGridView16.Columns[2].Width = 70;
                    dataGridView16.Columns[3].Width = 210;
                    dataGridView16.Columns[4].Width = 70;
                    dataGridView16.Columns[5].Width = 70;

                    CargaComboU();
                    Conexion.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al leer la base de datos o cargar las tablas!" + "\n" + e.ToString());
            }
        }

        public void CargaComboU()
        {
            if (comboBox20.Items.Count == 0)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "select * from perfiles";

                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        Agrega = false;

                        while (lee.Read())
                        {
                            perfil.Add(lee["perfil"].ToString());
                        }
                    }

                    lee.Close();
                    Conexion.Close();
                    perfil.Sort();

                    for (int i = 0; i < perfil.Count; i++)
                    {
                        comboBox20.Items.Add(perfil[i]);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ha ocurrido un error al leer la base de datos, máS información: \n\n " + e.ToString());
                }
            }
        }


        private void dataGridView16_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (DatosRowU.Count > 0)
                {
                    if (DatosIgualesU() == false)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Desea guardar los cambios antes de salir?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            dataGridView16.Rows[CeldaU].Selected = true;
                            textBox56.Text = DatosRowU[0];
                            textBox57.Text = DatosRowU[1];
                            comboBox20.Text = DatosRowU[2];
                            textBox58.Text = DatosRowU[3];
                            comboBox21.Text = DatosRowU[4];
                            textBox59.Text = DatosRowU[5];
                            dataGridView16_CellClick(null, new DataGridViewCellEventArgs(0, CeldaU));
                            MessageBox.Show("Los cambios se descartaron");
                        }
                        else
                        {
                            GuardaDatosU();
                            CargaTablaU();
                            CargaCellU(e);
                        }
                    }
                    else
                    {
                        CargaCellU(e);
                    }
                }
                else
                {
                    CargaCellU(e);
                }
            }
        }

        public void CargaCellU(DataGridViewCellEventArgs e)
        {
            dataGridView16.ClearSelection();
            dataGridView16.Rows[e.RowIndex].Selected = true;
            CeldaU = e.RowIndex;

            CantidadCeldasU = dataGridView16.GetCellCount(DataGridViewElementStates.Selected);
            Posicion = 0;
            DatosRowU = new List<string>();

            while (Posicion < CantidadCeldasU)
            {
                DatosRowU.Add(dataGridView16.Rows[e.RowIndex].Cells[Posicion].Value.ToString());
                Posicion++;
            }

            textBox56.Text = DatosRowU[0];
            textBox57.Text = DatosRowU[1];
            comboBox20.Text = DatosRowU[2];
            textBox58.Text = DatosRowU[3];
            comboBox21.Text = DatosRowU[4];
            textBox59.Text = DatosRowU[5];
        }

        public bool DatosIgualesU()
        {
            if (DatosRowU[0].Equals(textBox56.Text) && DatosRowU[1].Equals(textBox57.Text) && DatosRowU[2].Equals(comboBox20.Text)
                && DatosRowU[3].Equals(textBox58.Text) && DatosRowU[4].Equals(comboBox21.Text) && DatosRowU[5].Equals(textBox59.Text))
            {
                return (true);
            }
            else
            {
                return (false);
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (DatosIgualesU() == true)
            {
                MessageBox.Show("No se han detectado cambios");
            }
            else
            {
                GuardaDatosU();
                CargaTablaU();
                dataGridView16_CellClick(dataGridView16, new DataGridViewCellEventArgs(0, CeldaU));
            }
        }

        public void GuardaDatosU()
        {
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From usuarios";
                    using (adapter = new MySqlDataAdapter(query, conexion))
                    {
                        SalvaU = new DataTable();
                        adapter.Fill(SalvaU);
                    }
                    query = "delete From usuarios";
                    comando = new MySqlCommand(query, Conexion);
                    comando.ExecuteNonQuery();
                    CartgaDatosRowU();
                    comando = new MySqlCommand("insert into usuarios values (?id, ?NombreUsuario, ?perfil, ?email, ?activo, ?ingreso)", Conexion);
                    foreach (DataGridViewRow row in dataGridView16.Rows)
                    {
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("?id", Convert.ToString(row.Cells["id"].Value));
                        comando.Parameters.AddWithValue("?NombreUsuario", Convert.ToString(row.Cells["NombreUsuario"].Value));
                        comando.Parameters.AddWithValue("?perfil", Convert.ToString(row.Cells["perfil"].Value));
                        comando.Parameters.AddWithValue("?email", Convert.ToString(row.Cells["E-Mail"].Value));
                        comando.Parameters.AddWithValue("?activo", Convert.ToString(row.Cells["activo"].Value));
                        comando.Parameters.AddWithValue("?ingreso", Convert.ToString(row.Cells["ingreso"].Value));
                        comando.ExecuteNonQuery();
                    }
                    Conexion.Close();
                }

                MessageBox.Show("Los cambios se han guardado");
            }
            catch (Exception e)
            {
                MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + e.ToString());
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();

                        query = "delete From usuarios";
                        comando = new MySqlCommand(query, Conexion);
                        comando.ExecuteNonQuery();

                        dataGridView16.DataSource = SalvaU;
                        comando = new MySqlCommand("insert into usuarios values (?id, ?NombreUsuario, ?perfil, ?email, ?activo, ?ingreso)", Conexion);
                        foreach (DataGridViewRow row in dataGridView16.Rows)
                        {
                            comando.Parameters.Clear();
                            comando.Parameters.AddWithValue("?id", Convert.ToString(row.Cells["id"].Value));
                            comando.Parameters.AddWithValue("?NombreUsuario", Convert.ToString(row.Cells["NombreUsuario"].Value));
                            comando.Parameters.AddWithValue("?perfil", Convert.ToString(row.Cells["perfil"].Value));
                            comando.Parameters.AddWithValue("?email", Convert.ToString(row.Cells["E-Mail"].Value));
                            comando.Parameters.AddWithValue("?activo", Convert.ToString(row.Cells["activo"].Value));
                            comando.Parameters.AddWithValue("?ingreso", Convert.ToString(row.Cells["ingreso"].Value));
                            comando.ExecuteNonQuery();
                        }
                        Conexion.Close();
                    }
                    MessageBox.Show("Los cambios se han revertido");
                    dataGridView16_CellClick(dataGridView16, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception S)
                {
                    MessageBox.Show("No se ha podido revertir los cambios, ¿Está conectado a la base de datos?" + S.ToString());
                }
            }
        }

        public void CartgaDatosRowU()
        {
            DatosRowU[0] = textBox56.Text;
            DatosRowU[1] = textBox57.Text;
            DatosRowU[2] = comboBox20.Text;
            DatosRowU[3] = textBox58.Text;
            DatosRowU[4] = comboBox21.Text;
            DatosRowU[5] = textBox59.Text;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            DataRow row = DtableU.NewRow();
            DtableU.Rows.Add(row);
            dataGridView16.DataSource = DtableU;
            dataGridView16.Invoke(new Action(() => { dataGridView16.FirstDisplayedScrollingRowIndex = dataGridView16.RowCount - 1; }));
            dataGridView16_CellClick(null, new DataGridViewCellEventArgs(0, dataGridView16.RowCount - 1));
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea borrar el formato?", "Atención", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "delete from usuarios where id = ?CT";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?CT", DatosRowU[0]);
                        comando.ExecuteNonQuery();
                        Conexion.Close();
                    }
                    CargaTablaU();
                    MessageBox.Show("Elñ usuario se ha borrado exitosamente");
                    dataGridView16_CellClick(null, new DataGridViewCellEventArgs(0, 0));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Ha ocurrido un error al actualizar la tabla, más información: \n\n" + n.ToString());
                }
            }
        }

        private void textBox56_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[0].Value = textBox56.Text;
        }

        private void textBox57_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[1].Value = textBox57.Text;
        }

        private void comboBox20_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[2].Value = comboBox20.Text;
        }

        private void textBox58_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[3].Value = textBox58.Text;
        }

        private void comboBox21_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[4].Value = comboBox21.Text;
        }

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            dataGridView16.SelectedRows[0].Cells[5].Value = textBox59.Text;
        }

        #endregion

        #endregion

        #region Reportes

        #region Inicio

        TableLayoutPanel TableCE;
        TableLayoutPanel TableCL;
        TableLayoutPanel TableCC;
        TableLayoutPanel TableT;
        List<CheckBox> TableCEL;
        List<CheckBox> TableCLL;
        List<CheckBox> TableCCL;
        List<CheckBox> TableTL;
        List<string> CheckCE;
        List<string> CheckCL;
        List<string> CheckCC;
        List<string> CheckT;
        int Total = 0;
        string FormatoFecha = "";
        string FormatoHora = "";
        string FechaInicial = "";
        string HoraInicial = "";
        string FechaFinal = "";
        string HoraFinal = "";
        bool FiltrosCorretos = true;
        CheckBox CETodos;
        CheckBox CLTodos;
        CheckBox CCTodos;
        CheckBox TTodos;
        TableLayoutPanel TableCL2;
        List<CheckBox> TableCLL2;
        List<string> CheckCL2;
        CheckBox CLTodos2;
        bool SelecFiltro = true;
        string FormatoFechaFinal = "";
        string FormatoHoraFinal = "";
        int PosicionDiaI = 0;
        int PosicionDiaF = 0;
        int TotalRegistros = 0;
        int RegTot = 0;
        string PosDia = "";
        string PosHora = "";
        string PosMinutos = "";
        string FechaRow = "";
        string HoraRow = "";
        string NombreTabla = "";
        bool EsNumerico = false;
        int Out1 = 0;
        int Out2 = 0;
        int Out3 = 0;
        int Out4 = 0;
        int Minutos = 0;
        string HoraI = "";
        string HoraF = "";
        int MesI = 0;
        int MesF = 0;
        bool EnRango = false;
        List<string> TablasNumeros;
        List<string> Filtrados;
        List<string[]> LlamadasFiltradas;
        string[] llamadasFil;
        TableLayoutPanel Visor;
        Label Head;
        List<List<string[]>> EXT = new List<List<string[]>>();
        List<List<string[]>> EXT2 = new List<List<string[]>>();
        List<string[]> ext;
        string CodiCentro;
        int DurGen = 0;
        int VrNetoGen = 0;
        int VrRecargoGen = 0;
        int VrIvaGen = 0;
        int VrTotalGen = 0;
        string[] RowTotal;
        int LOCDur = 0; int LOCTot = 0; int LOCCant = 0;
        int DDNDur = 0; int DDNTot = 0; int DDNCant = 0;
        int CELDur = 0; int CELTot = 0; int CELCant = 0;
        int TOLDur = 0; int TOLTot = 0; int TOLCant = 0;
        int DDIDur = 0; int DDITot = 0; int DDICant = 0;
        int ENTDur = 0; int ENTTot = 0; int ENTCant = 0;
        int EXCDur = 0; int EXCTot = 0; int EXCCant = 0;
        int INTDur = 0; int INTTot = 0; int INTCant = 0;
        int INVDur = 0; int INVTot = 0; int INVCant = 0;
        int ITHDur = 0; int ITHTot = 0; int ITHCant = 0;
        int SATDur = 0; int SATTot = 0; int SATCant = 0;
        int TotalDuracion = 0;
        int TotalValores = 0;
        int TotalCantidad = 0;
        DataGridView DataLlamadas;
        Label lab;
        string LabRes = "";
        string CentroCostoRes = "";
        int CantRes = 0;
        int DurRes = 0;
        int VrNetoRes = 0;
        int VrRecargoRes = 0;
        int VrIVARes = 0;
        int VrTotalRes = 0;
        int IncrementoGen = 10;
        int Posc = 0;
        Document pdfDoc;
        PdfPTable pdfTable;
        PdfPCell cell;
        iTextSharp.text.Font Fuente;
        float[] AnchoPDF;
        int AnchoPDFpos;
        string[] t;
        bool FIltroPrincipalCorrecto = true;

        public void IniciaRep()
        {
            if (panel1.Controls.Count == 0 || panel2.Controls.Count == 0 || panel3.Controls.Count == 0 || panel4.Controls.Count == 0 || panel9.Controls.Count == 0 || panel10.Controls.Count == 0)
            {
                DialogResult dialogResult2 = MessageBox.Show("A continuación se cargarán los filtros, esto puede tardar un poco. ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
                if (dialogResult2 == DialogResult.Yes)
                {
                    FiltrosCorretos = true;

                    try
                    {
                        using (Conexion = new MySqlConnection(conexion))
                        {
                            Conexion.Open();
                            query = "select * from clase_extensiones";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            while (lee.Read())
                            {
                                Total++;
                            }
                            lee.Close();
                            query = "select * from clase_llamadas";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            while (lee.Read())
                            {
                                Total++;
                            }
                            lee.Close();
                            query = "select * from centros_costo";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            while (lee.Read())
                            {
                                Total++;
                            }
                            lee.Close();
                            query = "select * from troncales";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            while (lee.Read())
                            {
                                Total++;
                            }
                            lee.Close();
                            query = "select * from clase_llamadas";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            while (lee.Read())
                            {
                                Total++;
                            }
                            lee.Close();
                            Conexion.Close();
                        }

                        progressBar1.Maximum = Total;
                        progressBar1.Step = 1;
                        progressBar1.Value = 0;
                        Total = 0;

                        tabControl7.Enabled = false;
                        tabControl1.Enabled = false;
                        if (panel2.Controls.Count == 0)
                        {
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from clase_extensiones";
                                    comando = new MySqlCommand(query, Conexion);
                                    lee = comando.ExecuteReader();

                                    TableCE = new TableLayoutPanel();
                                    TableCE.AutoSize = true;
                                    TableCE.Controls.Add(new Label()
                                    {
                                        Text = "CLASE DE EXTENSIONES",
                                        BackColor = System.Drawing.Color.FromArgb(255, 217, 102),
                                        BorderStyle = BorderStyle.FixedSingle,
                                        Font = new System.Drawing.Font("OpenSymbol", 11F, FontStyle.Regular, GraphicsUnit.Point),
                                        Location = new Point(3, 0),
                                        Size = new Size(205, 20),
                                        TextAlign = ContentAlignment.MiddleCenter
                                    }, 1, 0);

                                    TableCE.ColumnCount = 1;
                                    TableCE.Location = new Point(3, 3);
                                    TableCE.RowCount = 1;
                                    TableCE.TabIndex = 0;
                                    TableCE.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset;

                                    panel2.Controls.Add(TableCE);

                                    while (lee.Read())
                                    {
                                        Application.DoEvents();
                                        Total++;
                                        progressBar1.Value = Total;
                                        TableCE.RowCount = TableCE.RowCount + 1;
                                        TableCE.Controls.Add(new CheckBox()
                                        {
                                            Text = lee["Clas_Extension"].ToString() + "   " + lee["Deta_Clase_Extension"].ToString(),
                                            BackColor = System.Drawing.Color.White,
                                            Location = new Point(3, 34),
                                            RightToLeft = RightToLeft.Yes,
                                            Size = new Size(205, 24),
                                            TextAlign = ContentAlignment.MiddleRight,
                                            UseVisualStyleBackColor = false
                                        }, 1, TableCE.RowCount - 1);
                                    }
                                    TableCE.RowCount = TableCE.RowCount + 1;
                                    TableCE.Controls.Add(CETodos = new CheckBox()
                                    {
                                        Text = "TODOS",
                                        BackColor = System.Drawing.Color.White,
                                        Location = new Point(3, 34),
                                        RightToLeft = RightToLeft.Yes,
                                        Size = new Size(205, 24),
                                        TextAlign = ContentAlignment.MiddleRight,
                                        UseVisualStyleBackColor = false,
                                        Name = "CLTodos",

                                    }, 1, TableCE.RowCount - 1);
                                    CETodos.CheckedChanged += CETodos_CheckedChanged;
                                    Conexion.Close();
                                    lee.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                tabControl7.Enabled = true;
                                tabControl1.Enabled = true;
                                MessageBox.Show("No se ha podido leer las clases de extensiones en la base de datos\n\n" + e);
                                FiltrosCorretos = false;
                            }
                        }

                        if (panel1.Controls.Count == 0)
                        {
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from clase_llamadas";
                                    comando = new MySqlCommand(query, Conexion);
                                    lee = comando.ExecuteReader();

                                    TableCL = new TableLayoutPanel();
                                    TableCL.AutoSize = true;
                                    TableCL.Controls.Add(new Label()
                                    {
                                        Text = "CLASE DE LLAMADAS",
                                        BackColor = System.Drawing.Color.FromArgb(255, 217, 102),
                                        BorderStyle = BorderStyle.FixedSingle,
                                        Font = new System.Drawing.Font("OpenSymbol", 11F, FontStyle.Regular, GraphicsUnit.Point),
                                        Location = new Point(3, 0),
                                        Size = new Size(205, 20),
                                        TextAlign = ContentAlignment.MiddleCenter
                                    }, 1, 0);

                                    TableCL.ColumnCount = 1;
                                    TableCL.Location = new Point(3, 3);
                                    TableCL.RowCount = 1;
                                    TableCL.TabIndex = 0;
                                    TableCL.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset;

                                    panel1.Controls.Add(TableCL);

                                    while (lee.Read())
                                    {
                                        Application.DoEvents();
                                        Total++;
                                        progressBar1.Value = Total;
                                        TableCL.RowCount = TableCL.RowCount + 1;
                                        TableCL.Controls.Add(new CheckBox()
                                        {
                                            Text = lee["Clase_Llamada"].ToString() + "   " + lee["Nombre_Clase_Llamada"].ToString(),
                                            BackColor = System.Drawing.Color.White,
                                            Location = new Point(3, 34),
                                            RightToLeft = RightToLeft.Yes,
                                            Size = new Size(205, 23),
                                            TextAlign = ContentAlignment.MiddleRight,
                                            UseVisualStyleBackColor = false
                                        }, 1, TableCL.RowCount - 1);
                                    }
                                    TableCL.RowCount = TableCL.RowCount + 1;
                                    TableCL.Controls.Add(CLTodos = new CheckBox()
                                    {
                                        Text = "TODOS",
                                        BackColor = System.Drawing.Color.White,
                                        Location = new Point(3, 34),
                                        RightToLeft = RightToLeft.Yes,
                                        Size = new Size(205, 23),
                                        TextAlign = ContentAlignment.MiddleRight,
                                        UseVisualStyleBackColor = false,
                                        Name = "CLTodos"
                                    }, 1, TableCL.RowCount - 1);
                                    CLTodos.CheckedChanged += CLTodos_CheckedChanged;
                                    Conexion.Close();
                                    lee.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                tabControl7.Enabled = true;
                                tabControl1.Enabled = true;
                                MessageBox.Show("No se ha podido leer las clases de extensiones en la base de datos\n\n" + e);
                                FiltrosCorretos = false;
                            }
                        }
                        if (panel3.Controls.Count == 0)
                        {
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from centros_costo";
                                    comando = new MySqlCommand(query, Conexion);
                                    lee = comando.ExecuteReader();

                                    TableCC = new TableLayoutPanel();
                                    TableCC.AutoSize = true;
                                    TableCC.Controls.Add(new Label()
                                    {
                                        Text = "CENTRO DE COSTO",
                                        BackColor = System.Drawing.Color.FromArgb(255, 217, 102),
                                        BorderStyle = BorderStyle.FixedSingle,
                                        Font = new System.Drawing.Font("OpenSymbol", 11F, FontStyle.Regular, GraphicsUnit.Point),
                                        Location = new Point(3, 0),
                                        Size = new Size(205, 20),
                                        TextAlign = ContentAlignment.MiddleCenter
                                    }, 1, 0);

                                    TableCC.ColumnCount = 1;
                                    TableCC.Location = new Point(3, 3);
                                    TableCC.RowCount = 1;
                                    TableCC.TabIndex = 0;
                                    TableCC.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset;

                                    panel3.Controls.Add(TableCC);

                                    while (lee.Read())
                                    {
                                        Application.DoEvents();
                                        Total++;
                                        progressBar1.Value = Total;
                                        TableCC.RowCount = TableCC.RowCount + 1;
                                        TableCC.Controls.Add(new CheckBox()
                                        {
                                            Text = lee["Codi_Centro"].ToString() + "   " + lee["Nomb_Centro"].ToString(),
                                            BackColor = System.Drawing.Color.White,
                                            Location = new Point(3, 34),
                                            RightToLeft = RightToLeft.Yes,
                                            Size = new Size(205, 23),
                                            TextAlign = ContentAlignment.MiddleRight,
                                            UseVisualStyleBackColor = false
                                        }, 1, TableCC.RowCount - 1);
                                    }
                                    TableCC.RowCount = TableCC.RowCount + 1;
                                    TableCC.Controls.Add(CCTodos = new CheckBox()
                                    {
                                        Text = "TODOS",
                                        BackColor = System.Drawing.Color.White,
                                        Location = new Point(3, 34),
                                        RightToLeft = RightToLeft.Yes,
                                        Size = new Size(205, 23),
                                        TextAlign = ContentAlignment.MiddleRight,
                                        UseVisualStyleBackColor = false,
                                        Name = "CLTodos"
                                    }, 1, TableCC.RowCount - 1);
                                    CCTodos.CheckedChanged += CCTodos_CheckedChanged;
                                    Conexion.Close();
                                    lee.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                tabControl7.Enabled = true;
                                tabControl1.Enabled = true;
                                MessageBox.Show("No se ha podido leer las clases de extensiones en la base de datos\n\n" + e);
                                FiltrosCorretos = false;
                            }
                        }
                        if (panel4.Controls.Count == 0)
                        {
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from troncales";
                                    comando = new MySqlCommand(query, Conexion);
                                    lee = comando.ExecuteReader();

                                    TableT = new TableLayoutPanel();
                                    TableT.AutoSize = true;
                                    TableT.Controls.Add(new Label()
                                    {
                                        Text = "TRONCALES",
                                        BackColor = System.Drawing.Color.FromArgb(255, 217, 102),
                                        BorderStyle = BorderStyle.FixedSingle,
                                        Font = new System.Drawing.Font("OpenSymbol", 11F, FontStyle.Regular, GraphicsUnit.Point),
                                        Location = new Point(3, 0),
                                        Size = new Size(205, 20),
                                        TextAlign = ContentAlignment.MiddleCenter
                                    }, 1, 0);

                                    TableT.ColumnCount = 1;
                                    TableT.Location = new Point(3, 3);
                                    TableT.RowCount = 1;
                                    TableT.TabIndex = 0;
                                    TableT.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset;

                                    panel4.Controls.Add(TableT);

                                    while (lee.Read())
                                    {
                                        Application.DoEvents();
                                        Total++;
                                        progressBar1.Value = Total;
                                        TableT.RowCount = TableT.RowCount + 1;
                                        TableT.Controls.Add(new CheckBox()
                                        {
                                            Text = lee["Line_Troncal"].ToString() + "   " + lee["Nume_Acceso_Troncal"].ToString(),
                                            BackColor = System.Drawing.Color.White,
                                            Location = new Point(3, 34),
                                            RightToLeft = RightToLeft.Yes,
                                            Size = new Size(205, 23),
                                            TextAlign = ContentAlignment.MiddleRight,
                                            UseVisualStyleBackColor = false
                                        }, 1, TableT.RowCount - 1);
                                    }
                                    TableT.RowCount = TableT.RowCount + 1;
                                    TableT.Controls.Add(TTodos = new CheckBox()
                                    {
                                        Text = "TODOS",
                                        BackColor = System.Drawing.Color.White,
                                        Location = new Point(3, 34),
                                        RightToLeft = RightToLeft.Yes,
                                        Size = new Size(205, 23),
                                        TextAlign = ContentAlignment.MiddleRight,
                                        UseVisualStyleBackColor = false,
                                        Name = "CLTodos"
                                    }, 1, TableT.RowCount - 1);
                                    TTodos.CheckedChanged += TTodos_CheckedChanged;
                                    Conexion.Close();
                                    lee.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                tabControl7.Enabled = true;
                                tabControl1.Enabled = true;
                                MessageBox.Show("No se ha podido leer las clases de extensiones en la base de datos\n\n" + e);
                                FiltrosCorretos = false;
                            }
                        }
                        if (panel10.Controls.Count == 0)
                        {
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from clase_llamadas";
                                    comando = new MySqlCommand(query, Conexion);
                                    lee = comando.ExecuteReader();

                                    TableCL2 = new TableLayoutPanel();
                                    TableCL2.AutoSize = true;
                                    TableCL2.Controls.Add(new Label()
                                    {
                                        Text = "CLASE DE LLAMADAS",
                                        BackColor = System.Drawing.Color.FromArgb(255, 217, 102),
                                        BorderStyle = BorderStyle.FixedSingle,
                                        Font = new System.Drawing.Font("OpenSymbol", 11F, FontStyle.Regular, GraphicsUnit.Point),
                                        Location = new Point(3, 0),
                                        Size = new Size(205, 20),
                                        TextAlign = ContentAlignment.MiddleCenter
                                    }, 1, 0);

                                    TableCL2.ColumnCount = 1;
                                    TableCL2.Location = new Point(3, 3);
                                    TableCL2.RowCount = 1;
                                    TableCL2.TabIndex = 0;
                                    TableCL2.CellBorderStyle = TableLayoutPanelCellBorderStyle.Outset;

                                    panel10.Controls.Add(TableCL2);

                                    while (lee.Read())
                                    {
                                        Application.DoEvents();
                                        Total++;
                                        progressBar1.Value = Total;
                                        TableCL2.RowCount = TableCL2.RowCount + 1;
                                        TableCL2.Controls.Add(new CheckBox()
                                        {
                                            Text = lee["Clase_Llamada"].ToString() + "   " + lee["Nombre_Clase_Llamada"].ToString(),
                                            BackColor = System.Drawing.Color.White,
                                            Location = new Point(3, 34),
                                            RightToLeft = RightToLeft.Yes,
                                            Size = new Size(205, 24),
                                            TextAlign = ContentAlignment.MiddleRight,
                                            UseVisualStyleBackColor = false
                                        }, 1, TableCL2.RowCount - 1);
                                    }
                                    TableCL2.RowCount = TableCL2.RowCount + 1;
                                    TableCL2.Controls.Add(CLTodos2 = new CheckBox()
                                    {
                                        Text = "TODOS",
                                        BackColor = System.Drawing.Color.White,
                                        Location = new Point(3, 34),
                                        RightToLeft = RightToLeft.Yes,
                                        Size = new Size(205, 24),
                                        TextAlign = ContentAlignment.MiddleRight,
                                        UseVisualStyleBackColor = false,
                                        Name = "CLTodos",

                                    }, 1, TableCL2.RowCount - 1);
                                    CLTodos2.CheckedChanged += CLTodos2_CheckedChanged;
                                    Conexion.Close();
                                    lee.Close();
                                }
                            }
                            catch (Exception ex)
                            {
                                tabControl7.Enabled = true;
                                tabControl1.Enabled = true;
                                MessageBox.Show("No se ha podido leer las clases de llamadas en la base de datos\n\n" + ex);
                                FiltrosCorretos = false;
                            }
                        }
                        if (string.IsNullOrEmpty(FormatoFechaFinal))
                        {
                            Envia("04");
                        }
                        if (string.IsNullOrEmpty(FormatoHoraFinal))
                        {
                            Envia("06");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocurrió un problema al cargar los filtros!\n\n" + ex.ToString());
                        FiltrosCorretos = false;
                    }
                    if (FiltrosCorretos == true)
                    {
                        CargaCheckBox();
                        label188.Visible = false;
                        progressBar1.Visible = false;
                        button46.Visible = true;
                        button46.Enabled = true;
                    }
                    else
                    {
                        panel1.Controls.Clear();
                        panel2.Controls.Clear();
                        panel3.Controls.Clear();
                        panel4.Controls.Clear();
                        panel10.Controls.Clear();
                        tabControl1.SelectTab(0);
                    }
                    tabControl7.Enabled = true;
                    tabControl1.Enabled = true;
                }
                else if (dialogResult2 == DialogResult.No)
                {
                    tabControl1.SelectTab(0);
                }
            }
        }

        private void CETodos_CheckedChanged(object sender, EventArgs e)
        {
            if (CETodos.Checked == true)
            {
                foreach (CheckBox s in TableCEL)
                {
                    s.Checked = true;
                }
            }
            else
            {
                foreach (CheckBox s in TableCEL)
                {
                    s.Checked = false;
                }
            }
        }

        private void CLTodos_CheckedChanged(object sender, EventArgs e)
        {
            if (CLTodos.Checked == true)
            {
                foreach (CheckBox s in TableCLL)
                {
                    s.Checked = true;
                }
            }
            else
            {
                foreach (CheckBox s in TableCLL)
                {
                    s.Checked = false;
                }
            }
        }

        private void CCTodos_CheckedChanged(object sender, EventArgs e)
        {
            if (CCTodos.Checked == true)
            {
                foreach (CheckBox s in TableCCL)
                {
                    s.Checked = true;
                }
            }
            else
            {
                foreach (CheckBox s in TableCCL)
                {
                    s.Checked = false;
                }
            }
        }

        private void TTodos_CheckedChanged(object sender, EventArgs e)
        {
            if (TTodos.Checked == true)
            {
                foreach (CheckBox s in TableTL)
                {
                    s.Checked = true;
                }
            }
            else
            {
                foreach (CheckBox s in TableTL)
                {
                    s.Checked = false;
                }
            }
        }
        private void CLTodos2_CheckedChanged(object sender, EventArgs e)
        {
            if (CLTodos2.Checked == true)
            {
                foreach (CheckBox s in TableCLL2)
                {
                    s.Checked = true;
                }
            }
            else
            {
                foreach (CheckBox s in TableCLL2)
                {
                    s.Checked = false;
                }
            }
        }

        public void CargaCheckBox()
        {
            TableCEL = new List<CheckBox>();
            TableCLL = new List<CheckBox>();
            TableCCL = new List<CheckBox>();
            TableTL = new List<CheckBox>();

            foreach (System.Windows.Forms.Control s in TableCE.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableCEL.Add((CheckBox)s);
                }
            }
            foreach (System.Windows.Forms.Control s in TableCL.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableCLL.Add((CheckBox)s);
                }
            }
            foreach (System.Windows.Forms.Control s in TableCC.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableCCL.Add((CheckBox)s);
                }
            }
            foreach (System.Windows.Forms.Control s in TableT.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableTL.Add((CheckBox)s);
                }
            }

            TableCLL2 = new List<CheckBox>();
            foreach (System.Windows.Forms.Control s in TableCL2.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableCLL2.Add((CheckBox)s);
                }
            }
        }

        #endregion

        #region Reportes Generales

        #region Procesa

        private void button46_Click(object sender, EventArgs e)
        {
            if (Correcto() == true)
            {
                if (SePuede() == true)
                {
                    FechaInicial = dateTimePicker1.Text.ToString();
                    HoraInicial = dateTimePicker3.Text.ToString();
                    FechaFinal = dateTimePicker2.Text.ToString();
                    HoraFinal = dateTimePicker4.Text.ToString();
                    DialogResult dialogResult2 = MessageBox.Show("Desde: " + FechaInicial + " a las: " + HoraInicial + " Hasta: " + FechaFinal + " a las: " + HoraFinal + ". ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
                    if (dialogResult2 == DialogResult.Yes)
                    {
                        FIltroPrincipalCorrecto = true;
                        button46.Enabled = false;
                        button46.Visible = false;
                        label188.Text = "Procesando...";
                        label188.Visible = true;
                        progressBar1.Value = 0;
                        progressBar1.Maximum = 100;
                        progressBar1.Visible = true;
                        CargaChecked();
                        LeePos();
                        CargaCheckBox();
                        using (Conexion = new MySqlConnection(conexion))
                        {
                            Conexion.Open();
                            query = "SHOW TABLES";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            TablasNumeros = new List<string>();
                            while (lee.Read())
                            {
                                NombreTabla = "";
                                EsNumerico = false;
                                NombreTabla = lee.GetValue(0).ToString();
                                EsNumerico = int.TryParse(NombreTabla.Split(' ')[0], out Out1);
                                if (EsNumerico == true)
                                {
                                    TablasNumeros.Add(NombreTabla);
                                }
                            }
                            lee.Close();
                            Conexion.Close();
                        }
                        if (TablasNumeros.Count >= 1)
                        {
                            Out1 = Convert.ToInt32(dateTimePicker1.Value.ToString("yyyy"));
                            Out2 = Convert.ToInt32(dateTimePicker2.Value.ToString("yyyy"));
                            Out3 = Convert.ToInt32(dateTimePicker1.Value.ToString("MM"));
                            Out4 = Convert.ToInt32(dateTimePicker2.Value.ToString("MM"));
                            Filtrados = new List<string>();
                            foreach (string S in TablasNumeros)
                            {
                                EnRango = true;

                                if (Convert.ToInt32(S.Split(' ')[0]) >= Out1 && Convert.ToInt32(S.Split(' ')[0]) <= Out2)
                                {
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out1)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) < Out3)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out2)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) > Out4)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                }
                                else
                                {
                                    EnRango = false;
                                }

                                if (EnRango == true)
                                {
                                    Filtrados.Add(S);
                                }
                            }
                            if (Filtrados.Count >= 1)
                            {
                                TablasNumeros = Filtrados;
                                if (TablasNumeros.Count == 1)
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        Out1 = Convert.ToInt32(dateTimePicker1.Value.ToString("dd"));
                                        Out2 = Convert.ToInt32(dateTimePicker2.Value.ToString("dd"));
                                        Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                        Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                        HoraI = dateTimePicker3.Value.ToString("HH:mm");
                                        HoraF = dateTimePicker4.Value.ToString("HH:mm");
                                        LlamadasFiltradas = new List<string[]>();
                                        TotalRegistros = 0;
                                        while (lee.Read())
                                        {
                                            TotalRegistros++;
                                        }
                                        lee.Close();
                                        RegTot = TotalRegistros;
                                        TotalRegistros = 0;
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        while (lee.Read())
                                        {
                                            Application.DoEvents();
                                            Minutos = 0;
                                            if (lee["Errores"].ToString().Equals("-"))
                                            {
                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1 && Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                {
                                                    if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1 && Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                    {
                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])) && Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[12];
                                                            FiltraCheck();
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                    {
                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[12];
                                                            FiltraCheck();
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                    {

                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[12];
                                                            FiltraCheck();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        llamadasFil = new string[12];
                                                        FiltraCheck();
                                                    }


                                                }
                                            }
                                            TotalRegistros++;
                                            progressBar1.Value = (int)((TotalRegistros * 100) / RegTot);
                                        }
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                else
                                {
                                    TotalRegistros = 0;
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        using (Conexion = new MySqlConnection(conexion))
                                        {
                                            Conexion.Open();
                                            query = "select * from `" + TablasNumeros[i] + "`";
                                            comando = new MySqlCommand(query, Conexion);
                                            lee = comando.ExecuteReader();
                                            while (lee.Read())
                                            {
                                                TotalRegistros++;
                                            }
                                            lee.Close();
                                            Conexion.Close();
                                        }
                                    }
                                    progressBar1.Maximum = 100;
                                    RegTot = TotalRegistros;
                                    TotalRegistros = 0;
                                    LlamadasFiltradas = new List<string[]>();
                                    Out1 = Convert.ToInt32(dateTimePicker1.Value.ToString("dd"));
                                    Out2 = Convert.ToInt32(dateTimePicker2.Value.ToString("dd"));
                                    Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                    Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                    HoraI = dateTimePicker3.Value.ToString("HH:mm");
                                    HoraF = dateTimePicker4.Value.ToString("HH:mm");
                                    MesI = Convert.ToInt32(dateTimePicker1.Value.ToString("MM"));
                                    MesF = Convert.ToInt32(dateTimePicker2.Value.ToString("MM"));
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        if (i == 0)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker1.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                                    {
                                                                        llamadasFil = new string[12];
                                                                        FiltraCheck();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    llamadasFil = new string[12];
                                                                    FiltraCheck();
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            llamadasFil = new string[12];
                                                            FiltraCheck();
                                                        }
                                                    }
                                                    
                                                    TotalRegistros++;
                                                    progressBar1.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else if (i == TablasNumeros.Count - 1)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker2.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                                    {
                                                                        llamadasFil = new string[12];
                                                                        FiltraCheck();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    llamadasFil = new string[12];
                                                                    FiltraCheck();
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            llamadasFil = new string[12];
                                                            FiltraCheck();
                                                        }
                                                    }
                                                    
                                                    TotalRegistros++;
                                                    progressBar1.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from " + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        llamadasFil = new string[12];
                                                        FiltraCheck();
                                                    }
                                                    TotalRegistros++;
                                                    progressBar1.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("No se han detectado llamadas dentro del rango especificado");
                                FIltroPrincipalCorrecto = false;
                                TerminaReporte();
                            }
                        }
                        else
                        {
                            MessageBox.Show("No se han detectado llamadas en la base de datos");
                            FIltroPrincipalCorrecto = false;
                            TerminaReporte();
                        }
                        if(FIltroPrincipalCorrecto == true)
                        {
                            TerminaReporte();
                            MuestraReporte(Directory.GetCurrentDirectory() + @"\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".pdf");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No se ha detectado una fecha en la cadena");
                    TerminaReporte();
                }
            }
        }

        bool FiltroParametro = true;
        public void FiltraCheck()
        {
            try
            {
                for (int i = 0; i < CheckCE.Count; i++)
                {
                    if (CheckCE[i].Equals(lee["ClaseExtension"].ToString()))
                    {
                        i = CheckCE.Count;
                        for (int n = 0; n < CheckCL.Count; n++)
                        {
                            if (CheckCL[n].Equals(lee["ClaseLlamada"].ToString()))
                            {
                                n = CheckCL.Count;
                                for (int m = 0; m < CheckCC.Count; m++)
                                {
                                    if (CheckCC[m].Equals(lee["CentroDeCosto"].ToString()))
                                    {
                                        m = CheckCC.Count;
                                        for (int s = 0; s < CheckT.Count; s++)
                                        {
                                            if (CheckT[s].Equals(lee["TTroncal"].ToString()))
                                            {
                                                FiltroParametro = true;
                                                if (checkBox2.Checked == true)
                                                {
                                                    using (Conexion = new MySqlConnection(conexion))
                                                    {
                                                        Conexion.Open();
                                                        query = "select * from parametros where parametro = 'Llamadas extensas general'";
                                                        comando2 = new MySqlCommand(query, Conexion);
                                                        lee2 = comando2.ExecuteReader();
                                                        lee2.Read();
                                                        if ((Convert.ToInt32(lee["DuracionLlamadaAproximada"].ToString())) < (Convert.ToInt32(lee2["seleccion"].ToString())))
                                                        {
                                                            FiltroParametro = false;
                                                        }
                                                        lee2.Close();
                                                        Conexion.Close();
                                                    }
                                                }
                                                if (checkBox3.Checked == true)
                                                {
                                                    using (Conexion = new MySqlConnection(conexion))
                                                    {
                                                        Conexion.Open();
                                                        query = "select * from parametros where parametro = 'Llamadas con valor general'";
                                                        comando2 = new MySqlCommand(query, Conexion);
                                                        lee2 = comando2.ExecuteReader();
                                                        lee2.Read();
                                                        if ((Convert.ToInt32(lee["ValorTotal"].ToString())) < (Convert.ToInt32(lee2["seleccion"].ToString())))
                                                        {
                                                            FiltroParametro = false;
                                                        }
                                                        lee2.Close();
                                                        Conexion.Close();
                                                    }
                                                    
                                                }

                                                if (FiltroParametro == true)
                                                {
                                                    s = CheckT.Count;
                                                    llamadasFil[0] = lee["FFechaFinalLlamada"].ToString();
                                                    llamadasFil[1] = lee["HHoraFinalLlamada"].ToString();
                                                    llamadasFil[2] = lee["NNumeroMarcado"].ToString();
                                                    llamadasFil[3] = lee["Destino"].ToString();
                                                    llamadasFil[4] = lee["ClaseLlamada"].ToString();
                                                    llamadasFil[5] = lee["DuracionLlamadaAproximada"].ToString();
                                                    llamadasFil[6] = lee["ValorLlamadaTarifa"].ToString();
                                                    llamadasFil[7] = lee["RecargoServicioValor"].ToString();
                                                    llamadasFil[8] = lee["ValorIVA"].ToString();
                                                    llamadasFil[9] = lee["ValorTotal"].ToString();
                                                    llamadasFil[10] = lee["EExtension"].ToString();
                                                    try
                                                    {
                                                        using (Conexion = new MySqlConnection(conexion))
                                                        {
                                                            Conexion.Open();
                                                            query = "select * from extensiones where Nume_Extension = '" + llamadasFil[10] + "'";
                                                            comando2 = new MySqlCommand(query, Conexion);
                                                            lee2 = comando2.ExecuteReader();
                                                            lee2.Read();
                                                            llamadasFil[11] = lee["Nomb_Extension"].ToString();
                                                            lee2.Close();
                                                            Conexion.Close();
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        llamadasFil[11] = "Extension desconocida: " + llamadasFil[10];
                                                    }
                                                    LlamadasFiltradas.Add(llamadasFil);
                                                    
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al aplicar los filtros!\n\n" + ex.ToString());
                FIltroPrincipalCorrecto = false;
            }
        }
       
        bool Repetido = false;
        public void MuestraReporte(string np)
        {
            using (FileStream stream = new FileStream(np, FileMode.Create))
            {
                MP.label1.Text = "Cargando reporte al visor, por favor espere...";
                MP.Show();
                Application.DoEvents();

                pdfDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                Fuente = new iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7, iTextSharp.text.Font.NORMAL);

                Filtrados = new List<string>();
                Head = new Label();
                Head.MaximumSize = new Size(720, 454);
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from parametros where parametro = 'Reportes Hotel'";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        lee.Read();
                        Head.Text = lee["seleccion"].ToString();
                        lee.Close();
                        Conexion.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problema al conectarse con la base de datos\n\n" + ex.ToString());
                }
                Head.Text += "\n\nREPORTE GENERAL DE LLAMADAS" + "\nDesde: " + FechaInicial + " a las: " + HoraInicial + " Hasta: " + FechaFinal + " a las: " + HoraFinal + "\n\n Extensiones: ";
                if (CheckCE.Count == TableCEL.Count)
                {
                    Head.Text += "Todos";
                }
                else
                {
                    foreach (string s in CheckCE)
                    {
                        Head.Text += s + ", ";
                    }
                }
                Head.Text += "\nCentros de costo: ";
                if (CheckCC.Count == TableCCL.Count)
                {
                    Head.Text += "Todos";
                }
                else
                {
                    foreach (string s in CheckCC)
                    {
                        Head.Text += s + ", ";
                    }
                }
                Head.Text += "\nClases de llamada: ";
                if (CheckCL.Count == TableCLL.Count)
                {
                    Head.Text += "Todos";
                }
                else
                {
                    foreach (string s in CheckCL)
                    {
                        Head.Text += s + ", ";
                    }
                }
                Head.Text += "\nTroncales: ";
                if (CheckT.Count == TableTL.Count)
                {
                    Head.Text += "Todos";
                }
                else
                {
                    foreach (string s in CheckT)
                    {
                        Head.Text += s + ", ";
                    }
                }
                if (checkBox1.Checked == true)
                {
                    Head.Text += "\n\nREPORTE RESUMIDO\n\n";
                }
                else
                {
                    Head.Text += "\n\nREPORTE DETALLADO\n\n";
                }
                Head.AutoSize = true;
                Head.BorderStyle = BorderStyle.None;

                Visor = new TableLayoutPanel();
                Visor.Controls.Add(Head);
                pdfDoc.Add(new Paragraph(Head.Text, Fuente));
                Visor.ColumnCount = 1;
                Visor.Location = new Point(0, 0);
                Visor.RowCount = 1;
                Visor.TabIndex = 0;
                Visor.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
                Visor.AutoSize = true;

                Visor.RowCount = Visor.RowCount + 1;
                EXT = new List<List<string[]>>();
                ext = new List<string[]>();

                for (int a = 1; a < LlamadasFiltradas.Count; a++)
                {
                    for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                    {
                        if (Convert.ToInt32(LlamadasFiltradas[b - 1][10]) > Convert.ToInt32(LlamadasFiltradas[b][10]))
                        {
                            t = LlamadasFiltradas[b - 1];
                            LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                            LlamadasFiltradas[b] = t;
                        }
                    }
                }
                    

                foreach (string[] s in LlamadasFiltradas)
                {
                    Repetido = false;
                    if (EXT.Count > 0)
                    {
                        foreach (List<string[]> l in EXT)
                        {
                            foreach (string[] c in l)
                            {
                                if (c[10].Equals(s[10]))
                                {
                                    Repetido = true;
                                }
                            }
                        }
                        if (Repetido == false)
                        {
                            foreach (string[] n in LlamadasFiltradas)
                            {
                                if (n[10].Equals(s[10]))
                                {
                                    ext.Add(n);
                                }
                            }
                            if (ext.Count > 0)
                            {
                                EXT.Add(ext);
                                ext = new List<string[]>();
                            }
                        }
                    }
                    else
                    {
                        foreach (string[] n in LlamadasFiltradas)
                        {
                            if (n[10].Equals(s[10]))
                            {
                                ext.Add(n);
                            }
                        }
                        if (ext.Count > 0)
                        {
                            EXT.Add(ext);
                            ext = new List<string[]>();
                        }
                    }
                }
                
                LOCDur = 0; LOCTot = 0; LOCCant = 0;
                DDNDur = 0; DDNTot = 0; DDNCant = 0;
                CELDur = 0; CELTot = 0; CELCant = 0;
                TOLDur = 0; TOLTot = 0; TOLCant = 0;
                DDIDur = 0; DDITot = 0; DDICant = 0;
                ENTDur = 0; ENTTot = 0; ENTCant = 0;
                EXCDur = 0; EXCTot = 0; EXCCant = 0;
                INTDur = 0; INTTot = 0; INTCant = 0;
                INVDur = 0; INVTot = 0; INVCant = 0;
                ITHDur = 0; ITHTot = 0; ITHCant = 0;
                SATDur = 0; SATTot = 0; SATCant = 0;
                TotalValores = 0;
                TotalDuracion = 0;
                TotalCantidad = 0;

                if (checkBox1.Checked == false)
                {
                    foreach (List<string[]> s in EXT)
                    {
                        Visor.RowCount = Visor.RowCount + 1;
                        lab = new Label();
                        lab.Text = "EXT: " + (s[0][10]) + "    ";
                        try
                        {
                            using (Conexion = new MySqlConnection(conexion))
                            {
                                Conexion.Open();
                                query = "select * from extensiones where Nume_Extension = ?e";
                                comando = new MySqlCommand(query, Conexion);
                                comando.Parameters.AddWithValue("?e", (s[0][10]));
                                lee = comando.ExecuteReader();
                                lee.Read();
                                lab.Text += lee["Nomb_Extension"].ToString();
                                CodiCentro = lee["Codi_Centro"].ToString();
                                lab.Text += "           CENTRO: " + CodiCentro + " ";
                                lee.Close();
                                Conexion.Close();
                                try
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from centros_costo where Codi_Centro = ?e";
                                        comando = new MySqlCommand(query, Conexion);
                                        comando.Parameters.AddWithValue("?e", CodiCentro);
                                        lee = comando.ExecuteReader();
                                        lee.Read();
                                        lab.Text += lee["Nomb_Centro"].ToString();
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                catch
                                {
                                    lab.Text += "Nombre de centro desconocido";
                                }
                            }
                        }
                        catch
                        {
                            lab.Text += "Extension desconocida";
                            lab.Text += "Centro de costo desconocido";
                        }
                        lab.AutoSize = true;
                        lab.BorderStyle = BorderStyle.None;
                        Visor.Controls.Add(lab);
                        pdfDoc.Add(new Paragraph(lab.Text, Fuente));
                        Visor.RowCount = Visor.RowCount + 1;
                        DataLlamadas = new DataGridView();
                        Visor.Controls.Add(DataLlamadas);
                        DataLlamadas.ColumnCount = 10;
                        DataLlamadas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        DataLlamadas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        DataLlamadas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        DataLlamadas.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        DataLlamadas.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        DataLlamadas.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        DataLlamadas.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        DataLlamadas.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        DataLlamadas.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        DataLlamadas.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        DataLlamadas.AutoSize = true;
                        DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                        DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                        DataLlamadas.BorderStyle = BorderStyle.None;
                        DataLlamadas.AllowUserToAddRows = false;
                        DataLlamadas.AllowUserToDeleteRows = false;
                        DataLlamadas.AllowUserToResizeRows = false;
                        DataLlamadas.RowHeadersVisible = false;
                        DataLlamadas.MultiSelect = false;
                        DataLlamadas.ReadOnly = true;
                        DataLlamadas.Enabled = false;

                        DataLlamadas.Columns[0].HeaderText = "FECHA";
                        DataLlamadas.Columns[0].Width = 57;
                        DataLlamadas.Columns[1].HeaderText = "HORA";
                        DataLlamadas.Columns[1].Width = 57;
                        DataLlamadas.Columns[2].HeaderText = "NUM.MARCADO";
                        DataLlamadas.Columns[2].Width = 110;
                        DataLlamadas.Columns[3].HeaderText = "DESTINO";
                        DataLlamadas.Columns[3].Width = 97;
                        DataLlamadas.Columns[4].HeaderText = "Cl.Llam";
                        DataLlamadas.Columns[4].Width = 67;
                        DataLlamadas.Columns[5].HeaderText = "DUR";
                        DataLlamadas.Columns[5].Width = 52;
                        DataLlamadas.Columns[6].HeaderText = "Vr.neto";
                        DataLlamadas.Columns[6].Width = 67;
                        DataLlamadas.Columns[7].HeaderText = "Vr.Recargo";
                        DataLlamadas.Columns[7].Width = 67;
                        DataLlamadas.Columns[8].HeaderText = "Vr.IVA";
                        DataLlamadas.Columns[8].Width = 67;
                        DataLlamadas.Columns[9].HeaderText = "Vr.Total";
                        DataLlamadas.Columns[9].Width = 73;


                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);
                            DataLlamadas.Rows.Add(n);
                        }
                        RowTotal = new string[10] { "TOTAL:", " ", " ", " ", " ", DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                        AnchoPDF = new float[10] { 8.77f, 7.48f, 20, 20, 8.77f, 5.60f, 9f, 9f, 9f, 9f };
                        pdfTable.SetWidths(AnchoPDF);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SetWidths(AnchoPDF);
                        foreach (DataGridViewColumn column in DataLlamadas.Columns)
                        {
                            cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                            pdfTable.AddCell(cell);
                        }
                        foreach (DataGridViewRow row in DataLlamadas.Rows)
                        {
                            foreach (DataGridViewCell celda in row.Cells)
                            {
                                cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                pdfTable.AddCell(cell);
                            }
                        }
                        pdfDoc.Add(pdfTable);
                        pdfDoc.Add(new Paragraph("\n\n"));
                    }
                }
                else
                {
                    IncrementoGen = 10;
                    Visor.RowCount = Visor.RowCount + 1;
                    DataLlamadas = new DataGridView();
                    Visor.Controls.Add(DataLlamadas);
                    DataLlamadas.ColumnCount = 8;
                    DataLlamadas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    DataLlamadas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    DataLlamadas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.AutoSize = true;
                    DataLlamadas.RowHeadersVisible = false;
                    DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                    DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    DataLlamadas.BorderStyle = BorderStyle.None;
                    DataLlamadas.AllowUserToAddRows = false;
                    DataLlamadas.AllowUserToDeleteRows = false;
                    DataLlamadas.AllowUserToResizeRows = false;
                    DataLlamadas.MultiSelect = false;
                    DataLlamadas.ReadOnly = true;
                    DataLlamadas.Enabled = false;


                    DataLlamadas.Columns[0].HeaderText = "EXTENSIÓN";
                    DataLlamadas.Columns[0].Width = 210;
                    DataLlamadas.Columns[1].HeaderText = "C.COSTO";
                    DataLlamadas.Columns[1].Width = 62;
                    DataLlamadas.Columns[2].HeaderText = "CANTIDAD";
                    DataLlamadas.Columns[2].Width = 65;
                    DataLlamadas.Columns[3].HeaderText = "DURACIÓN";
                    DataLlamadas.Columns[3].Width = 68;
                    DataLlamadas.Columns[4].HeaderText = "Vr.Neto";
                    DataLlamadas.Columns[4].Width = 72;
                    DataLlamadas.Columns[5].HeaderText = "Vr.Recargo";
                    DataLlamadas.Columns[5].Width = 72;
                    DataLlamadas.Columns[6].HeaderText = "Vr.IVA";
                    DataLlamadas.Columns[6].Width = 82;
                    DataLlamadas.Columns[7].HeaderText = "Vr.Total";
                    DataLlamadas.Columns[7].Width = 82;

                    CantRes = 0;
                    DurRes = 0;
                    VrNetoRes = 0;
                    VrRecargoRes = 0;
                    VrIVARes = 0;
                    VrTotalRes = 0;

                    foreach (List<string[]> s in EXT)
                    {
                        LabRes = "";
                        CentroCostoRes = "";
                        LabRes = "EXT: " + (s[0][10]) + "    ";
                        try
                        {
                            using (Conexion = new MySqlConnection(conexion))
                            {
                                Conexion.Open();
                                query = "select * from extensiones where Nume_Extension = ?e";
                                comando = new MySqlCommand(query, Conexion);
                                comando.Parameters.AddWithValue("?e", (s[0][10]));
                                lee = comando.ExecuteReader();
                                lee.Read();
                                LabRes += lee["Nomb_Extension"].ToString();
                                CentroCostoRes = lee["Codi_Centro"].ToString();
                                lee.Close();
                                Conexion.Close();
                            }
                        }
                        catch
                        {
                            LabRes += "Extension desconocida";
                            CentroCostoRes = "Centro de costo desconocido";
                        }

                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);

                        }
                        RowTotal = new string[8] { LabRes, CentroCostoRes, s.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        DataLlamadas.Size = new Size(485, IncrementoGen += 10);
                        DurRes += DurGen;
                        VrNetoRes += VrNetoGen;
                        VrRecargoRes += VrRecargoGen;
                        VrIVARes += VrIvaGen;
                        VrTotalRes += VrTotalGen;
                        CantRes += s.Count;
                    }
                    RowTotal = new string[8] { "TOTAL", " ", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                    DataLlamadas.Rows.Add(RowTotal);
                    pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    AnchoPDF = new float[8] { 29f, 20f, 8f, 8f, 10f, 10f, 10f, 10f };
                    pdfTable.SetWidths(AnchoPDF);
                    pdfTable.WidthPercentage = 100;
                    pdfTable.SetWidths(AnchoPDF);
                    foreach (DataGridViewColumn column in DataLlamadas.Columns)
                    {
                        cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                        pdfTable.AddCell(cell);
                    }
                    foreach (DataGridViewRow row in DataLlamadas.Rows)
                    {
                        AnchoPDFpos = 0;
                        foreach (DataGridViewCell celda in row.Cells)
                        {
                            cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                            if(AnchoPDFpos == 0)
                            {
                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                            }
                            else
                            {
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                            }
                            pdfTable.AddCell(cell);
                            AnchoPDFpos++;
                        }
                    }
                    pdfDoc.Add(pdfTable);
                    pdfDoc.Add(new Paragraph("\n\n"));
                }

                Visor.RowCount = Visor.RowCount + 1;
                lab = new Label();
                lab.Text = "TOTAL:";
                lab.AutoSize = true;
                lab.BorderStyle = BorderStyle.None;
                Visor.Controls.Add(lab);
                pdfDoc.Add(new Paragraph(lab.Text, Fuente));

                Visor.RowCount = Visor.RowCount + 1;
                DataLlamadas = new DataGridView();
                Visor.Controls.Add(DataLlamadas);
                DataLlamadas.ColumnCount = 4;
                DataLlamadas.AutoSize = true;
                DataLlamadas.RowHeadersVisible = false;
                DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                DataLlamadas.BorderStyle = BorderStyle.None;
                DataLlamadas.AllowUserToAddRows = false;
                DataLlamadas.AllowUserToDeleteRows = false;
                DataLlamadas.AllowUserToResizeRows = false;
                DataLlamadas.MultiSelect = false;
                DataLlamadas.ReadOnly = true;
                DataLlamadas.Enabled = false;

                DataLlamadas.Columns[0].HeaderText = "Cl.Llam";
                DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                DataLlamadas.Columns[3].HeaderText = "Vr.Total";

                foreach (string n in CheckCL)
                {
                    if (n.Equals("LOC"))
                    {
                        RowTotal = new string[4] { n, LOCCant.ToString(), LOCDur.ToString(), LOCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDN"))
                    {
                        RowTotal = new string[4] { n, DDNCant.ToString(), DDNDur.ToString(), DDNTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("CEL"))
                    {
                        RowTotal = new string[4] { n, CELCant.ToString(), CELDur.ToString(), CELTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("TOL"))
                    {
                        RowTotal = new string[4] { n, TOLCant.ToString(), TOLDur.ToString(), TOLTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDI"))
                    {
                        RowTotal = new string[4] { n, DDICant.ToString(), DDIDur.ToString(), DDITot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }

                    else if (n.Equals("ENT"))
                    {
                        RowTotal = new string[4] { n, ENTCant.ToString(), ENTDur.ToString(), ENTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("EXC"))
                    {
                        RowTotal = new string[4] { n, EXCCant.ToString(), EXCDur.ToString(), EXCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INT"))
                    {
                        RowTotal = new string[4] { n, INTCant.ToString(), INTDur.ToString(), INTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INV"))
                    {
                        RowTotal = new string[4] { n, INVCant.ToString(), INVDur.ToString(), INVTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("ITH"))
                    {
                        RowTotal = new string[4] { n, ITHCant.ToString(), ITHDur.ToString(), ITHTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("SAT"))
                    {
                        RowTotal = new string[4] { n, SATCant.ToString(), SATDur.ToString(), SATTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                }
                TotalValores = LOCTot + DDNTot + CELTot + TOLTot + DDITot + ENTTot + EXCTot + INTTot + INVTot + ITHTot + SATTot;
                TotalDuracion = LOCDur + DDNDur + CELDur + TOLDur + DDIDur + ENTDur + EXCDur + INTDur + INVDur + ITHDur + SATDur;
                TotalCantidad = LOCCant + DDNCant + CELCant + TOLCant + DDICant + ENTCant + EXCCant + INTCant + INVCant + ITHCant + SATCant;
                RowTotal = new string[4] { "TOTAL:", TotalCantidad.ToString(), TotalDuracion.ToString(), TotalValores.ToString() };
                DataLlamadas.Rows.Add(RowTotal);
                pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                pdfTable.DefaultCell.PaddingBottom = 3;
                pdfTable.DefaultCell.PaddingTop = 3;
                pdfTable.WidthPercentage = 30;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
                foreach (DataGridViewColumn column in DataLlamadas.Columns)
                {
                    cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                    pdfTable.AddCell(cell);
                }
                foreach (DataGridViewRow row in DataLlamadas.Rows)
                {
                    foreach (DataGridViewCell celda in row.Cells)
                    {
                        cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                        pdfTable.AddCell(cell);
                    }
                }
                pdfDoc.Add(pdfTable);
                pdfDoc.Add(new Paragraph("\n\n"));
                
                panel5.Controls.Add(Visor);
                
                pdfDoc.Close();
                stream.Close();
                MP.Hide();
                MP.label1.Text = "Enviando reporte programado, por favor espere";
                MuestraMensaje(np);
            }
        }

        public void TerminaReporte()
        {
            panel5.Controls.Clear();
            label188.Visible = false;
            progressBar1.Visible = false;
            button46.Visible = true;
            button46.Enabled = true;
        }

        public bool SePuede()
        {
            FechaRow = "";
            HoraRow = "";
            using (Conexion = new MySqlConnection(conexion))
            {
                Conexion.Open();
                query = "select * from llamadas_telefonicas where Errores = '-'";
                comando = new MySqlCommand(query, Conexion);
                lee = comando.ExecuteReader();
                lee.Read();
                if (!lee["FFechaFinalLlamada"].ToString().Equals("-"))
                {
                    FechaRow = "FFechaFinalLlamada";
                    if (!lee["HHoraFinalLlamada"].ToString().Equals("-"))
                    {
                        HoraRow = "HHoraFinalLlamada";
                        lee.Close();
                        Conexion.Close();
                        return (true);
                    }
                    else if (!lee["jHoraInicialLlamada"].ToString().Equals("-"))
                    {
                        HoraRow = "jHoraInicialLlamada";
                        lee.Close();
                        Conexion.Close();
                        return (true);
                    }
                    else
                    {
                        FechaRow = "";
                        HoraRow = "";
                        lee.Close();
                        Conexion.Close();
                        return (false);
                    }
                }
                else if (!lee["mFechaInicialLlamada"].ToString().Equals("-"))
                {
                    FechaRow = "mFechaInicialLlamada";
                    if (!lee["HHoraFinalLlamada"].ToString().Equals("-"))
                    {
                        HoraRow = "HHoraFinalLlamada";
                        lee.Close();
                        Conexion.Close();
                        return (true);
                    }
                    else if (!lee["jHoraInicialLlamada"].ToString().Equals("-"))
                    {
                        HoraRow = "jHoraInicialLlamada";
                        lee.Close();
                        Conexion.Close();
                        return (true);
                    }
                    else
                    {
                        FechaRow = "";
                        HoraRow = "";
                        lee.Close();
                        Conexion.Close();
                        return (false);
                    }
                }
                else
                {
                    FechaRow = "";
                    HoraRow = "";
                    lee.Close();
                    Conexion.Close();
                    return (false);
                }
            }
        }
        
        public void LeePos()
        {
            PosDia = "";
            Posc = 0;
            for (int i = 0; i < FormatoFechaFinal.Length; i++)
            {
                if (FormatoFechaFinal.ToLower()[i].Equals('d'))
                {
                    PosDia = i.ToString();
                    i = FormatoFechaFinal.Length;
                    foreach (char s in FormatoFechaFinal.ToLower())
                    {
                        if (s.Equals('d'))
                        {
                            Posc++;
                        }
                    }
                }
            }
            PosDia += "-" + Posc.ToString();

            PosHora = "";
            Posc = 0;
            for (int i = 0; i < FormatoHoraFinal.Length; i++)
            {
                if (FormatoHoraFinal.ToLower()[i].Equals('h'))
                {
                    PosHora = i.ToString();
                    i = FormatoHoraFinal.Length;
                    foreach (char s in FormatoHoraFinal.ToLower())
                    {
                        if (s.Equals('h'))
                        {
                            Posc++;
                        }
                    }
                }
            }
            PosHora += "-" + Posc.ToString();

            PosMinutos = "";
            Posc = 0;
            for (int i = 0; i < FormatoHoraFinal.Length; i++)
            {
                if (FormatoHoraFinal.ToLower()[i].Equals('m'))
                {
                    PosMinutos = i.ToString();
                    i = FormatoHoraFinal.Length;
                    foreach (char s in FormatoHoraFinal.ToLower())
                    {
                        if (s.Equals('m'))
                        {
                            Posc++;
                        }
                    }
                }
            }
            PosMinutos += "-" + Posc.ToString();
        }

        public void CargaChecked()
        {
            CheckCE = new List<string>();
            CheckCL = new List<string>();
            CheckCC = new List<string>();
            CheckT = new List<string>();

            foreach (CheckBox s in TableCEL)
            {
                if (s.Checked == true)
                {
                    CheckCE.Add(s.Text.Split(' ')[0]);
                }
            }
            foreach (CheckBox s in TableCLL)
            {
                if (s.Checked == true)
                {
                    CheckCL.Add(s.Text.Split(' ')[0]);
                }
            }
            foreach (CheckBox s in TableCCL)
            {
                if (s.Checked == true)
                {
                    CheckCC.Add(s.Text.Split(' ')[0]);
                }
            }
            foreach (CheckBox s in TableTL)
            {
                if (s.Checked == true)
                {
                    CheckT.Add(s.Text.Split(' ')[0]);
                }
            }
        }
        
        public void FiltraFecha(string Formato)
        {
            try
            {
                foreach (char s in Formato.Split('-')[0].ToLower())
                {
                    if (s.Equals('m'))
                    {
                        FormatoFechaFinal += 'M';
                    }
                    else
                    {
                        FormatoFechaFinal += s;
                    }
                }
                dateTimePicker1.Invoke(new Action(() => { dateTimePicker1.Format = DateTimePickerFormat.Custom; dateTimePicker1.CustomFormat = FormatoFechaFinal; }));
                dateTimePicker8.Invoke(new Action(() => { dateTimePicker8.Format = DateTimePickerFormat.Custom; dateTimePicker8.CustomFormat = FormatoFechaFinal; }));
                dateTimePicker2.Invoke(new Action(() => { dateTimePicker2.Format = DateTimePickerFormat.Custom; dateTimePicker2.CustomFormat = FormatoFechaFinal; }));
                dateTimePicker7.Invoke(new Action(() => { dateTimePicker7.Format = DateTimePickerFormat.Custom; dateTimePicker7.CustomFormat = FormatoFechaFinal; }));
                label190.Invoke(new Action(() => { label190.Text = "Formato: " + FormatoFechaFinal; }));
                label190.Invoke(new Action(() => { label194.Text = "Formato: " + FormatoFechaFinal; }));

                PosicionDiaI = Convert.ToInt32(Formato.Split('-')[1].Split('.')[0]);
                PosicionDiaF = Convert.ToInt32(Formato.Split('-')[1].Split('.')[1]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al procesar el formato de la fecha\n\n" + ex.ToString());
            }
        }

        public void FiltraHora(string Formato)
        {
            Formato = Formato.ToLower();
            foreach (char s in Formato)
            {
                if (s.Equals('h'))
                {
                    FormatoHoraFinal += 'H';
                }
                else if(s.Equals('m'))
                {
                    FormatoHoraFinal += 'm';
                }
            }
            dateTimePicker3.Invoke(new Action(() => { dateTimePicker3.Format = DateTimePickerFormat.Custom; dateTimePicker3.CustomFormat = FormatoHoraFinal; }));
            dateTimePicker6.Invoke(new Action(() => { dateTimePicker6.Format = DateTimePickerFormat.Custom; dateTimePicker6.CustomFormat = FormatoHoraFinal; }));
            dateTimePicker4.Invoke(new Action(() => { dateTimePicker4.Format = DateTimePickerFormat.Custom; dateTimePicker4.CustomFormat = FormatoHoraFinal; }));
            dateTimePicker5.Invoke(new Action(() => { dateTimePicker5.Format = DateTimePickerFormat.Custom; dateTimePicker5.CustomFormat = FormatoHoraFinal; }));
            label191.Invoke(new Action(() => { label191.Text = "Formato: " + FormatoHoraFinal; }));
            label191.Invoke(new Action(() => { label193.Text = "Formato: " + FormatoHoraFinal; }));
        }
        
        private void button52_Click(object sender, EventArgs e)
        {
            panel5.Controls.Clear();
        }


        #endregion

        #region Correcto

        public bool Correcto()
        {
            SelecFiltro = false;
            foreach (CheckBox s in TableCEL)
            {
                if (s.Checked == true)
                {
                    SelecFiltro = true;
                }
            }
            if (SelecFiltro == true)
            {
                SelecFiltro = false;
                foreach (CheckBox s in TableCLL)
                {
                    if (s.Checked == true)
                    {
                        SelecFiltro = true;
                    }
                }
                if (SelecFiltro == true)
                {
                    SelecFiltro = false;
                    foreach (CheckBox s in TableCCL)
                    {
                        if (s.Checked == true)
                        {
                            SelecFiltro = true;
                        }
                    }
                    if (SelecFiltro == true)
                    {
                        SelecFiltro = false;
                        foreach (CheckBox s in TableTL)
                        {
                            if (s.Checked == true)
                            {
                                SelecFiltro = true;
                            }
                        }
                        if (SelecFiltro == true)
                        {
                            return (true);
                        }
                        else
                        {
                            MessageBox.Show("No ha seleccionado un filtro en las troncales");
                            return (false);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No ha seleccionado un filtro en los centros de costo");
                        return (false);
                    }
                }
                else
                {
                    MessageBox.Show("No ha seleccionado un filtro en las calses de llamadas!");
                    return (false);
                }
            }
            else
            {
                MessageBox.Show("No ha seleccionado un filtro en las calses de extensión!");
                return (false);
            }
        }

        #endregion

        #region Guarda y carga

        private void button49_Click(object sender, EventArgs e)
        {
            try
            {
                if (Correcto() == true)
                {
                    saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
                    saveFileDialog1.FilterIndex = 1;
                    saveFileDialog1.RestoreDirectory = true;
                    CargaChecked();
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        using (StreamWriter escritor = new StreamWriter(saveFileDialog1.OpenFile()))
                        {
                            escritor.WriteLine("Reporte general");
                            escritor.WriteLine(dateTimePicker1.Value.ToString());
                            escritor.WriteLine(dateTimePicker2.Value.ToString());
                            escritor.WriteLine(dateTimePicker3.Value.ToString());
                            escritor.WriteLine(dateTimePicker4.Value.ToString());
                            escritor.WriteLine(checkBox1.CheckState);
                            escritor.WriteLine(checkBox2.CheckState);
                            escritor.WriteLine(checkBox3.CheckState);
                            escritor.WriteLine("Filtros CE");
                            foreach (string s in CheckCE)
                            {
                                escritor.WriteLine(s);
                            }
                            escritor.WriteLine("Filtros CL");
                            foreach (string s in CheckCL)
                            {
                                escritor.WriteLine(s);
                            }
                            escritor.WriteLine("Filtros CC");
                            foreach (string s in CheckCC)
                            {
                                escritor.WriteLine(s);
                            }
                            escritor.WriteLine("Filtros T");
                            foreach (string s in CheckT)
                            {
                                escritor.WriteLine(s);
                            }
                            escritor.Close();
                        }
                        MessageBox.Show("La configuración se ha guardado con éxito");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al guardar la configuración :\n\n" + ex.ToString());
            }
        }
        string Linea = "";
        List<string> Lineas;
        List<string> FiltrosCE;
        List<string> FiltrosCC;
        List<string> FiltrosCL;
        List<string> FiltrosT;
        bool CE = false;
        bool CC = false;
        bool CL = false;
        bool T = false;
        private void button50_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            Lineas = new List<string>();
            FiltrosCE = new List<string>();
            FiltrosCL = new List<string>();
            FiltrosCC = new List<string>();
            FiltrosT = new List<string>();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (StreamReader lector = new StreamReader(openFileDialog1.OpenFile()))
                    {
                        while((Linea = lector.ReadLine()) != null)
                        {
                            Lineas.Add(Linea);
                        }
                        lector.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al leer el archivo\n\n" + ex.ToString());
                }
                if (Lineas.Count != 0)
                {
                    if (Lineas[0].Equals("Reporte general"))
                    {
                        foreach (string s in Lineas)
                        {
                            if (s.Equals("Filtros CE"))
                            {
                                CE = true;
                                CL = false;
                                CC = false;
                                T = false;
                            }
                            else if (s.Equals("Filtros CL"))
                            {
                                CE = false;
                                CL = true;
                                CC = false;
                                T = false;
                            }
                            else if (s.Equals("Filtros CC"))
                            {
                                CE = false;
                                CL = false;
                                CC = true;
                                T = false;
                            }
                            else if (s.Equals("Filtros T"))
                            {
                                CE = false;
                                CL = false;
                                CC = false;
                                T = true;
                            }


                            if (CE == true)
                            {
                                if (!s.Equals("Filtros CE"))
                                {
                                    FiltrosCE.Add(s);
                                }
                            }
                            else if (CL == true)
                            {
                                if (!s.Equals("Filtros CL"))
                                {
                                    FiltrosCL.Add(s);
                                }
                            }
                            else if (CC == true)
                            {
                                if (!s.Equals("Filtros CC"))
                                {
                                    FiltrosCC.Add(s);
                                }
                            }
                            else if (T == true)
                            {
                                if (!s.Equals("Filtros T"))
                                {
                                    FiltrosT.Add(s);
                                }
                            }

                        }

                        dateTimePicker1.Text = Lineas[1];
                        dateTimePicker2.Text = Lineas[2];
                        dateTimePicker3.Text = Lineas[3];
                        dateTimePicker4.Text = Lineas[4];
                        if (Lineas[5].Equals("Checked")) { checkBox1.Checked = true; } else if (Lineas[5].Equals("Unchecked")) { checkBox1.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de reportes resumen"); checkBox1.Checked = false; }
                        if (Lineas[6].Equals("Checked")) { checkBox2.Checked = true; } else if (Lineas[6].Equals("Unchecked")) { checkBox2.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de llamadas extensas"); checkBox2.Checked = false; }
                        if (Lineas[7].Equals("Checked")) { checkBox3.Checked = true; } else if (Lineas[7].Equals("Unchecked")) { checkBox3.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de llamadas con valor"); checkBox3.Checked = false; }
                        if (FiltrosCE.Count != 0)
                        {
                            foreach (string s in FiltrosCE)
                            {
                                foreach (CheckBox n in TableCEL)
                                {
                                    if (n.Text.Split(' ')[0].Equals(s))
                                    {
                                        n.Checked = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Error al cargar las clases de extensión");
                        }
                        if (FiltrosCL.Count != 0)
                        {
                            foreach (string s in FiltrosCL)
                            {
                                foreach (CheckBox n in TableCLL)
                                {
                                    if (n.Text.Split(' ')[0].Equals(s))
                                    {
                                        n.Checked = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Error al cargar las clases de llamadas");
                        }
                        if (FiltrosCC.Count != 0)
                        {
                            foreach (string s in FiltrosCC)
                            {
                                foreach (CheckBox n in TableCCL)
                                {
                                    if (n.Text.Split(' ')[0].Equals(s))
                                    {
                                        n.Checked = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Error al cargar las clases de extensión");
                        }
                        if (FiltrosT.Count != 0)
                        {
                            foreach (string s in FiltrosT)
                            {
                                foreach (CheckBox n in TableTL)
                                {
                                    if (n.Text.Split(' ')[0].Equals(s))
                                    {
                                        n.Checked = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Error al cargar las clases de extensión");
                        }
                        MessageBox.Show("La configuración se ha cargado");
                    }
                    else
                    {
                        MessageBox.Show("El archivo de configuración seleccionado no corresponde a una configuración de un reporte general");
                    }
                }
                else
                {
                    MessageBox.Show("No se ha detectado una configuración en el archivo seleccionado");
                }
            }
        }

        #endregion

        #endregion

        #region Reportes Específicos

        #region Procesa
        
        List<string> RangoRad;
        private void button56_Click(object sender, EventArgs e)
        {
            if (Correcto2() == true)
            {
                if (SePuede() == true)
                {
                    FechaInicial = dateTimePicker8.Text.ToString();
                    HoraInicial = dateTimePicker6.Text.ToString();
                    FechaFinal = dateTimePicker7.Text.ToString();
                    HoraFinal = dateTimePicker5.Text.ToString();
                    DialogResult dialogResult2 = MessageBox.Show("Desde: " + FechaInicial + " a las: " + HoraInicial + " Hasta: " + FechaFinal + " a las: " + HoraFinal + ". ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
                    if (dialogResult2 == DialogResult.Yes)
                    {
                        FIltroPrincipalCorrecto = true;
                        button56.Enabled = false;
                        button56.Visible = false;
                        label196.Visible = true;
                        progressBar2.Value = 0;
                        progressBar2.Maximum = 100;
                        progressBar2.Visible = true;
                        CargaChecked2();
                        LeePos();
                        CargaRad();
                        CargaCheckBox2();

                        using (Conexion = new MySqlConnection(conexion))
                        {
                            Conexion.Open();
                            query = "SHOW TABLES";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            TablasNumeros = new List<string>();
                            while (lee.Read())
                            {
                                NombreTabla = "";
                                EsNumerico = false;
                                NombreTabla = lee.GetValue(0).ToString();
                                EsNumerico = int.TryParse(NombreTabla.Split(' ')[0], out Out1);
                                if (EsNumerico == true)
                                {
                                    TablasNumeros.Add(NombreTabla);
                                }
                            }
                            lee.Close();
                            Conexion.Close();
                        }
                        if (TablasNumeros.Count >= 1)
                        {
                            Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("yyyy"));
                            Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("yyyy"));
                            Out3 = Convert.ToInt32(dateTimePicker8.Value.ToString("MM"));
                            Out4 = Convert.ToInt32(dateTimePicker7.Value.ToString("MM"));
                            Filtrados = new List<string>();
                            foreach (string S in TablasNumeros)
                            {
                                EnRango = true;

                                if (Convert.ToInt32(S.Split(' ')[0]) >= Out1 && Convert.ToInt32(S.Split(' ')[0]) <= Out2)
                                {
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out1)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) < Out3)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out2)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) > Out4)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                }
                                else
                                {
                                    EnRango = false;
                                }

                                if (EnRango == true)
                                {
                                    Filtrados.Add(S);
                                }
                            }
                            if (Filtrados.Count >= 1)
                            {
                                TablasNumeros = Filtrados;
                                if (TablasNumeros.Count == 1)
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("dd"));
                                        Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("dd"));
                                        Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                        Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                        HoraI = dateTimePicker6.Value.ToString("HH:mm");
                                        HoraF = dateTimePicker5.Value.ToString("HH:mm");
                                        LlamadasFiltradas = new List<string[]>();
                                        TotalRegistros = 0;
                                        while (lee.Read())
                                        {
                                            TotalRegistros++;
                                        }
                                        lee.Close();
                                        progressBar2.Value = 0;
                                        RegTot = TotalRegistros;
                                        TotalRegistros = 0;
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        while (lee.Read())
                                        {
                                            Application.DoEvents();
                                            Minutos = 0;
                                            if (lee["Errores"].ToString().Equals("-"))
                                            {
                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1 && Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                {
                                                    if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1 && Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                    {
                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])) && Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[16];
                                                            FiltraCheck2();
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                    {
                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[16];
                                                            FiltraCheck2();
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                    {

                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                        {
                                                            llamadasFil = new string[16];
                                                            FiltraCheck2();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2();
                                                    }


                                                }
                                            }
                                            TotalRegistros++;
                                            progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);

                                        }
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                else
                                {
                                    progressBar2.Visible = true;
                                    progressBar2.Value = 0;
                                    TotalRegistros = 0;
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        using (Conexion = new MySqlConnection(conexion))
                                        {
                                            Conexion.Open();
                                            query = "select * from `" + TablasNumeros[i] + "`";
                                            comando = new MySqlCommand(query, Conexion);
                                            lee = comando.ExecuteReader();
                                            while (lee.Read())
                                            {
                                                TotalRegistros++;
                                            }
                                            lee.Close();
                                            Conexion.Close();
                                        }
                                    }
                                    progressBar2.Maximum = 100;
                                    RegTot = TotalRegistros;
                                    TotalRegistros = 0;
                                    LlamadasFiltradas = new List<string[]>();
                                    Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("dd"));
                                    Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("dd"));
                                    Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                    Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                    HoraI = dateTimePicker6.Value.ToString("HH:mm");
                                    HoraF = dateTimePicker5.Value.ToString("HH:mm");
                                    MesI = Convert.ToInt32(dateTimePicker8.Value.ToString("MM"));
                                    MesF = Convert.ToInt32(dateTimePicker7.Value.ToString("MM"));
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        if (i == 0)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker8.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                                    {
                                                                        llamadasFil = new string[16];
                                                                        FiltraCheck2();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    llamadasFil = new string[16];
                                                                    FiltraCheck2();
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            llamadasFil = new string[16];
                                                            FiltraCheck2();
                                                        }
                                                    }

                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else if (i == TablasNumeros.Count - 1)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker7.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                                    {
                                                                        llamadasFil = new string[16];
                                                                        FiltraCheck2();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    llamadasFil = new string[16];
                                                                    FiltraCheck2();
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            llamadasFil = new string[16];
                                                            FiltraCheck2();
                                                        }
                                                    }

                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from " + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2();
                                                    }
                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);

                                                }
                                                Conexion.Close();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("No se han detectado llamadas dentro del rango especificado");
                                FIltroPrincipalCorrecto = false;
                                TerminaReporte2();
                            }

                        }
                        else
                        {
                            MessageBox.Show("No se han detectado llamadas en la base de datos");
                            FIltroPrincipalCorrecto = false;
                            TerminaReporte2();
                        }
                        if(FIltroPrincipalCorrecto == true)
                        {
                            TerminaReporte2();
                            MuestraReporte2(Directory.GetCurrentDirectory() + @"\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".pdf");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No se ha detectado una fecha válida en la cadena");
                    TerminaReporte2();
                }
            }
        }

        bool EncontradoRad = false;
        public void FiltraCheck2()
        {
            try
            {
                for (int s = 0; s < CheckCL2.Count; s++)
                {
                    if (CheckCL2[s].Equals(lee["ClaseLlamada"].ToString()))
                    {
                        FiltroParametro = true;
                        if (checkBox5.Checked == true)
                        {
                            using (Conexion = new MySqlConnection(conexion))
                            {
                                Conexion.Open();
                                query = "select * from parametros where parametro = 'Llamadas extensas especificos'";
                                comando2 = new MySqlCommand(query, Conexion);
                                lee2 = comando2.ExecuteReader();
                                lee2.Read();
                                if ((Convert.ToInt32(lee["DuracionLlamadaAproximada"].ToString())) < (Convert.ToInt32(lee2["seleccion"].ToString())))
                                {
                                    FiltroParametro = false;
                                }
                                lee2.Close();
                                Conexion.Close();
                            }
                        }
                        if (checkBox4.Checked == true)
                        {
                            using (Conexion = new MySqlConnection(conexion))
                            {
                                Conexion.Open();
                                query = "select * from parametros where parametro = 'Llamadas con valor especificos'";
                                comando2 = new MySqlCommand(query, Conexion);
                                lee2 = comando2.ExecuteReader();
                                lee2.Read();
                                if ((Convert.ToInt32(lee["ValorTotal"].ToString())) < (Convert.ToInt32(lee2["seleccion"].ToString())))
                                {
                                    FiltroParametro = false;
                                }
                                lee2.Close();
                                Conexion.Close();
                            }

                        }
                        if(FiltroParametro == true)
                        {
                            EncontradoRad = false;
                            if (radioButton1.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    if (n.Equals(lee["EExtension"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (radioButton2.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    if (n.Equals(lee["CentroDeCosto"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (radioButton3.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    if (n.Equals(lee["TTroncal"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (radioButton5.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    if (n.Equals(lee["NNumeroMarcado"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (radioButton6.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    if (n.Equals(lee["NumeFolio"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (radioButton4.Checked)
                            {
                                foreach (string n in RangoRad)
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from codigos_personales where Nomb_Cod_Personal = '" + n + "'";
                                        comando2 = new MySqlCommand(query, Conexion);
                                        lee2 = comando2.ExecuteReader();
                                        while (lee2.Read())
                                        {
                                            if (lee2["Codi_Personal"].ToString().Equals(lee["PCodigoPersonal"].ToString()))
                                            {
                                                s = CheckCL2.Count;
                                                llamadasFil[0] = lee["FFechaFinalLlamada"].ToString();
                                                llamadasFil[1] = lee["HHoraFinalLlamada"].ToString();
                                                llamadasFil[2] = lee["NNumeroMarcado"].ToString();
                                                llamadasFil[3] = lee["Destino"].ToString();
                                                llamadasFil[4] = lee["ClaseLlamada"].ToString();
                                                llamadasFil[5] = lee["DuracionLlamadaAproximada"].ToString();
                                                llamadasFil[6] = lee["ValorLlamadaTarifa"].ToString();
                                                llamadasFil[7] = lee["RecargoServicioValor"].ToString();
                                                llamadasFil[8] = lee["ValorIVA"].ToString();
                                                llamadasFil[9] = lee["ValorTotal"].ToString();
                                                llamadasFil[10] = lee["EExtension"].ToString();
                                                llamadasFil[11] = "--";
                                                llamadasFil[12] = lee["CentroDeCosto"].ToString();
                                                llamadasFil[13] = lee["TTroncal"].ToString();
                                                llamadasFil[14] = n;
                                                llamadasFil[15] = lee["NumeFolio"].ToString();
                                                LlamadasFiltradas.Add(llamadasFil);
                                            }
                                        }
                                        lee2.Close();
                                        Conexion.Close();
                                    }
                                }
                                EncontradoRad = false;
                            }
                            if (EncontradoRad == true)
                            {
                                s = CheckCL2.Count;
                                llamadasFil[0] = lee["FFechaFinalLlamada"].ToString();
                                llamadasFil[1] = lee["HHoraFinalLlamada"].ToString();
                                llamadasFil[2] = lee["NNumeroMarcado"].ToString();
                                llamadasFil[3] = lee["Destino"].ToString();
                                llamadasFil[4] = lee["ClaseLlamada"].ToString();
                                llamadasFil[5] = lee["DuracionLlamadaAproximada"].ToString();
                                llamadasFil[6] = lee["ValorLlamadaTarifa"].ToString();
                                llamadasFil[7] = lee["RecargoServicioValor"].ToString();
                                llamadasFil[8] = lee["ValorIVA"].ToString();
                                llamadasFil[9] = lee["ValorTotal"].ToString();
                                llamadasFil[10] = lee["EExtension"].ToString();
                                llamadasFil[11] = "--";
                                llamadasFil[12] = lee["CentroDeCosto"].ToString();
                                llamadasFil[13] = lee["TTroncal"].ToString();
                                llamadasFil[14] = lee["PCodigoPersonal"].ToString();
                                llamadasFil[15] = lee["NumeFolio"].ToString();
                                LlamadasFiltradas.Add(llamadasFil);

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al aplicar los filtros!\n\n" + ex.ToString());
                FIltroPrincipalCorrecto = false;
            }
        }
        
        public void MuestraReporte2(string np)
        {
            using (FileStream stream = new FileStream(np, FileMode.Create))
            {
                MP.label1.Text = "Cargando reporte al visor, por favor espere...";
                MP.Show();
                Application.DoEvents();

                pdfDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                Fuente = new iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7, iTextSharp.text.Font.NORMAL);

                Filtrados = new List<string>();
                Head = new Label();
                try
                {
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from parametros where parametro = 'Reportes Hotel'";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        lee.Read();
                        Head.Text = lee["seleccion"].ToString();
                        lee.Close();
                        Conexion.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problema al conectarse con la base de datos\n\n" + ex.ToString());
                }
                Head.Text += "\n\nREPORTE ESPECÍFICO DE LLAMADAS" + "\nDesde: " + FechaInicial + " a las: " + HoraInicial + " Hasta: " + FechaFinal + " a las: " + HoraFinal + "\n\n Clase de llamadas: ";
                if (CheckCL2.Count == TableCLL2.Count)
                {
                    Head.Text += "Todos";
                }
                else
                {
                    foreach (string s in CheckCL2)
                    {
                        Head.Text += s + ", ";
                    }
                }
                if (radioButton1.Checked) { Head.Text += "\n\nREPORTE: " + radioButton1.Text; }
                else if (radioButton2.Checked) { Head.Text += "\n\nREPORTE: " + radioButton2.Text; }
                else if (radioButton3.Checked) { Head.Text += "\n\nREPORTE: " + radioButton3.Text; }
                else if (radioButton4.Checked) { Head.Text += "\n\nREPORTE: " + radioButton4.Text; }
                else if (radioButton5.Checked) { Head.Text += "\n\nREPORTE: " + radioButton5.Text; }
                else if (radioButton6.Checked) { Head.Text += "\n\nREPORTE: " + radioButton6.Text; }
                Head.Text += "\nDesde: " + comboBox22.Text.Split(' ')[0] + " Hasta: " + comboBox23.Text.Split(' ')[0] + "\n";
                if (checkBox6.Checked == true)
                {
                    Head.Text += "\nREPORTE RESUMIDO\n\n";
                }
                else
                {
                    Head.Text += "\nREPORTE DETALLADO\n\n";
                }

                Head.AutoSize = true;
                Head.BorderStyle = BorderStyle.None;

                Visor = new TableLayoutPanel();
                Visor.Controls.Add(Head);
                pdfDoc.Add(new Paragraph(Head.Text, Fuente));
                Visor.ColumnCount = 1;
                Visor.Location = new Point(0, 0);
                Visor.RowCount = 1;
                Visor.TabIndex = 0;
                Visor.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
                Visor.AutoSize = true;

                Visor.RowCount = Visor.RowCount + 1;
                EXT = new List<List<string[]>>();
                ext = new List<string[]>();

                if (radioButton1.Checked)
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][10]) > Convert.ToInt32(LlamadasFiltradas[b][10]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                else if (radioButton2.Checked)
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][12]) > Convert.ToInt32(LlamadasFiltradas[b][12]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                else if (radioButton3.Checked)
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            try
                            {
                                if (Convert.ToInt32(LlamadasFiltradas[b - 1][13]) > Convert.ToInt32(LlamadasFiltradas[b][13]))
                                {
                                    t = LlamadasFiltradas[b - 1];
                                    LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                    LlamadasFiltradas[b] = t;
                                }
                            }
                            catch
                            {

                            }
                            
                        }
                    }
                }
                else if (radioButton6.Checked)
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][15]) > Convert.ToInt32(LlamadasFiltradas[b][15]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }

                TotalRegistros = 0;
                progressBar2.Value = 0;
                progressBar2.Maximum = 100;
                foreach (string[] s in LlamadasFiltradas)
                {
                    Application.DoEvents();
                    Repetido = false;
                    if (EXT.Count > 0)
                    {
                        foreach (List<string[]> l in EXT)
                        {
                            foreach (string[] c in l)
                            {
                                if (radioButton1.Checked) { if (c[10].Equals(s[10])) { Repetido = true; } }
                                else if (radioButton2.Checked) { if (c[12].Equals(s[12])) { Repetido = true; } }
                                else if (radioButton3.Checked) { if (c[13].Equals(s[13])) { Repetido = true; } }
                                else if (radioButton4.Checked) { if (c[14].Equals(s[14])) { Repetido = true; } }
                                else if (radioButton5.Checked) { if (c[2].Equals(s[2])) { Repetido = true; } }
                                else if (radioButton6.Checked) { if (c[15].Equals(s[15])) { Repetido = true; } }
                            }
                        }
                        if (Repetido == false)
                        {
                            foreach (string[] n in LlamadasFiltradas)
                            {
                                if (radioButton1.Checked) { if (n[10].Equals(s[10])) { ext.Add(n); } }
                                else if (radioButton2.Checked) { if (n[12].Equals(s[12])) { ext.Add(n); } }
                                else if (radioButton3.Checked) { if (n[13].Equals(s[13])) { ext.Add(n); } }
                                else if (radioButton4.Checked) { if (n[14].Equals(s[14])) { ext.Add(n); } }
                                else if (radioButton5.Checked) { if (n[2].Equals(s[2])) { ext.Add(n); } }
                                else if (radioButton6.Checked) { if (n[15].Equals(s[15])) { ext.Add(n); } }
                            }
                            if (ext.Count > 0)
                            {
                                EXT.Add(ext);
                                ext = new List<string[]>();
                            }
                        }
                    }
                    else
                    {
                        foreach (string[] n in LlamadasFiltradas)
                        {
                            if (radioButton1.Checked) { if (n[10].Equals(s[10])) { ext.Add(n); } }
                            else if (radioButton2.Checked) { if (n[12].Equals(s[12])) { ext.Add(n); } }
                            else if (radioButton3.Checked) { if (n[13].Equals(s[13])) { ext.Add(n); } }
                            else if (radioButton4.Checked) { if (n[14].Equals(s[14])) { ext.Add(n); } }
                            else if (radioButton5.Checked) { if (n[2].Equals(s[2])) { ext.Add(n); } }
                            else if (radioButton6.Checked) { if (n[15].Equals(s[15])) { ext.Add(n); } }
                        }
                        if (ext.Count > 0)
                        {
                            EXT.Add(ext);
                            ext = new List<string[]>();
                        }
                    }
                    TotalRegistros++;
                    progressBar2.Value = (int)((TotalRegistros * 100) / LlamadasFiltradas.Count);
                }

                LOCDur = 0; LOCTot = 0; LOCCant = 0;
                DDNDur = 0; DDNTot = 0; DDNCant = 0;
                CELDur = 0; CELTot = 0; CELCant = 0;
                TOLDur = 0; TOLTot = 0; TOLCant = 0;
                DDIDur = 0; DDITot = 0; DDICant = 0;
                ENTDur = 0; ENTTot = 0; ENTCant = 0;
                EXCDur = 0; EXCTot = 0; EXCCant = 0;
                INTDur = 0; INTTot = 0; INTCant = 0;
                INVDur = 0; INVTot = 0; INVCant = 0;
                ITHDur = 0; ITHTot = 0; ITHCant = 0;
                SATDur = 0; SATTot = 0; SATCant = 0;
                TotalValores = 0;
                TotalDuracion = 0;
                TotalCantidad = 0;
                
                if (checkBox6.Checked == false)
                {
                    if(radioButton1.Checked || radioButton3.Checked || radioButton4.Checked || radioButton6.Checked || radioButton5.Checked)
                    {
                        TotalRegistros = 0;
                        progressBar2.Value = 0;
                        progressBar2.Maximum = 100;
                        foreach (List<string[]> s in EXT)
                        {
                            Application.DoEvents();
                            Visor.RowCount = Visor.RowCount + 1;
                            lab = new Label();
                            if (radioButton1.Checked)
                            {
                                lab.Text = "EXT: " + (s[0][10]) + "    ";
                                try
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from extensiones where Nume_Extension = ?e";
                                        comando = new MySqlCommand(query, Conexion);
                                        comando.Parameters.AddWithValue("?e", (s[0][10]));
                                        lee = comando.ExecuteReader();
                                        lee.Read();
                                        lab.Text += lee["Nomb_Extension"].ToString();
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                catch
                                {
                                    lab.Text += "Extension desconocida";
                                }
                            }
                            else if (radioButton3.Checked) { lab.Text = "TRONCAL: " + (s[0][13]); }
                            else if (radioButton4.Checked) { lab.Text = "CÓDIGO: " + (s[0][14]); }
                            else if (radioButton5.Checked) { lab.Text = "NÚMERO: " + (s[0][2]); }
                            else if (radioButton6.Checked) { lab.Text = "FOLIO: " + (s[0][15]); }

                            lab.AutoSize = true;
                            lab.BorderStyle = BorderStyle.None;
                            Visor.Controls.Add(lab);
                            pdfDoc.Add(new Paragraph(lab.Text, Fuente));
                            Visor.RowCount = Visor.RowCount + 1;
                            DataLlamadas = new DataGridView();
                            Visor.Controls.Add(DataLlamadas);
                            DataLlamadas.ColumnCount = 10;
                            DataLlamadas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            DataLlamadas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            DataLlamadas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            DataLlamadas.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            DataLlamadas.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            DataLlamadas.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataLlamadas.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataLlamadas.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataLlamadas.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataLlamadas.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataLlamadas.AutoSize = true;
                            DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                            DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                            DataLlamadas.BorderStyle = BorderStyle.None;
                            DataLlamadas.AllowUserToAddRows = false;
                            DataLlamadas.AllowUserToDeleteRows = false;
                            DataLlamadas.AllowUserToResizeRows = false;
                            DataLlamadas.RowHeadersVisible = false;
                            DataLlamadas.MultiSelect = false;
                            DataLlamadas.ReadOnly = true;
                            DataLlamadas.Enabled = false;

                            DataLlamadas.Columns[0].HeaderText = "FECHA";
                            DataLlamadas.Columns[0].Width = 57;
                            DataLlamadas.Columns[1].HeaderText = "HORA";
                            DataLlamadas.Columns[1].Width = 57;
                            DataLlamadas.Columns[2].HeaderText = "NUM.MARCADO";
                            DataLlamadas.Columns[2].Width = 110;
                            DataLlamadas.Columns[3].HeaderText = "DESTINO";
                            DataLlamadas.Columns[3].Width = 97;
                            DataLlamadas.Columns[4].HeaderText = "Cl.Llam";
                            DataLlamadas.Columns[4].Width = 67;
                            DataLlamadas.Columns[5].HeaderText = "DUR";
                            DataLlamadas.Columns[5].Width = 52;
                            DataLlamadas.Columns[6].HeaderText = "Vr.neto";
                            DataLlamadas.Columns[6].Width = 67;
                            DataLlamadas.Columns[7].HeaderText = "Vr.Recargo";
                            DataLlamadas.Columns[7].Width = 67;
                            DataLlamadas.Columns[8].HeaderText = "Vr.IVA";
                            DataLlamadas.Columns[8].Width = 67;
                            DataLlamadas.Columns[9].HeaderText = "Vr.Total";
                            DataLlamadas.Columns[9].Width = 73;


                            DurGen = 0;
                            VrNetoGen = 0;
                            VrRecargoGen = 0;
                            VrIvaGen = 0;
                            VrTotalGen = 0;
                            
                            foreach (string[] n in s)
                            {
                                Application.DoEvents();
                                if (n[4].Equals("LOC"))
                                {
                                    LOCDur += Convert.ToInt32(n[5]);
                                    LOCTot += Convert.ToInt32(n[9]);
                                    LOCCant++;
                                }
                                else if (n[4].Equals("DDN"))
                                {
                                    DDNDur += Convert.ToInt32(n[5]);
                                    DDNTot += Convert.ToInt32(n[9]);
                                    DDNCant++;
                                }
                                else if (n[4].Equals("CEL"))
                                {
                                    CELDur += Convert.ToInt32(n[5]);
                                    CELTot += Convert.ToInt32(n[9]);
                                    CELCant++;
                                }
                                else if (n[4].Equals("TOL"))
                                {
                                    TOLDur += Convert.ToInt32(n[5]);
                                    TOLTot += Convert.ToInt32(n[9]);
                                    TOLCant++;
                                }
                                else if (n[4].Equals("DDI"))
                                {
                                    DDIDur += Convert.ToInt32(n[5]);
                                    DDITot += Convert.ToInt32(n[9]);
                                    DDICant++;
                                }
                                else if (n[4].Equals("ENT"))
                                {
                                    ENTDur += Convert.ToInt32(n[5]);
                                    ENTTot += Convert.ToInt32(n[9]);
                                    ENTCant++;
                                }
                                else if (n[4].Equals("EXC"))
                                {
                                    EXCDur += Convert.ToInt32(n[5]);
                                    EXCTot += Convert.ToInt32(n[9]);
                                    EXCCant++;
                                }
                                else if (n[4].Equals("INT"))
                                {
                                    INTDur += Convert.ToInt32(n[5]);
                                    INTTot += Convert.ToInt32(n[9]);
                                    INTCant++;
                                }
                                else if (n[4].Equals("INV"))
                                {
                                    INVDur += Convert.ToInt32(n[5]);
                                    INVTot += Convert.ToInt32(n[9]);
                                    INVCant++;
                                }
                                else if (n[4].Equals("ITH"))
                                {
                                    ITHDur += Convert.ToInt32(n[5]);
                                    ITHTot += Convert.ToInt32(n[9]);
                                    ITHCant++;
                                }
                                else if (n[4].Equals("SAT"))
                                {
                                    SATDur += Convert.ToInt32(n[5]);
                                    SATTot += Convert.ToInt32(n[9]);
                                    SATCant++;
                                }

                                DurGen += Convert.ToInt32(n[5]);
                                VrNetoGen += Convert.ToInt32(n[6]);
                                VrRecargoGen += Convert.ToInt32(n[7]);
                                VrIvaGen += Convert.ToInt32(n[8]);
                                VrTotalGen += Convert.ToInt32(n[9]);
                                DataLlamadas.Rows.Add(n);
                            }
                            RowTotal = new string[10] { "TOTAL:", " ", " ", " ", " ", DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                            DataLlamadas.Rows.Add(RowTotal);

                            pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                            AnchoPDF = new float[10] { 8.77f, 7.48f, 20, 20, 8.77f, 5.60f, 9f, 9f, 9f, 9f };
                            pdfTable.SetWidths(AnchoPDF);
                            pdfTable.WidthPercentage = 100;
                            pdfTable.SetWidths(AnchoPDF);
                            foreach (DataGridViewColumn column in DataLlamadas.Columns)
                            {
                                cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                                pdfTable.AddCell(cell);
                            }
                            foreach (DataGridViewRow row in DataLlamadas.Rows)
                            {
                                foreach (DataGridViewCell celda in row.Cells)
                                {
                                    cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                    pdfTable.AddCell(cell);
                                }
                            }
                            pdfDoc.Add(pdfTable);
                            pdfDoc.Add(new Paragraph("\n\n"));
                            TotalRegistros++;
                            progressBar2.Value = (int)((TotalRegistros * 100) / EXT.Count);
                        }
                    }
                    else
                    {
                        if (radioButton2.Checked)
                        {
                            TotalRegistros = 0;
                            progressBar2.Value = 0;
                            progressBar2.Maximum = 100;
                            foreach (List<string[]> s in EXT)
                            {
                                Application.DoEvents();
                                Visor.RowCount = Visor.RowCount + 1;
                                lab = new Label();
                                lab.Text = "CENTRO: " + (s[0][12]) + "    ";
                                try
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from centros_costo where Codi_Centro = ?e";
                                        comando = new MySqlCommand(query, Conexion);
                                        comando.Parameters.AddWithValue("?e", (s[0][12]));
                                        lee = comando.ExecuteReader();
                                        lee.Read();
                                        lab.Text += lee["Nomb_Centro"].ToString();
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                catch
                                {
                                    lab.Text += "Centro de costo desconocida";
                                }
                                lab.AutoSize = true;
                                lab.BorderStyle = BorderStyle.None;
                                Visor.Controls.Add(lab);
                                pdfDoc.Add(new Paragraph(lab.Text, Fuente));
                                Visor.RowCount = Visor.RowCount + 1;
                                DataLlamadas = new DataGridView();
                                Visor.Controls.Add(DataLlamadas);
                                DataLlamadas.ColumnCount = 7;
                                DataLlamadas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                                DataLlamadas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                                DataLlamadas.AutoSize = true;
                                DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                                DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                DataLlamadas.BorderStyle = BorderStyle.None;
                                DataLlamadas.AllowUserToAddRows = false;
                                DataLlamadas.AllowUserToDeleteRows = false;
                                DataLlamadas.AllowUserToResizeRows = false;
                                DataLlamadas.RowHeadersVisible = false;
                                DataLlamadas.MultiSelect = false;
                                DataLlamadas.ReadOnly = true;
                                DataLlamadas.Enabled = false;

                                DataLlamadas.Columns[0].HeaderText = "EXTENSIÓN";
                                DataLlamadas.Columns[0].Width = 242;
                                DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                                DataLlamadas.Columns[1].Width = 70;
                                DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                                DataLlamadas.Columns[2].Width = 73;
                                DataLlamadas.Columns[3].HeaderText = "Vr.Neto";
                                DataLlamadas.Columns[3].Width = 77;
                                DataLlamadas.Columns[4].HeaderText = "Vr.Recargo";
                                DataLlamadas.Columns[4].Width = 77;
                                DataLlamadas.Columns[5].HeaderText = "Vr.IVA";
                                DataLlamadas.Columns[5].Width = 87;
                                DataLlamadas.Columns[6].HeaderText = "Vr.Total";
                                DataLlamadas.Columns[6].Width = 87;
                                
                                EXT2 = new List<List<string[]>>();
                                foreach (string[] n in s)
                                {
                                    Application.DoEvents();
                                    if (n[4].Equals("LOC"))
                                    {
                                        LOCDur += Convert.ToInt32(n[5]);
                                        LOCTot += Convert.ToInt32(n[9]);
                                        LOCCant++;
                                    }
                                    else if (n[4].Equals("DDN"))
                                    {
                                        DDNDur += Convert.ToInt32(n[5]);
                                        DDNTot += Convert.ToInt32(n[9]);
                                        DDNCant++;
                                    }
                                    else if (n[4].Equals("CEL"))
                                    {
                                        CELDur += Convert.ToInt32(n[5]);
                                        CELTot += Convert.ToInt32(n[9]);
                                        CELCant++;
                                    }
                                    else if (n[4].Equals("TOL"))
                                    {
                                        TOLDur += Convert.ToInt32(n[5]);
                                        TOLTot += Convert.ToInt32(n[9]);
                                        TOLCant++;
                                    }
                                    else if (n[4].Equals("DDI"))
                                    {
                                        DDIDur += Convert.ToInt32(n[5]);
                                        DDITot += Convert.ToInt32(n[9]);
                                        DDICant++;
                                    }
                                    else if (n[4].Equals("ENT"))
                                    {
                                        ENTDur += Convert.ToInt32(n[5]);
                                        ENTTot += Convert.ToInt32(n[9]);
                                        ENTCant++;
                                    }
                                    else if (n[4].Equals("EXC"))
                                    {
                                        EXCDur += Convert.ToInt32(n[5]);
                                        EXCTot += Convert.ToInt32(n[9]);
                                        EXCCant++;
                                    }
                                    else if (n[4].Equals("INT"))
                                    {
                                        INTDur += Convert.ToInt32(n[5]);
                                        INTTot += Convert.ToInt32(n[9]);
                                        INTCant++;
                                    }
                                    else if (n[4].Equals("INV"))
                                    {
                                        INVDur += Convert.ToInt32(n[5]);
                                        INVTot += Convert.ToInt32(n[9]);
                                        INVCant++;
                                    }
                                    else if (n[4].Equals("ITH"))
                                    {
                                        ITHDur += Convert.ToInt32(n[5]);
                                        ITHTot += Convert.ToInt32(n[9]);
                                        ITHCant++;
                                    }
                                    else if (n[4].Equals("SAT"))
                                    {
                                        SATDur += Convert.ToInt32(n[5]);
                                        SATTot += Convert.ToInt32(n[9]);
                                        SATCant++;
                                    }

                                    Repetido = false;
                                    if(EXT2.Count != 0)
                                    {
                                        foreach (List<string[]> g in EXT2)
                                        {
                                            if (g[0][10].Equals(n[10]))
                                            {
                                                Repetido = true;
                                            }
                                        }
                                    }
                                    if(Repetido == false)
                                    {
                                        ext = new List<string[]>();
                                        foreach (string[] l in s)
                                        {
                                            if (n[10].Equals(l[10]))
                                            {
                                                ext.Add(l);
                                            }
                                        }
                                        EXT2.Add(ext);
                                    }

                                }

                                DurRes = 0;
                                VrNetoRes = 0;
                                VrRecargoRes = 0;
                                VrIVARes = 0;
                                VrTotalRes = 0;
                                CantRes = 0;

                                foreach (List<string[]> n in EXT2)
                                {
                                    Application.DoEvents();
                                    DurGen = 0;
                                    VrNetoGen = 0;
                                    VrRecargoGen = 0;
                                    VrIvaGen = 0;
                                    VrTotalGen = 0;

                                    foreach (string[] r in n)
                                    {
                                        Application.DoEvents();
                                        DurGen += Convert.ToInt32(r[5]);
                                        VrNetoGen += Convert.ToInt32(r[6]);
                                        VrRecargoGen += Convert.ToInt32(r[7]);
                                        VrIvaGen += Convert.ToInt32(r[8]);
                                        VrTotalGen += Convert.ToInt32(r[9]);
                                    }
                                    try
                                    {
                                        using (Conexion = new MySqlConnection(conexion))
                                        {
                                            Conexion.Open();
                                            query = "select * from extensiones where Nume_Extension = ?e";
                                            comando = new MySqlCommand(query, Conexion);
                                            comando.Parameters.AddWithValue("?e", (n[0][10]));
                                            lee = comando.ExecuteReader();
                                            lee.Read();
                                            RowTotal = new string[7] { "EXT: " + n[0][10] + " " + lee["Nomb_Extension"].ToString(), n.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                                            lee.Close();
                                            Conexion.Close();
                                        }
                                    }
                                    catch
                                    {
                                        RowTotal = new string[7] { "EXT: " + n[0][10] + "Desconocida", n.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                                    }
                                    DataLlamadas.Rows.Add(RowTotal);
                                    DurRes += DurGen;
                                    VrNetoRes += VrNetoGen;
                                    VrRecargoRes += VrRecargoGen;
                                    VrIVARes += VrIvaGen;
                                    VrTotalRes += VrTotalGen;
                                    CantRes += n.Count;
                                }
                                RowTotal = new string[7] { "TOTAL", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                                DataLlamadas.Rows.Add(RowTotal);

                                pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                                AnchoPDF = new float[7] { 39f, 16f, 10f, 10f, 10f, 10f, 10f };
                                pdfTable.SetWidths(AnchoPDF);
                                pdfTable.WidthPercentage = 100;
                                pdfTable.SetWidths(AnchoPDF);
                                foreach (DataGridViewColumn column in DataLlamadas.Columns)
                                {
                                    cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                                    pdfTable.AddCell(cell);
                                }
                                foreach (DataGridViewRow row in DataLlamadas.Rows)
                                {
                                    AnchoPDFpos = 0;
                                    foreach (DataGridViewCell celda in row.Cells)
                                    {
                                        cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                        if (AnchoPDFpos == 0)
                                        {
                                            cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                        }
                                        else
                                        {
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        }
                                        pdfTable.AddCell(cell);
                                        AnchoPDFpos++;
                                    }
                                }
                                pdfDoc.Add(pdfTable);
                                pdfDoc.Add(new Paragraph("\n\n"));
                                TotalRegistros++;
                                progressBar2.Value = (int)((TotalRegistros * 100) / EXT.Count);
                            }
                        }
                    }
                }
                else
                {
                    IncrementoGen = 10;
                    Visor.RowCount = Visor.RowCount + 1;
                    DataLlamadas = new DataGridView();
                    Visor.Controls.Add(DataLlamadas);
                    DataLlamadas.ColumnCount = 7;
                    DataLlamadas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    DataLlamadas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    DataLlamadas.AutoSize = true;
                    DataLlamadas.RowHeadersVisible = false;
                    DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                    DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    DataLlamadas.BorderStyle = BorderStyle.None;
                    DataLlamadas.AllowUserToAddRows = false;
                    DataLlamadas.AllowUserToDeleteRows = false;
                    DataLlamadas.AllowUserToResizeRows = false;
                    DataLlamadas.MultiSelect = false;
                    DataLlamadas.ReadOnly = true;
                    DataLlamadas.Enabled = false;

                    if (radioButton1.Checked) { DataLlamadas.Columns[0].HeaderText = "EXTENSIÓN"; }
                    else if (radioButton2.Checked) { DataLlamadas.Columns[0].HeaderText = "CENTRO DE COSTO"; }
                    else if (radioButton3.Checked) { DataLlamadas.Columns[0].HeaderText = "TRONCAL"; }
                    else if (radioButton4.Checked) { DataLlamadas.Columns[0].HeaderText = "CÓDIGO PERSONAL"; }
                    else if (radioButton5.Checked) { DataLlamadas.Columns[0].HeaderText = "´NÚMERO MARCADO"; }
                    else if (radioButton6.Checked) { DataLlamadas.Columns[0].HeaderText = "NÚMERO DE FOLIO"; }
                    DataLlamadas.Columns[0].Width = 242;
                    DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                    DataLlamadas.Columns[1].Width = 70;
                    DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                    DataLlamadas.Columns[2].Width = 73;
                    DataLlamadas.Columns[3].HeaderText = "Vr.Neto";
                    DataLlamadas.Columns[3].Width = 77;
                    DataLlamadas.Columns[4].HeaderText = "Vr.Recargo";
                    DataLlamadas.Columns[4].Width = 77;
                    DataLlamadas.Columns[5].HeaderText = "Vr.IVA";
                    DataLlamadas.Columns[5].Width = 87;
                    DataLlamadas.Columns[6].HeaderText = "Vr.Total";
                    DataLlamadas.Columns[6].Width = 87;

                    CantRes = 0;
                    DurRes = 0;
                    VrNetoRes = 0;
                    VrRecargoRes = 0;
                    VrIVARes = 0;
                    VrTotalRes = 0;

                    TotalRegistros = 0;
                    progressBar2.Value = 0;
                    progressBar2.Maximum = 100;
                    foreach (List<string[]> s in EXT)
                    {
                        Application.DoEvents();
                        LabRes = "";
                        if (radioButton1.Checked)
                        {
                            LabRes = "EXT: " + (s[0][10]) + "    ";
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from extensiones where Nume_Extension = ?e";
                                    comando = new MySqlCommand(query, Conexion);
                                    comando.Parameters.AddWithValue("?e", (s[0][10]));
                                    lee = comando.ExecuteReader();
                                    lee.Read();
                                    LabRes += lee["Nomb_Extension"].ToString();
                                    lee.Close();
                                    Conexion.Close();
                                }
                            }
                            catch
                            {
                                LabRes += "Extension desconocida";
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            LabRes = "CENTRO: " + (s[0][12]) + "    ";
                            try
                            {
                                using (Conexion = new MySqlConnection(conexion))
                                {
                                    Conexion.Open();
                                    query = "select * from centros_costo where Codi_Centro = ?e";
                                    comando = new MySqlCommand(query, Conexion);
                                    comando.Parameters.AddWithValue("?e", (s[0][12]));
                                    lee = comando.ExecuteReader();
                                    lee.Read();
                                    LabRes += lee["Nomb_Centro"].ToString();
                                    lee.Close();
                                    Conexion.Close();
                                }
                            }
                            catch
                            {
                                LabRes += "Centro de costo desconocida";
                            }
                        }
                        else if (radioButton3.Checked) { LabRes = "TRONCAL: " + (s[0][13]); }
                        else if (radioButton4.Checked) { LabRes = "CÓDIGO: " + (s[0][14]); }
                        else if (radioButton5.Checked) { LabRes = "NÚMERO: " + (s[0][2]); }
                        else if (radioButton6.Checked) { LabRes = "FOLIO: " + (s[0][15]); }

                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            Application.DoEvents();
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);

                        }
                        RowTotal = new string[7] { LabRes, s.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        DataLlamadas.Size = new Size(485, IncrementoGen += 10);
                        DurRes += DurGen;
                        VrNetoRes += VrNetoGen;
                        VrRecargoRes += VrRecargoGen;
                        VrIVARes += VrIvaGen;
                        VrTotalRes += VrTotalGen;
                        CantRes += s.Count;
                        TotalRegistros++;
                        progressBar2.Value = (int)((TotalRegistros * 100) / EXT.Count);
                    }
                    RowTotal = new string[7] { "TOTAL", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                    DataLlamadas.Rows.Add(RowTotal);

                    pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    AnchoPDF = new float[7] { 39f, 16f, 10f, 10f, 10f, 10f, 10f };
                    pdfTable.SetWidths(AnchoPDF);
                    pdfTable.WidthPercentage = 100;
                    pdfTable.SetWidths(AnchoPDF);
                    foreach (DataGridViewColumn column in DataLlamadas.Columns)
                    {
                        cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                        pdfTable.AddCell(cell);
                    }
                    foreach (DataGridViewRow row in DataLlamadas.Rows)
                    {
                        AnchoPDFpos = 0;
                        foreach (DataGridViewCell celda in row.Cells)
                        {
                            cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                            if(AnchoPDFpos == 0)
                            {
                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                            }
                            else
                            {
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                            }
                            pdfTable.AddCell(cell);
                            AnchoPDFpos++;
                        }
                    }
                    pdfDoc.Add(pdfTable);
                    pdfDoc.Add(new Paragraph("\n\n"));
                }

                Visor.RowCount = Visor.RowCount + 1;
                lab = new Label();
                lab.Text = "TOTAL:";
                lab.AutoSize = true;
                lab.BorderStyle = BorderStyle.None;
                Visor.Controls.Add(lab);
                pdfDoc.Add(new Paragraph(lab.Text, Fuente));

                Visor.RowCount = Visor.RowCount + 1;
                DataLlamadas = new DataGridView();
                Visor.Controls.Add(DataLlamadas);
                DataLlamadas.ColumnCount = 4;
                DataLlamadas.AutoSize = true;
                DataLlamadas.RowHeadersVisible = false;
                DataLlamadas.BackgroundColor = System.Drawing.Color.White;
                DataLlamadas.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                DataLlamadas.BorderStyle = BorderStyle.None;
                DataLlamadas.AllowUserToAddRows = false;
                DataLlamadas.AllowUserToDeleteRows = false;
                DataLlamadas.AllowUserToResizeRows = false;
                DataLlamadas.MultiSelect = false;
                DataLlamadas.ReadOnly = true;
                DataLlamadas.Enabled = false;

                DataLlamadas.Columns[0].HeaderText = "Cl.Llam";
                DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                DataLlamadas.Columns[3].HeaderText = "Vr.Total";

                foreach (string n in CheckCL2)
                {
                    if (n.Equals("LOC"))
                    {
                        RowTotal = new string[4] { n, LOCCant.ToString(), LOCDur.ToString(), LOCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDN"))
                    {
                        RowTotal = new string[4] { n, DDNCant.ToString(), DDNDur.ToString(), DDNTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("CEL"))
                    {
                        RowTotal = new string[4] { n, CELCant.ToString(), CELDur.ToString(), CELTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("TOL"))
                    {
                        RowTotal = new string[4] { n, TOLCant.ToString(), TOLDur.ToString(), TOLTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDI"))
                    {
                        RowTotal = new string[4] { n, DDICant.ToString(), DDIDur.ToString(), DDITot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }

                    else if (n.Equals("ENT"))
                    {
                        RowTotal = new string[4] { n, ENTCant.ToString(), ENTDur.ToString(), ENTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("EXC"))
                    {
                        RowTotal = new string[4] { n, EXCCant.ToString(), EXCDur.ToString(), EXCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INT"))
                    {
                        RowTotal = new string[4] { n, INTCant.ToString(), INTDur.ToString(), INTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INV"))
                    {
                        RowTotal = new string[4] { n, INVCant.ToString(), INVDur.ToString(), INVTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("ITH"))
                    {
                        RowTotal = new string[4] { n, ITHCant.ToString(), ITHDur.ToString(), ITHTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("SAT"))
                    {
                        RowTotal = new string[4] { n, SATCant.ToString(), SATDur.ToString(), SATTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                }
                TotalValores = LOCTot + DDNTot + CELTot + TOLTot + DDITot + ENTTot + EXCTot + INTTot + INVTot + ITHTot + SATTot;
                TotalDuracion = LOCDur + DDNDur + CELDur + TOLDur + DDIDur + ENTDur + EXCDur + INTDur + INVDur + ITHDur + SATDur;
                TotalCantidad = LOCCant + DDNCant + CELCant + TOLCant + DDICant + ENTCant + EXCCant + INTCant + INVCant + ITHCant + SATCant;
                RowTotal = new string[4] { "TOTAL:", TotalCantidad.ToString(), TotalDuracion.ToString(), TotalValores.ToString() };
                DataLlamadas.Rows.Add(RowTotal);
                pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                pdfTable.DefaultCell.PaddingBottom = 3;
                pdfTable.DefaultCell.PaddingTop = 3;
                pdfTable.WidthPercentage = 30;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
                foreach (DataGridViewColumn column in DataLlamadas.Columns)
                {
                    cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                    pdfTable.AddCell(cell);
                }
                foreach (DataGridViewRow row in DataLlamadas.Rows)
                {
                    foreach (DataGridViewCell celda in row.Cells)
                    {
                        cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                        pdfTable.AddCell(cell);
                    }
                }
                pdfDoc.Add(pdfTable);
                pdfDoc.Add(new Paragraph("\n\n"));

                panel6.Controls.Add(Visor);
                
                pdfDoc.Close();
                stream.Close();
                MP.Hide();
                MP.label1.Text = "Enviando reporte programado, por favor espere";
                MuestraMensaje(np);
            }
        }

        public void CargaChecked2()
        {
            CheckCL2 = new List<string>();

            foreach (CheckBox s in TableCLL2)
            {
                if (s.Checked == true)
                {
                    CheckCL2.Add(s.Text.Split(' ')[0]);
                }
            }
        }

        public void CargaCheckBox2()
        {
            TableCLL2 = new List<CheckBox>();
            foreach (System.Windows.Forms.Control s in TableCL2.Controls)
            {
                if (s is CheckBox && !s.Text.Equals("TODOS"))
                {
                    TableCLL2.Add((CheckBox)s);
                }
            }
        }

        public void TerminaReporte2()
        {
            panel6.Controls.Clear();
            label196.Visible = false;
            progressBar2.Visible = false;
            button56.Visible = true;
            button56.Enabled = true;
        }

        private void button53_Click(object sender, EventArgs e)
        {
            panel6.Controls.Clear();
        }

        bool YaEncontrado = false;
        public void CargaRad()
        {
            RangoRad = new List<string>();
            foreach (string s in comboBox22.Items)
            {
                if (s.Equals(comboBox22.Text))
                {
                    YaEncontrado = true;
                }
                if (s.Equals(comboBox23.Text))
                {
                    YaEncontrado = false;
                    RangoRad.Add(s.Split(' ')[0]);
                }
                if (YaEncontrado == true)
                {
                    RangoRad.Add(s.Split(' ')[0]);
                }
            }
        }

        #endregion

        #region Correcto

        public bool Correcto2()
        {
            SelecFiltro = false;
            foreach (CheckBox s in TableCLL2)
            {
                if (s.Checked == true)
                {
                    SelecFiltro = true;
                }
            }
            if (SelecFiltro == true)
            {
                return (true);
            }
            else
            {
                MessageBox.Show("No ha seleccionado un filtro en las clases de llamada");
                return (false);
            }
        }

        #endregion

        #region Guarda y Carga

        private void button55_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton5.Checked)
                {
                    MessageBox.Show("La configuración no se puede guardar en el filtro de número marcado!");
                }
                else
                {
                    if (Correcto2() == true)
                    {
                        saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
                        saveFileDialog1.FilterIndex = 1;
                        saveFileDialog1.RestoreDirectory = true;
                        CargaChecked2();
                        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            using (StreamWriter escritor = new StreamWriter(saveFileDialog1.OpenFile()))
                            {
                                escritor.WriteLine("Reporte especifico");
                                escritor.WriteLine(dateTimePicker8.Value.ToString());
                                escritor.WriteLine(dateTimePicker7.Value.ToString());
                                escritor.WriteLine(dateTimePicker6.Value.ToString());
                                escritor.WriteLine(dateTimePicker5.Value.ToString());
                                escritor.WriteLine(checkBox6.CheckState);
                                escritor.WriteLine(checkBox5.CheckState);
                                escritor.WriteLine(checkBox4.CheckState);
                                if (radioButton1.Checked) { escritor.WriteLine(radioButton1.Text); }
                                else if (radioButton2.Checked) { escritor.WriteLine(radioButton2.Text); }
                                else if (radioButton3.Checked) { escritor.WriteLine(radioButton3.Text); }
                                else if (radioButton4.Checked) { escritor.WriteLine(radioButton4.Text); }
                                else if (radioButton6.Checked) { escritor.WriteLine(radioButton6.Text); }
                                escritor.WriteLine(comboBox22.Text);
                                escritor.WriteLine(comboBox23.Text);
                                escritor.WriteLine("Filtros CL");
                                foreach (string s in CheckCL2)
                                {
                                    escritor.WriteLine(s);
                                }
                                escritor.Close();
                            }
                            MessageBox.Show("La configuración se ha guardado con éxito");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al guardar la configuración :\n\n" + ex.ToString());
            }
        }

        List<string> FiltrosCL2;
        bool CL2 = false;
        private void button54_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            FiltrosCL2 = new List<string>();
            Linea = "";
            Lineas = new List<string>();
            CL2 = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (StreamReader lector = new StreamReader(openFileDialog1.OpenFile()))
                    {
                        while ((Linea = lector.ReadLine()) != null)
                        {
                            Lineas.Add(Linea);
                        }
                        lector.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al leer el archivo\n\n" + ex.ToString());
                }
                if (Lineas.Count != 0)
                {
                    if (Lineas[0].Equals("Reporte especifico"))
                    {
                        foreach (string s in Lineas)
                        {
                            if (s.Equals("Filtros CL"))
                            {
                                CL2 = true;
                            }
                            if (CL2 == true)
                            {
                                if (!s.Equals("Filtros CL"))
                                {
                                    FiltrosCL2.Add(s);
                                }
                            }

                        }

                        dateTimePicker8.Text = Lineas[1];
                        dateTimePicker7.Text = Lineas[2];
                        dateTimePicker6.Text = Lineas[3];
                        dateTimePicker5.Text = Lineas[4];
                        if (Lineas[5].Equals("Checked")) { checkBox6.Checked = true; } else if (Lineas[5].Equals("Unchecked")) { checkBox6.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de reportes resumen"); checkBox6.Checked = false; }
                        if (Lineas[6].Equals("Checked")) { checkBox5.Checked = true; } else if (Lineas[6].Equals("Unchecked")) { checkBox5.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de llamadas extensas"); checkBox5.Checked = false; }
                        if (Lineas[7].Equals("Checked")) { checkBox4.Checked = true; } else if (Lineas[7].Equals("Unchecked")) { checkBox4.Checked = false; } else { MessageBox.Show("Error al cargar la configuración de llamadas con valor"); checkBox4.Checked = false; }
                        if (Lineas[8].Equals("Extensiones")) { radioButton1.Checked = true; }
                        else if (Lineas[8].Equals("Centros de costo")) { radioButton2.Checked = true; }
                        else if (Lineas[8].Equals("Troncales")) { radioButton3.Checked = true; }
                        else if (Lineas[8].Equals("Códigos personales")) { radioButton4.Checked = true; }
                        else if (Lineas[8].Equals("Número de folio")) { radioButton6.Checked = true; }
                        comboBox22.Text = Lineas[9];
                        comboBox23.Text = Lineas[10];
                        if (FiltrosCL2.Count != 0)
                        {
                            foreach (string s in FiltrosCL2)
                            {
                                foreach (CheckBox n in TableCLL2)
                                {
                                    if (n.Text.Split(' ')[0].Equals(s))
                                    {
                                        n.Checked = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Error al cargar las clases de llamadas");
                        }
                        MessageBox.Show("La configuración se ha cargado");
                    }
                    else
                    {
                        MessageBox.Show("El archivo de configuración seleccionado no corresponde a una configuración de un reporte especifio");
                    }
                }
                else
                {
                    MessageBox.Show("No se ha detectado una configuración en el archivo seleccionado");
                }
            }
        }

        #endregion

        #region RadioB

        private void tabControl7_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPageIndex == 1)
            {
                radioButton1.Checked = true;
            }
            if (e.TabPageIndex == 2)
            {
                listBox1.Items.Clear();
                try
                {
                    if (File.Exists("repro.txt"))
                    {
                        using (StreamReader lector = new StreamReader("repro.txt"))
                        {
                            LineaCortaReporte = "";
                            ReportesProgramadosComp = new List<string>();
                            while ((LineaCortaReporte = lector.ReadLine()) != null)
                            {
                                ReportesProgramadosComp.Add(LineaCortaReporte);
                            }
                            lector.Close();
                        }
                        ReportesProgramadosAux = new List<string>();
                        ReportesProgramados = new List<List<string>>();
                        foreach (string s in ReportesProgramadosComp)
                        {
                            if (s.Equals("--"))
                            {
                                ReportesProgramados.Add(ReportesProgramadosAux);
                                ReportesProgramadosAux = new List<string>();
                            }
                            else
                            {
                                ReportesProgramadosAux.Add(s);
                            }
                        }
                        foreach (List<string> s in ReportesProgramados)
                        {
                            listBox1.Items.Add(s[0]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ha ocurrido un error al cargar los reportes programados!\n\n" + ex.ToString());
                    tabControl1.SelectTab(0);
                }
            }
        }

        int PosRep = 0;
        List<string> Repetidos;
        bool RepetidoB;
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                PosRep = 0;
                if (radioButton1.Checked)
                {
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            PosRep++;
                            comboBox22.Items.Add(lee["Nume_Extension"].ToString() + " " + lee["Nomb_Extension"].ToString());
                            comboBox23.Items.Add(lee["Nume_Extension"].ToString() + " " + lee["Nomb_Extension"].ToString());
                        }
                        lee.Close();
                        Conexion.Close();
                    }
                    comboBox22.Text = comboBox22.Items[0].ToString();
                    comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se hapodido conectar con la base de datos\n\n" + ex.ToString());
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                PosRep = 0;
                if (radioButton2.Checked)
                {
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from centros_costo";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            PosRep++;
                            comboBox22.Items.Add(lee["Codi_Centro"].ToString() + " " + lee["Nomb_Centro"].ToString());
                            comboBox23.Items.Add(lee["Codi_Centro"].ToString() + " " + lee["Nomb_Centro"].ToString());
                        }
                        lee.Close();
                        Conexion.Close();
                    }
                    comboBox22.Text = comboBox22.Items[0].ToString();
                    comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se hapodido conectar con la base de datos\n\n" + ex.ToString());
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                PosRep = 0;
                if (radioButton3.Checked)
                {
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from troncales";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            PosRep++;
                            comboBox22.Items.Add(lee["Line_Troncal"].ToString() + " " + lee["Nume_Acceso_Troncal"].ToString());
                            comboBox23.Items.Add(lee["Line_Troncal"].ToString() + " " + lee["Nume_Acceso_Troncal"].ToString());
                        }
                        lee.Close();
                        Conexion.Close();
                    }
                    comboBox22.Text = comboBox22.Items[0].ToString();
                    comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se hapodido conectar con la base de datos\n\n" + ex.ToString());
            }
        }
        
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            Repetidos = new List<string>();
            try
            {
                PosRep = 0;
                if (radioButton4.Checked)
                {
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from codigos_personales";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            RepetidoB = false;
                            PosRep++;
                            foreach (string s in Repetidos)
                            {
                                if (s.Equals(lee["Nomb_Cod_Personal"].ToString()))
                                {
                                    RepetidoB = true;
                                }
                            }
                            if(RepetidoB == false)
                            {
                                comboBox22.Items.Add(lee["Nomb_Cod_Personal"].ToString());
                                comboBox23.Items.Add(lee["Nomb_Cod_Personal"].ToString());
                                Repetidos.Add(lee["Nomb_Cod_Personal"].ToString());
                            }
                        }
                        lee.Close();
                        Conexion.Close();
                    }
                    comboBox22.Text = comboBox22.Items[0].ToString();
                    comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se hapodido conectar con la base de datos\n\n" + ex.ToString());
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            Repetidos = new List<string>();
            try
            {
                PosRep = 0;
                if (radioButton6.Checked)
                {
                    comboBox22.Items.Clear();
                    comboBox23.Items.Clear();
                    using (Conexion = new MySqlConnection(conexion))
                    {
                        Conexion.Open();
                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            RepetidoB = false;
                            PosRep++;
                            foreach (string s in Repetidos)
                            {
                                if (s.Equals(lee["Nume_Folio"].ToString()))
                                {
                                    RepetidoB = true;
                                }
                            }

                            if (RepetidoB == false)
                            {
                                comboBox22.Items.Add(lee["Nume_Folio"].ToString());
                                comboBox23.Items.Add(lee["Nume_Folio"].ToString());
                                Repetidos.Add(lee["Nume_Folio"].ToString());
                            }
                        }
                        lee.Close();
                        Conexion.Close();
                    }
                    comboBox22.Text = comboBox22.Items[0].ToString();
                    comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se hapodido conectar con la base de datos\n\n" + ex.ToString());
            }
        }
        bool NumeroRepetido;
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                FechaInicial = dateTimePicker8.Text.ToString();
                HoraInicial = dateTimePicker6.Text.ToString();
                FechaFinal = dateTimePicker7.Text.ToString();
                HoraFinal = dateTimePicker5.Text.ToString();
                LeePos();
                if (SePuede() == true)
                {
                    DialogResult dialogResult2 = MessageBox.Show("Para este filtro primero debe seleccionar un rango de fecha, en este momento tiene seleccionado desde: " + FechaInicial + " a las: " + HoraInicial + " Hasta: " + FechaFinal + " a las: " + HoraFinal + ". ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
                    if (dialogResult2 == DialogResult.Yes)
                    {
                        button56.Enabled = false;
                        button56.Visible = false;
                        label196.Visible = true;
                        progressBar2.Value = 0;
                        progressBar2.Maximum = 100;
                        progressBar2.Visible = true;
                        comboBox22.Items.Clear();
                        comboBox23.Items.Clear();

                        using (Conexion = new MySqlConnection(conexion))
                        {
                            Conexion.Open();
                            query = "SHOW TABLES";
                            comando = new MySqlCommand(query, Conexion);
                            lee = comando.ExecuteReader();
                            TablasNumeros = new List<string>();
                            while (lee.Read())
                            {
                                NombreTabla = "";
                                EsNumerico = false;
                                NombreTabla = lee.GetValue(0).ToString();
                                EsNumerico = int.TryParse(NombreTabla.Split(' ')[0], out Out1);
                                if (EsNumerico == true)
                                {
                                    TablasNumeros.Add(NombreTabla);
                                }
                            }
                            lee.Close();
                            Conexion.Close();
                        }
                        if (TablasNumeros.Count >= 1)
                        {
                            Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("yyyy"));
                            Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("yyyy"));
                            Out3 = Convert.ToInt32(dateTimePicker8.Value.ToString("MM"));
                            Out4 = Convert.ToInt32(dateTimePicker7.Value.ToString("MM"));
                            Filtrados = new List<string>();
                            foreach (string S in TablasNumeros)
                            {
                                EnRango = true;

                                if (Convert.ToInt32(S.Split(' ')[0]) >= Out1 && Convert.ToInt32(S.Split(' ')[0]) <= Out2)
                                {
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out1)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) < Out3)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                    if (Convert.ToInt32(S.Split(' ')[0]) == Out2)
                                    {
                                        if (Convert.ToInt32(S.Split(' ')[1]) > Out4)
                                        {
                                            EnRango = false;
                                        }
                                    }
                                }
                                else
                                {
                                    EnRango = false;
                                }

                                if (EnRango == true)
                                {
                                    Filtrados.Add(S);
                                }
                            }
                            if (Filtrados.Count >= 1)
                            {
                                TablasNumeros = Filtrados;
                                if (TablasNumeros.Count == 1)
                                {
                                    using (Conexion = new MySqlConnection(conexion))
                                    {
                                        Conexion.Open();
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("dd"));
                                        Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("dd"));
                                        Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                        Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                        HoraI = dateTimePicker6.Value.ToString("HH:mm");
                                        HoraF = dateTimePicker5.Value.ToString("HH:mm");
                                        LlamadasFiltradas = new List<string[]>();
                                        TotalRegistros = 0;
                                        while (lee.Read())
                                        {
                                            TotalRegistros++;
                                        }
                                        lee.Close();
                                        progressBar2.Value = 0;
                                        RegTot = TotalRegistros;
                                        TotalRegistros = 0;
                                        query = "select * from `" + TablasNumeros[0] + "`";
                                        comando = new MySqlCommand(query, Conexion);
                                        lee = comando.ExecuteReader();
                                        while (lee.Read())
                                        {
                                            Application.DoEvents();
                                            Minutos = 0;
                                            if (lee["Errores"].ToString().Equals("-"))
                                            {
                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1 && Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                {
                                                    if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                    {
                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                        {
                                                            if (comboBox22.Items.Count == 0)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                            else
                                                            {
                                                                NumeroRepetido = false;
                                                                foreach (string s in comboBox22.Items)
                                                                {
                                                                    if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                    {
                                                                        NumeroRepetido = true;
                                                                    }
                                                                }
                                                                if (NumeroRepetido == false)
                                                                {
                                                                    comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                    comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                }
                                                            }

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                    {

                                                        Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                        if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                        {
                                                            if (comboBox22.Items.Count == 0)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                            else
                                                            {
                                                                NumeroRepetido = false;
                                                                foreach (string s in comboBox22.Items)
                                                                {
                                                                    if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                    {
                                                                        NumeroRepetido = true;
                                                                    }
                                                                }
                                                                if (NumeroRepetido == false)
                                                                {
                                                                    comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                    comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (comboBox22.Items.Count == 0)
                                                        {
                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                        }
                                                        else
                                                        {
                                                            NumeroRepetido = false;
                                                            foreach (string s in comboBox22.Items)
                                                            {
                                                                if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                {
                                                                    NumeroRepetido = true;
                                                                }
                                                            }
                                                            if (NumeroRepetido == false)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            TotalRegistros++;
                                            progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);

                                        }
                                        lee.Close();
                                        Conexion.Close();
                                    }
                                }
                                else
                                {
                                    progressBar2.Visible = true;
                                    progressBar2.Value = 0;
                                    TotalRegistros = 0;
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        using (Conexion = new MySqlConnection(conexion))
                                        {
                                            Conexion.Open();
                                            query = "select * from `" + TablasNumeros[i] + "`";
                                            comando = new MySqlCommand(query, Conexion);
                                            lee = comando.ExecuteReader();
                                            while (lee.Read())
                                            {
                                                TotalRegistros++;
                                            }
                                            lee.Close();
                                            Conexion.Close();
                                        }
                                    }
                                    progressBar2.Maximum = 100;
                                    RegTot = TotalRegistros;
                                    TotalRegistros = 0;
                                    LlamadasFiltradas = new List<string[]>();
                                    Out1 = Convert.ToInt32(dateTimePicker8.Value.ToString("dd"));
                                    Out2 = Convert.ToInt32(dateTimePicker7.Value.ToString("dd"));
                                    Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                    Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                    HoraI = dateTimePicker6.Value.ToString("HH:mm");
                                    HoraF = dateTimePicker5.Value.ToString("HH:mm");
                                    MesI = Convert.ToInt32(dateTimePicker8.Value.ToString("MM"));
                                    MesF = Convert.ToInt32(dateTimePicker7.Value.ToString("MM"));
                                    for (int i = 0; i < TablasNumeros.Count; i++)
                                    {
                                        if (i == 0)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker8.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) >= Out1)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                                    {
                                                                        if (comboBox22.Items.Count == 0)
                                                                        {
                                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                        }
                                                                        else
                                                                        {
                                                                            NumeroRepetido = false;
                                                                            foreach (string s in comboBox22.Items)
                                                                            {
                                                                                if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                                {
                                                                                    NumeroRepetido = true;
                                                                                }
                                                                            }
                                                                            if (NumeroRepetido == false)
                                                                            {
                                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (comboBox22.Items.Count == 0)
                                                                    {
                                                                        comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                        comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                    }
                                                                    else
                                                                    {
                                                                        NumeroRepetido = false;
                                                                        foreach (string s in comboBox22.Items)
                                                                        {
                                                                            if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                            {
                                                                                NumeroRepetido = true;
                                                                            }
                                                                        }
                                                                        if (NumeroRepetido == false)
                                                                        {
                                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (comboBox22.Items.Count == 0)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                            else
                                                            {
                                                                NumeroRepetido = false;
                                                                foreach (string s in comboBox22.Items)
                                                                {
                                                                    if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                    {
                                                                        NumeroRepetido = true;
                                                                    }
                                                                }
                                                                if (NumeroRepetido == false)
                                                                {
                                                                    comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                    comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                }
                                                            }
                                                        }
                                                    }

                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else if (i == TablasNumeros.Count - 1)
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from `" + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker7.Value.ToString("MM")))
                                                        {
                                                            if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                            {
                                                                if (Convert.ToInt32(lee[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                                {
                                                                    Minutos = (Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(lee[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                    if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                                    {
                                                                        if (comboBox22.Items.Count == 0)
                                                                        {
                                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                        }
                                                                        else
                                                                        {
                                                                            NumeroRepetido = false;
                                                                            foreach (string s in comboBox22.Items)
                                                                            {
                                                                                if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                                {
                                                                                    NumeroRepetido = true;
                                                                                }
                                                                            }
                                                                            if (NumeroRepetido == false)
                                                                            {
                                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (comboBox22.Items.Count == 0)
                                                                    {
                                                                        comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                        comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                    }
                                                                    else
                                                                    {
                                                                        NumeroRepetido = false;
                                                                        foreach (string s in comboBox22.Items)
                                                                        {
                                                                            if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                            {
                                                                                NumeroRepetido = true;
                                                                            }
                                                                        }
                                                                        if (NumeroRepetido == false)
                                                                        {
                                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (comboBox22.Items.Count == 0)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                            else
                                                            {
                                                                NumeroRepetido = false;
                                                                foreach (string s in comboBox22.Items)
                                                                {
                                                                    if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                    {
                                                                        NumeroRepetido = true;
                                                                    }
                                                                }
                                                                if (NumeroRepetido == false)
                                                                {
                                                                    comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                    comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                                }
                                                            }
                                                        }
                                                    }

                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);
                                                }
                                                Conexion.Close();
                                            }
                                        }
                                        else
                                        {
                                            using (Conexion = new MySqlConnection(conexion))
                                            {
                                                Conexion.Open();
                                                query = "select * from " + TablasNumeros[i] + "`";
                                                comando = new MySqlCommand(query, Conexion);
                                                lee = comando.ExecuteReader();
                                                while (lee.Read())
                                                {
                                                    Application.DoEvents();
                                                    if (lee["Errores"].ToString().Equals("-"))
                                                    {
                                                        if (comboBox22.Items.Count == 0)
                                                        {
                                                            comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                            comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                        }
                                                        else
                                                        {
                                                            NumeroRepetido = false;
                                                            foreach (string s in comboBox22.Items)
                                                            {
                                                                if (s.Equals(lee["NNumeroMarcado"].ToString()))
                                                                {
                                                                    NumeroRepetido = true;
                                                                }
                                                            }
                                                            if (NumeroRepetido == false)
                                                            {
                                                                comboBox22.Items.Add(lee["NNumeroMarcado"]);
                                                                comboBox23.Items.Add(lee["NNumeroMarcado"]);
                                                            }
                                                        }
                                                    }
                                                    TotalRegistros++;
                                                    progressBar2.Value = (int)((TotalRegistros * 100) / RegTot);

                                                }
                                                Conexion.Close();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                comboBox22.Items.Clear();
                                comboBox22.Text = "";
                                comboBox23.Items.Clear();
                                comboBox23.Text = "";
                                radioButton5.Checked = false;
                                MessageBox.Show("No se han detectado llamadas dentro del rango especificado");
                            }

                        }
                        else
                        {
                            comboBox22.Items.Clear();
                            comboBox22.Text = "";
                            comboBox23.Items.Clear();
                            comboBox23.Text = "";
                            radioButton5.Checked = false;
                            MessageBox.Show("No se han detectado llamadas en la base de datos");
                        }
                        if (comboBox22.Items.Count > 0 && comboBox23.Items.Count > 0)
                        {
                            comboBox22.Text = comboBox22.Items[0].ToString();
                            comboBox23.Text = comboBox22.Items[comboBox22.Items.Count - 1].ToString();
                        }
                        else
                        {
                            comboBox22.Items.Clear();
                            comboBox22.Text = "";
                            comboBox23.Items.Clear();
                            comboBox23.Text = "";
                            radioButton5.Checked = false;
                            MessageBox.Show("No se ha detectado números marcados en el rango especificado");
                        }
                        TerminaReporte2();
                    }
                    else
                    {
                        comboBox22.Items.Clear();
                        comboBox22.Text = "";
                        comboBox23.Items.Clear();
                        comboBox23.Text = "";
                        radioButton5.Checked = false;
                    }
                }
            }
        }

        #endregion

        #endregion

        #region Guarda o envía
        string NombreArchivo;
        public void MuestraMensaje(string np)
        {
            NombreArchivo = np;
            Mensaje.Seleccion = ""; 
            Mensaje.Seleccionado = false;
            Mensaje ms = new Mensaje();
            ms.FormClosing += Ms_FormClosing;
            ms.ShowDialog();
        }

        private void Ms_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Mensaje.Seleccion.Equals("0"))
            {
                GuardaPDF();
            }
            else if (Mensaje.Seleccion.Equals("1"))
            {
                EnviaPDF();
            }
            else if (Mensaje.Seleccion.Equals("2"))
            {
                DescartaPDF();
            }
            Mensaje.Seleccion = "";
            Mensaje.Seleccionado = false;
            NombreArchivo = "";
        }

        public void GuardaPDF()
        {
            try
            {
                saveFileDialog1.Filter = "pdf files (*.pdf)|*.pdf";
                saveFileDialog1.FilterIndex = 1;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = NombreArchivo;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    File.Move(NombreArchivo, saveFileDialog1.FileName);
                }
                File.Delete(NombreArchivo);
                MessageBox.Show("El archivo se ha guardado con éxito");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al ejecutar la opción : \n\n" + ex.ToString());
            }
        }
        string TablaPDF;
        public void EnviaPDF()
        {
            TablaPDF = "";
            try
            {
                if(tabControl7.SelectedIndex == 0)
                {
                    TablaPDF = "Generales";
                }
                else if(tabControl7.SelectedIndex == 1)
                {
                    if (radioButton1.Checked) { TablaPDF = radioButton1.Text; }
                    else if (radioButton2.Checked) { TablaPDF = radioButton2.Text; }
                    else if (radioButton3.Checked) { TablaPDF = radioButton3.Text; }
                    else if (radioButton4.Checked) { TablaPDF = radioButton4.Text; }
                    else if (radioButton6.Checked) { TablaPDF = radioButton6.Text; }
                }
                Envio_correo ec = new Envio_correo(TablaPDF, NombreArchivo);
                ec.ShowDialog();
                File.Delete(NombreArchivo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al ejecutar la opción : \n\n" + ex.ToString());
            }
        }

        public void DescartaPDF()
        {
            try
            {
                File.Delete(NombreArchivo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al ejecutar la opción : \n\n" + ex.ToString());
            }
        }

        #endregion

        #endregion
        
        #region Cambia Tab

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPageIndex != 0)
            {
                Permitido = false;
                foreach (string s in Permisos)
                {
                    if (e.TabPageIndex == Convert.ToInt32(s.Substring(0, 1)) && s.Substring(1, 2).Equals("SI"))
                    {
                        Permitido = true;
                    }
                }

                if (Permitido == false)
                {
                    e.Cancel = true;
                    MessageBox.Show("No tienes permisos para entrar en esta sección");
                }
            }
        }

        private void tabControl7_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if(e.TabPageIndex == 2)
            {
                if(!Permisos[7].Substring(1, 2).Equals("SI"))
                {
                    e.Cancel = true;
                    MessageBox.Show("No tienes permisos para entrar en esta sección");
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                aTimer.Enabled = true;
                tabPage1.Controls.Add(label68);
                tabPage1.Controls.Add(label69);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                aTimer.Enabled = false;
                tabPage2.Controls.Add(label68);
                tabPage2.Controls.Add(label69);
                CargaTablas("0");
                dataGridView2_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, 0));
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                aTimer.Enabled = false;
                tabPage3.Controls.Add(label68);
                tabPage3.Controls.Add(label69);
                CargaInd("0");
                CargaCombo("bg");
                dataGridView3_CellClick(dataGridView2, new DataGridViewCellEventArgs(0, 0));
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                aTimer.Enabled = false;
                tabPage4.Controls.Add(label68);
                tabPage4.Controls.Add(label69);
                CargaTablasEx();
                dataGridView4_CellClick(null, new DataGridViewCellEventArgs(0, 0));
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                aTimer.Enabled = false;
                tabPage6.Controls.Add(label68);
                tabPage6.Controls.Add(label69);
                CargaTablaTF();
                dataGridView11_CellClick(null, new DataGridViewCellEventArgs(0, 0));
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                aTimer.Enabled = false;
                tabPage7.Controls.Add(label68);
                tabPage7.Controls.Add(label69);
                CargaTablaP();
                dataGridView14_CellClick(dataGridView14, new DataGridViewCellEventArgs(0, 0));
            }
            else if(tabControl1.SelectedIndex == 7)
            {
                aTimer.Enabled = false;
                if (!string.IsNullOrEmpty(IP1) || !string.IsNullOrEmpty(IP2) || !string.IsNullOrEmpty(P1) || !string.IsNullOrEmpty(P2))
                {
                    textBox60.Text = IP1;
                    textBox61.Text = P1;
                    textBox62.Text = IP2;
                    textBox63.Text = P2;
                }
            }
            else if(tabControl1.SelectedIndex == 4)
            {
                aTimer.Enabled = false;
                tabPage5.Controls.Add(label68);
                tabPage5.Controls.Add(label69);
                IniciaRep();
            }

        }

        #endregion

        #region Login

        public List<string> Permisos = new List<string>();

        private void button42_Click(object sender, EventArgs e)
        {
            if (button42.Text.Equals("Login"))
            {
                Login lg = new Login();
                lg.Show();
                lg.FormClosing += Lg_FormClosing;
            }
            else if (button42.Text.Equals("Logout"))
            {
                Login.Permisos = new List<string>();
                Login.usuario = "";
                Permisos = new List<string>();
                button42.Text = "Login";
                label180.Text = "";
            }
        }

        private void Lg_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(Login.Logeado == true)
            {
                Permisos = Login.Permisos;
                button42.Text = "Logout";
                label180.Text = Login.usuario;
            }
            else
            {
                Permisos = new List<string>();
            }
        }

        #endregion

        #region Conecta RecepDatos

        public void IniciaConRecep()
        {
            sck = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            sck.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            ConectaRecep();
        }

        Socket sck;
        EndPoint epLocal, epRemote;
        byte[] bufferExpo = new byte[1024];
        string IP1 = "";
        string IP2 = "";
        string P1 = "";
        string P2 = "";
        int size;

        private void MessageCallBack(IAsyncResult aResult)
        {
            size = 0;
            try
            {
                size = sck.EndReceiveFrom(aResult, ref epRemote);
                if (size > 0)
                {
                    byte[] recivedData = new byte[1024];
                    recivedData = (byte[])aResult.AsyncState;
                    ASCIIEncoding eEncoding = new ASCIIEncoding();
                    string recivedMessage = eEncoding.GetString(recivedData);
                    if (recivedMessage.Split('|')[0].Equals("mensaje"))
                    {
                        MessageBox.Show(recivedMessage.Split('|')[1]);
                    }

                    if (recivedMessage.Split('|')[0].Equals("formatof"))
                    {
                        FormatoFecha = recivedMessage.Split('|')[1];
                        if (string.IsNullOrEmpty(FormatoFecha))
                        {
                            MessageBox.Show("ha ocurrido un error al obtener el formato de la hora, ¿está conectado a RecepDatos?");
                            FiltrosCorretos = false;
                        }
                        else if (FormatoFecha.Equals("error"))
                        {
                            MessageBox.Show("No se ha detectado algún formato de hora!");
                            FiltrosCorretos = false;
                        }
                        else
                        {
                            FiltraFecha(FormatoFecha);
                        }
                    }

                    if (recivedMessage.Split('|')[0].Equals("formatoh"))
                    {
                        FormatoHora = recivedMessage.Split('|')[1];
                        if (string.IsNullOrEmpty(FormatoHora))
                        {
                            MessageBox.Show("ha ocurrido un error al obtener el formato de la hora, ¿está conectado a RecepDatos?");
                            FiltrosCorretos = false;
                        }
                        else if (FormatoHora.Equals("error"))
                        {
                            MessageBox.Show("No se ha detectado algún formato de hora!");
                            FiltrosCorretos = false;
                        }
                        else
                        {
                            FiltraHora(FormatoHora);
                        }
                    }
                }
                bufferExpo = new byte[1024];
                sck.BeginReceiveFrom(bufferExpo, 0, bufferExpo.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MessageCallBack), bufferExpo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al recibir un dato! compruebe que el programa esté conectado correctamente a RecepDatos\n\n" + ex.ToString());
            }
        }

        public void ConectaRecep()
        {
            if (!sck.Connected)
            {
                try
                {
                    if (string.IsNullOrEmpty(IP1) || string.IsNullOrEmpty(IP2) || string.IsNullOrEmpty (P1) || string.IsNullOrEmpty(P2))
                    {
                        MessageBox.Show("No se ha establecio una configuración para comunicarse con RecepDatos!");
                    }
                    else
                    {
                        epLocal = new IPEndPoint(IPAddress.Parse(IP1), Convert.ToInt32(P1));
                        sck.Bind(epLocal);
                        epRemote = new IPEndPoint(IPAddress.Parse(IP2), Convert.ToInt32(P2));
                        sck.Connect(epRemote);
                        sck.BeginReceiveFrom(bufferExpo, 0, bufferExpo.Length, SocketFlags.None, ref epRemote, new AsyncCallback(MessageCallBack), bufferExpo);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al emparejar con RecepDatos!\n\n" + ex.ToString());
                    MessageBox.Show("Revise la conexión con RecepDatos");
                }
            }
        }
        
        public void Envia(string Respuesta)
        {
            try
            {
                if (sck.Connected)
                {
                    ASCIIEncoding enc = new ASCIIEncoding();
                    byte[] msg = new byte[1024];
                    msg = enc.GetBytes(Respuesta);
                    sck.Send(msg);
                }
                else
                {
                    MessageBox.Show("No se ha establecido una conexión con RecepDatos!");
                }
            }
            catch
            {
                MessageBox.Show("Error al enviar un comando, compruebe que el programa esté correctamente conectado a RecepDatos");
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(textBox60.Text) || string.IsNullOrEmpty(textBox61.Text) ||
               string.IsNullOrEmpty(textBox62.Text) || string.IsNullOrEmpty(textBox63.Text))
            {
                MessageBox.Show("Hay un campo vacío!");
                IP1 = "";
                P1 = "";
                IP2 = "";
                P2 = "";
            }
            else
            {
                IP1 = textBox60.Text;
                P1 = textBox61.Text;
                IP2 = textBox62.Text;
                P2 = textBox63.Text;
                Config = new List<string>();
                Config.Add(conexion);
                Config.Add(IP1);
                Config.Add(P1);
                Config.Add(IP2);
                Config.Add(P2);

                try {
                    File.Delete(Archivo);
                    using (StreamWriter escritor = new StreamWriter(Archivo))
                    {
                        foreach (string s in Config)
                        {
                            escritor.WriteLine(s);
                        }
                        escritor.Close();
                    }
                    MessageBox.Show("Los cambios se guardaron, la aplicación se reiniciará");
                    SaleEm = true;
                    Application.Restart();
                }
                catch (Exception r)
                {
                    MessageBox.Show("ha ocurrido un error al guardar la configuración en el documento de texto, la configuración será guardad en la memoria del programa\n\n" + r);
                }
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            try
            {
                Envia("01");
            }
            catch
            {
                MessageBox.Show("Error al conectar");
            }
        }

        private void label69_Click(object sender, EventArgs e)
        {
            Permitido = false;
            foreach (string s in Permisos)
            {
                if (8 == Convert.ToInt32(s.Substring(0, 1)) && s.Substring(1, 2).Equals("SI"))
                {
                    Permitido = true;
                }
            }
            if(Permitido == true)
            {
                if (label69.Text.Equals("SI"))
                {
                    Envia("02");
                }
                else if (label69.Text.Equals("NO"))
                {
                    Envia("03");
                }
            }
            else
            {
                MessageBox.Show("No tienes permisos para detener el servicio!");
            }
        }
        
        private void button44_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox60.Text) || string.IsNullOrEmpty(textBox61.Text) ||
               string.IsNullOrEmpty(textBox62.Text) || string.IsNullOrEmpty(textBox63.Text))
            {
                MessageBox.Show("Los cambios se descartaron");
            }
            tabControl1.SelectTab(0);
        }

        #endregion

        #region Licencia y timers
        
        string Hotel = "";
        string Subject = "";
        string Body = "";

        private void Revisa()
        {
            if (RevisaLic() == false)
            {
                Subject = Hotel + " Licencia expirada";
                Body = "La licencia de " + Hotel + " expirado " + DateTime.Now.ToString("yyyy/MM/dd");
                EnvioCorreo(Subject, Body);
                SaleEm = true;
                MessageBox.Show("La licencia ha expirado!");
                Application.Exit();
            }
            else
            {
                bTimer = new System.Timers.Timer();
                bTimer.Interval = 600000;
                bTimer.Elapsed += new ElapsedEventHandler(RevisaTimer);
                bTimer.Enabled = true;

                cTimer = new System.Timers.Timer();
                cTimer.Interval = 30000;
                cTimer.Elapsed += new ElapsedEventHandler(ReporteProgramado);
                cTimer.Enabled = true;
            }
        }

        private void RevisaTimer(object source, ElapsedEventArgs e)
        {
            if (RevisaLic() == false)
            {
                Subject = Hotel + " Licencia expirada";
                Body = "La licencia de " + Hotel + " expirado " + DateTime.Now.ToString("yyyy/MM/dd");
                EnvioCorreo(Subject, Body);
                SaleEm = true;
                Application.Exit();
            }
        }

        string Line = "";
        public bool RevisaLic()
        {
            try
            {
                aTimer.Enabled = false;
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From parametros where parametro = 'Reportes Hotel'";
                    comando = new MySqlCommand(query, Conexion);
                    lee = comando.ExecuteReader();
                    lee.Read();
                    Hotel = lee["seleccion"].ToString();
                    label217.Text = "Hotel: " + Hotel;
                    Conexion.Close();
                }
                if (File.Exists(@"C:\Windows\bfsvc.txt"))
                {
                    Line = "";
                    using (StreamReader lector = new StreamReader(@"C:\Windows\bfsvc.txt"))
                    {
                        Line = lector.ReadLine();
                        string horas = DateTime.Now.ToString(@"yyyy/MM/dd");
                        if (!string.IsNullOrEmpty(Line))
                        {
                            try
                            {
                                if (descifrar(Line).Split('-')[0].Equals(Hotel))
                                {
                                    Line = descifrar(Line).Split('-')[1];
                                    label203.Text = "Licencia expira el: " + Line + "\n(yyyy/MM/dd)";
                                    if (Convert.ToInt32(Line.Split('/')[0]) > Convert.ToInt32(horas.Split('/')[0]))
                                    {
                                        lector.Close();
                                        return (true);
                                    }
                                    else if (Convert.ToInt32(Line.Split('/')[0]) < Convert.ToInt32(horas.Split('/')[0]))
                                    {
                                        lector.Close();
                                        return (false);
                                    }
                                    else if (Convert.ToInt32(Line.Split('/')[0]) == Convert.ToInt32(horas.Split('/')[0]))
                                    {
                                        if (Convert.ToInt32(Line.Split('/')[1]) > Convert.ToInt32(horas.Split('/')[1]))
                                        {
                                            lector.Close();
                                            return (true);
                                        }
                                        else if (Convert.ToInt32(Line.Split('/')[1]) < Convert.ToInt32(horas.Split('/')[1]))
                                        {
                                            lector.Close();
                                            return (false);
                                        }
                                        else if (Convert.ToInt32(Line.Split('/')[1]) == Convert.ToInt32(horas.Split('/')[1]))
                                        {
                                            if (Convert.ToInt32(Line.Split('/')[2]) > Convert.ToInt32(horas.Split('/')[2]))
                                            {
                                                lector.Close();
                                                return (true);
                                            }
                                            else if (Convert.ToInt32(Line.Split('/')[2]) < Convert.ToInt32(horas.Split('/')[2]))
                                            {
                                                lector.Close();
                                                return (false);
                                            }
                                            else if (Convert.ToInt32(Line.Split('/')[2]) == Convert.ToInt32(horas.Split('/')[2]))
                                            {
                                                lector.Close();
                                                return (true);
                                            }
                                            else
                                            {
                                                lector.Close();
                                                return (false);
                                            }
                                        }
                                        else
                                        {
                                            lector.Close();
                                            return (false);
                                        }
                                    }
                                    else
                                    {
                                        lector.Close();
                                        return (false);
                                    }
                                }
                                else
                                {
                                    lector.Close();
                                    return (false);
                                }
                            }
                            catch
                            {
                                lector.Close();
                                return (false);
                            }
                        }
                        else
                        {
                            lector.Close();
                            return (false);
                        }
                    }
                }
                else
                {
                    return (false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return (true);
            }
            finally
            {
                tabControl1.Invoke(new Action(() => {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        aTimer.Enabled = true;
                    }
                }));
            }
        }

        string clave = "llave0315RD";
        byte[] llave;
        byte[] arreglo;
        MD5CryptoServiceProvider md5;
        TripleDESCryptoServiceProvider tripledes;
        byte[] resultado;
        string Correos = "";

        public string descifrar(string cadena)
        {
            arreglo = Convert.FromBase64String(cadena);
            md5 = new MD5CryptoServiceProvider();
            llave = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(clave));
            md5.Clear();

            tripledes = new TripleDESCryptoServiceProvider();
            tripledes.Key = llave;
            tripledes.Mode = CipherMode.ECB;
            tripledes.Padding = PaddingMode.PKCS7;
            ICryptoTransform convertir = tripledes.CreateDecryptor();
            resultado = convertir.TransformFinalBlock(arreglo, 0, arreglo.Length);
            tripledes.Clear();

            return (UTF8Encoding.UTF8.GetString(resultado));
        }
        
        public void EnvioCorreo(string subject, string body)
        {
            SmtpClient client = new SmtpClient("", 0);
            NetworkCredential credentials = new NetworkCredential("", "");
            try
            {
                using (Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    query = "Select * From parametros where parametro = 'Correos ExpoDatos'";
                    comando = new MySqlCommand(query, Conexion);
                    lee = comando.ExecuteReader();
                    lee.Read();
                    Correos = lee["seleccion"].ToString();
                    Conexion.Close();
                }
                
                client = new SmtpClient("smtp.gmail.com", 587);
                client.EnableSsl = true;
                credentials = new NetworkCredential("activadorreex@gmail.com", "RD0315ED");
                client.Credentials = credentials;
                try
                {
                    MailMessage Mensaje = new MailMessage();
                    foreach(string s in Correos.Split(','))
                    {
                        Mensaje.To.Add(new MailAddress(s));
                    }
                    Mensaje.From = new MailAddress("activadorreex@gmail.com");
                    Mensaje.Subject = Hotel + ": " + subject;
                    Mensaje.Body = body;
                    try
                    {
                        client.Send(Mensaje);
                    }
                    catch
                    {
                        MessageBox.Show("Error al enviar correo e reporte");
                    }
                }
                catch
                {
                    MessageBox.Show("Error al cargar destinatarios del correo de reporte");
                }
            }
            catch
            {
                MessageBox.Show("Error al configurar correo de reporte");

            }
        }

        #endregion

        #region Reportes Programados

        List<string> ProgramadoConfig;
        List<string> ReportesProgramadosComp;
        List<List<string>> ReportesProgramados;
        List<string> ReportesProgramadosAux;
        List<List<string>> Repro;
        List<List<string>> ReproGuarda;
        List<string> repro;
        List<string> reproAux;
        string LineaRepro;
        SmtpClient client;
        NetworkCredential credentials;
        MailMessage MensajeProg;
        string Correo;
        string Contraseña;
        int TodosProg;
        string HeadProg;
        string LabProg;

        MySqlConnection ConexionProg;
        MySqlCommand ComandoProg;
        MySqlCommand ComandoProg2;
        MySqlDataReader LeeProg;
        MySqlDataReader LeeProg2;

        #region Carga y configura

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked)
            {
                label210.Enabled = false;
                numericUpDown1.Enabled = false;

                label209.Enabled = true;
                textBox72.Enabled = true;
                button58.Enabled = true;
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked)
            {
                label210.Enabled = false;
                numericUpDown1.Enabled = false;

                label209.Enabled = true;
                textBox72.Enabled = true;
                button58.Enabled = true;
            }
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked)
            {
                label210.Enabled = true;
                numericUpDown1.Enabled = true;

                label209.Enabled = true;
                textBox72.Enabled = true;
                button58.Enabled = true;
            }
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked)
            {
                label210.Enabled = false;
                numericUpDown1.Enabled = false;

                label209.Enabled = true;
                textBox72.Enabled = true;
                button58.Enabled = true;
            }
        }

        private void button58_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 1;
            ProgramadoConfig = new List<string>();
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox72.Text = openFileDialog1.FileName;
                    using (StreamReader lector = new StreamReader(openFileDialog1.OpenFile()))
                    {
                        while ((Linea = lector.ReadLine()) != null)
                        {
                            ProgramadoConfig.Add(Linea);
                        }
                        lector.Close();
                    }
                }
                catch
                {
                    textBox72.Text = "Ocurrió un error al leer el archivo";
                }
            }
        }

        int PosHoraProg;
        int PosHoraProg2;
        private void button57_Click(object sender, EventArgs e)
        {
            if ((radioButton7.Checked && string.IsNullOrEmpty(textBox72.Text)) || (radioButton9.Checked && string.IsNullOrEmpty(textBox72.Text)) || (radioButton10.Checked && (numericUpDown1.Value == 0 || string.IsNullOrEmpty(textBox72.Text))) || (string.IsNullOrEmpty(textBox68.Text) && string.IsNullOrEmpty(textBox69.Text) && string.IsNullOrEmpty(textBox70.Text) && string.IsNullOrEmpty(textBox71.Text)) || string.IsNullOrEmpty(textBox79.Text))
            {
                MessageBox.Show("Falta definir una configuración!");
            }
            else
            {
                using (StreamWriter escritor = new StreamWriter("repro.txt", true))
                {
                    escritor.WriteLine(textBox79.Text + "   " + DateTime.Now.ToString("yyyy/MM/dd"));
                    if (radioButton7.Checked) { escritor.WriteLine("diario"); } else if (radioButton9.Checked) { escritor.WriteLine("mensual"); } else if (radioButton10.Checked) { escritor.WriteLine("custom"); } else if (radioButton8.Checked) { escritor.WriteLine("diario2"); }
                    escritor.WriteLine(DateTime.Now.ToString("yyyy/MM/dd"));
                    escritor.WriteLine(DateTime.Now.ToString("yyyy/MM/dd"));
                    escritor.WriteLine(numericUpDown1.Value.ToString());
                    escritor.WriteLine("0");
                    if (string.IsNullOrEmpty(textBox68.Text)) { escritor.WriteLine("-"); } else { escritor.WriteLine(textBox68.Text); }
                    if (string.IsNullOrEmpty(textBox69.Text)) { escritor.WriteLine("-"); } else { escritor.WriteLine(textBox69.Text); }
                    if (string.IsNullOrEmpty(textBox70.Text)) { escritor.WriteLine("-"); } else { escritor.WriteLine(textBox70.Text); }
                    if (string.IsNullOrEmpty(textBox71.Text)) { escritor.WriteLine("-"); } else { escritor.WriteLine(textBox71.Text); }
                    PosHoraProg = 0;
                    PosHoraProg2 = 0;
                    using (StreamReader lector = new StreamReader(textBox72.Text))
                    {
                        Linea = "";
                        while ((Linea = lector.ReadLine()) != null)
                        {
                            if(PosHoraProg == 0)
                            {
                                if(Linea.Equals("Reporte general"))
                                {
                                    PosHoraProg2 = 9;
                                }
                                else if (Linea.Equals("Reporte especifico"))
                                {
                                    PosHoraProg2 = 12;
                                }
                            }
                            PosHoraProg++;
                            if(PosHoraProg == PosHoraProg2)
                            {
                                escritor.WriteLine(dateTimePicker9.Value.ToString("HH:mm"));
                                escritor.WriteLine(Linea);
                            }
                            else
                            {
                                escritor.WriteLine(Linea);
                            }
                            
                        }
                        lector.Close();
                    }
                    escritor.WriteLine("--");
                    escritor.Close();
                }
                MessageBox.Show("El reporte automático se ha guardado exitosamente");
                tabControl7_Selected(null, new TabControlEventArgs(tabPage31, 2, TabControlAction.Selected));
                numericUpDown1.Value = 0;
                textBox72.Text = "";
                textBox68.Text = "";
                textBox69.Text = "";
                textBox70.Text = "";
                textBox71.Text = "";
                textBox79.Text = "";
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                foreach (List<string> s in ReportesProgramados)
                {
                    if (s[0].Equals(listBox1.SelectedItem.ToString()))
                    {
                        if (s[1].Equals("diario") || s[1].Equals("mensual")) { label214.Enabled = false; textBox78.Enabled = false; } else { textBox78.Text = s[4]; }
                        textBox73.Text = s[1];
                        textBox74.Text = s[6];
                        textBox75.Text = s[7];
                        textBox76.Text = s[8];
                        textBox77.Text = s[9];
                        listBox2.Items.Clear();
                        listBox2.Items.Add(s[10]);
                        if (!s[15].Equals("Unchecked")) { listBox2.Items.Add("Reporte resumen"); }
                        if (!s[16].Equals("Unchecked")) { listBox2.Items.Add("Solo llamadas extensas"); }
                        if (!s[17].Equals("Unchecked")) { listBox2.Items.Add("Solo llamadas con valor"); }
                        if (s[10].Equals("Reporte general")) { textBox80.Text = s[18]; } else if (s[10].Equals("Reporte especifico")) { textBox80.Text = s[21]; }
                        for (int i = 19; i < s.Count; i++)
                        {
                            listBox2.Items.Add(s[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("No se ha detectado un reporte");
            }
        }

        private void button59_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < ReportesProgramados.Count; i++)
                {
                    if (ReportesProgramados[i][0].Equals(listBox1.SelectedItem.ToString()))
                    {
                        ReportesProgramados.RemoveAt(i);
                        i = ReportesProgramados.Count + 1;
                    }
                }
                File.Delete("repro.txt");
                using (StreamWriter escritor = new StreamWriter("repro.txt"))
                {
                    foreach (List<string> s in ReportesProgramados)
                    {
                        foreach (string n in s)
                        {
                            escritor.WriteLine(n);
                        }
                        escritor.WriteLine("--");
                    }
                    escritor.Close();
                }
                MessageBox.Show("El reporte programado se ha eliminado exitosamente");
                tabControl7_Selected(null, new TabControlEventArgs(tabPage31, 2, TabControlAction.Selected));
                if(listBox1.Items.Count > 0)
                {
                    listBox1.SetSelected(0, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al eliminar el reporte programado!\n\n" + ex.ToString());
            }
        }

        #endregion

        #region timer
        
        string PDFprog;
        int MesActual;
        int Año;
        int MesSiguiente;
        int PosIRep;

        private void ReporteProgramado(object source, ElapsedEventArgs e)
        {
            try
            {
                aTimer.Enabled = false;
                if (File.Exists("repro.txt"))
                {
                    Repro = new List<List<string>>();
                    repro = new List<string>();
                    LineaRepro = "";
                    using (StreamReader lector = new StreamReader("repro.txt"))
                    {
                        while ((LineaRepro = lector.ReadLine()) != null)
                        {
                            if (LineaRepro.Equals("--"))
                            {
                                if (repro.Count > 0)
                                {
                                    Repro.Add(repro);
                                }
                                repro = new List<string>();
                            }
                            else
                            {
                                repro.Add(LineaRepro);
                            }
                        }
                    }
                    if (Repro.Count > 0)
                    {
                        ReproGuarda = new List<List<string>>();
                        reproAux = new List<string>();
                        foreach (List<string> s in Repro)
                        {
                            reproAux = new List<string>();
                            foreach (string n in s)
                            {
                                reproAux.Add(n);
                            }
                            ReproGuarda.Add(reproAux);
                        }
                        PDFprog = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".pdf";
                        cTimer.Enabled = false;
                        for (int i = 0; i < Repro.Count; i++)
                        {
                            reproAux = new List<string>();
                            if (Repro[i][10].Equals("Reporte general"))
                            {
                                PosHoraProg2 = 18;
                            }
                            else if (Repro[i][10].Equals("Reporte especifico"))
                            {
                                PosHoraProg2 = 21;
                            }
                            if (Repro[i][1].Equals("diario"))
                            {
                                if (Convert.ToInt32(DateTime.Now.ToString("dd")) != Convert.ToInt32(Repro[i][2].Split('/')[2]) || Convert.ToInt32(DateTime.Now.ToString("MM")) != Convert.ToInt32(Repro[i][2].Split('/')[1]))
                                {
                                    if ((Convert.ToInt32(DateTime.Now.ToString("HH")) * 60) + Convert.ToInt32(DateTime.Now.ToString("mm")) >= (Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[0]) * 60) + Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[1]))
                                    {
                                        PosIRep = i;
                                        i = Repro.Count;
                                        foreach (string h in Repro[PosIRep])
                                        {
                                            reproAux.Add(h);
                                        }
                                        ReproGuarda[PosIRep][3] = DateTime.Now.ToString("yyyy/MM/dd");
                                        Thread Procesa = new Thread(delegate ()
                                        {
                                            CargaCheckedProg(Repro[PosIRep], Repro[PosIRep][10], reproAux[2], PDFprog, reproAux, PosIRep, reproAux[2], reproAux[PosHoraProg2]);

                                        });
                                        MP.Invoke(new Action(() => { MP.Show(); }));
                                        Application.DoEvents();
                                        this.Invoke(new Action(() => { this.Enabled = false; }));
                                        Procesa.Start();
                                        Procesa.Join();
                                        this.Invoke(new Action(() => { this.Enabled = true; }));
                                        MP.Invoke(new Action(() => { MP.Hide(); }));
                                    }
                                }
                            }
                            else if (Repro[i][1].Equals("diario2"))
                            {
                                if (Convert.ToInt32(DateTime.Now.ToString("dd")) == Convert.ToInt32(Repro[i][2].Split('/')[2]) || Convert.ToInt32(DateTime.Now.ToString("MM")) != Convert.ToInt32(Repro[i][2].Split('/')[1]))
                                {
                                    if ((Convert.ToInt32(DateTime.Now.ToString("HH")) * 60) + Convert.ToInt32(DateTime.Now.ToString("mm")) >= (Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[0]) * 60) + Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[1]))
                                    {
                                        PosIRep = i;
                                        i = Repro.Count;
                                        foreach (string h in Repro[PosIRep])
                                        {
                                            reproAux.Add(h);
                                        }
                                        ReproGuarda[PosIRep][3] = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
                                        Thread Procesa = new Thread(delegate ()
                                        {
                                            CargaCheckedProg(Repro[PosIRep], Repro[PosIRep][10], reproAux[2], PDFprog, reproAux, PosIRep, reproAux[2], reproAux[PosHoraProg2]);

                                        });
                                        MP.Invoke(new Action(() => { MP.Show(); }));
                                        Application.DoEvents();
                                        this.Invoke(new Action(() => { this.Enabled = false; }));
                                        Procesa.Start();
                                        Procesa.Join();
                                        this.Invoke(new Action(() => { this.Enabled = true; }));
                                        MP.Invoke(new Action(() => { MP.Hide(); }));
                                    }
                                }
                            }
                            else if (Repro[i][1].Equals("mensual"))
                            {
                                if (Convert.ToInt32(DateTime.Now.ToString("MM")) != Convert.ToInt32(Repro[i][2].Split('/')[1]))
                                {
                                    if ((Convert.ToInt32(DateTime.Now.ToString("HH")) * 60) + Convert.ToInt32(DateTime.Now.ToString("mm")) >= (Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[0]) * 60) + Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[1]))
                                    {
                                        PosIRep = i;
                                        i = Repro.Count;
                                        foreach (string h in Repro[PosIRep])
                                        {
                                            reproAux.Add(h);
                                        }
                                        Año = DateTime.Now.Year;
                                        MesSiguiente = DateTime.Now.Month;
                                        MesActual = DateTime.Now.Month - 1;
                                        ReproGuarda[PosIRep][3] = Convert.ToDateTime(Año + "/" + MesSiguiente).AddDays(-1).ToString("yyy/MM/dd");
                                        reproAux[2] = Convert.ToDateTime(Año + "/" + MesActual + "/01").ToString("yyyy/MM/dd");
                                        Thread Procesa = new Thread(delegate ()
                                        {
                                            CargaCheckedProg(Repro[PosIRep], Repro[PosIRep][10], ReproGuarda[PosIRep][3], PDFprog, reproAux, PosIRep, reproAux[2], reproAux[PosHoraProg2]);
                                        });
                                        MP.Invoke(new Action(() => { MP.Show(); }));
                                        Application.DoEvents();
                                        this.Invoke(new Action(() => { this.Enabled = false; }));
                                        Procesa.Start();
                                        Procesa.Join();
                                        this.Invoke(new Action(() => { this.Enabled = true; }));
                                        MP.Invoke(new Action(() => { MP.Hide(); }));
                                    }
                                }
                            }
                            else if (Repro[i][1].Equals("custom"))
                            {
                                if (Convert.ToInt32(DateTime.Now.ToString("dd")) == Convert.ToDateTime(Repro[i][2]).AddDays(Convert.ToInt32(Repro[i][4]) + 1).Day)
                                {
                                    if ((Convert.ToInt32(DateTime.Now.ToString("HH")) * 60) + Convert.ToInt32(DateTime.Now.ToString("mm")) >= (Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[0]) * 60) + Convert.ToInt32(Repro[i][PosHoraProg2].Split(':')[1]))
                                    {
                                        PosIRep = i;
                                        i = Repro.Count;
                                        foreach (string h in Repro[PosIRep])
                                        {
                                            reproAux.Add(h);
                                        }
                                        ReproGuarda[PosIRep][3] = DateTime.Now.ToString("yyyy/MM/dd");
                                        Thread Procesa = new Thread(delegate ()
                                        {
                                            CargaCheckedProg(Repro[PosIRep], Repro[PosIRep][10], DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd"), PDFprog, reproAux, PosIRep, reproAux[2], reproAux[PosHoraProg2]);
                                        });
                                        MP.Invoke(new Action(() => { MP.Show(); }));
                                        Application.DoEvents();
                                        this.Invoke(new Action(() => { this.Enabled = false; }));
                                        Procesa.Start();
                                        Procesa.Join();
                                        this.Invoke(new Action(() => { this.Enabled = true; }));
                                        MP.Invoke(new Action(() => { MP.Hide(); }));
                                    }
                                }
                            }
                        }
                        cTimer.Enabled = true;
                    }
                }
            }
            catch
            {
                Correo = "";
                Contraseña = "";
                using (ConexionProg = new MySqlConnection(ExpoDatos.conexion))
                {
                    ConexionProg.Open();
                    query = "select * from parametros where parametro = 'Correo envio ExpoDatos'";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    LeeProg.Read();
                    Correo = LeeProg["seleccion"].ToString().Split(',')[0];
                    Contraseña = LeeProg["seleccion"].ToString().Split(',')[1];
                    ConexionProg.Close();
                }

                client = new SmtpClient("", 0);
                credentials = new NetworkCredential("", "");
                client = new SmtpClient("smtp.gmail.com", 587);
                client.EnableSsl = true;
                credentials = new NetworkCredential(Correo, Contraseña);
                client.Credentials = credentials;
                client.Timeout = 50000;
                MensajeProg = new MailMessage();

                if (!reproAux[6].Equals("-")) { MensajeProg.To.Add(new MailAddress(reproAux[6])); }
                if (!reproAux[7].Equals("-")) { MensajeProg.To.Add(new MailAddress(reproAux[7])); }
                if (!reproAux[8].Equals("-")) { MensajeProg.To.Add(new MailAddress(reproAux[8])); }
                if (!reproAux[9].Equals("-")) { MensajeProg.To.Add(new MailAddress(reproAux[9])); }

                MensajeProg.Subject = "Error al enviar reporte programado";
                MensajeProg.Body = "Ha ocurrido un error al enviar un reporte programado. EL reporte se intentará enviar nuevamente";
                MensajeProg.From = new MailAddress(Correo);
                try
                {
                    client.Send(MensajeProg);
                    
                }
                catch
                {
                    MensajeProg.Dispose();
                    client.Dispose();
                }
            }
            finally
            {
                tabControl1.Invoke(new Action(() => {
                    if (tabControl1.SelectedIndex == 0)
                    {
                        aTimer.Enabled = true;
                    }
                }));
            }
        }

        #endregion

        #region Carga y filtra checked

        List<string> CheckCL2Prog;
        List<string> CheckCLProg;
        List<string> CheckCEProg;
        List<string> CheckCCProg;
        List<string> CheckTProg;
        List<string> RangoradProg;
        bool CEprog;
        bool CLprog;
        bool CCprog;
        bool Tprog;
        public void CargaCheckedProg(List<string> n, string p1, string p2, string p4, List<string> p5, int i, string p6, string p7)
        {
            CheckCL2Prog = new List<string>();
            CheckCLProg = new List<string>();
            CheckCEProg = new List<string>();
            CheckCCProg = new List<string>();
            CheckTProg = new List<string>();

            CEprog = false;
            CLprog = false;
            CCprog = false;
            Tprog = false;

            foreach (string s in n)
            {
                if (s.Equals("Filtros CE"))
                {
                    CEprog = true;
                    CLprog = false;
                    CCprog = false;
                    Tprog = false;
                }
                else if (s.Equals("Filtros CL"))
                {
                    CEprog = false;
                    CLprog = true;
                    CCprog = false;
                    Tprog = false;
                }
                else if (s.Equals("Filtros CC"))
                {
                    CEprog = false;
                    CLprog = false;
                    CCprog = true;
                    Tprog = false;
                }
                else if (s.Equals("Filtros T"))
                {
                    CEprog = false;
                    CLprog = false;
                    CCprog = false;
                    Tprog = true;
                }


                if (CEprog == true)
                {
                    if (!s.Equals("Filtros CE"))
                    {
                        CheckCEProg.Add(s);
                    }
                }
                else if (CLprog == true)
                {
                    if (!s.Equals("Filtros CL"))
                    {
                        CheckCLProg.Add(s);
                        CheckCL2Prog.Add(s);
                    }
                }
                else if (CCprog == true)
                {
                    if (!s.Equals("Filtros CC"))
                    {
                        CheckCCProg.Add(s);
                    }
                }
                else if (Tprog == true)
                {
                    if (!s.Equals("Filtros T"))
                    {
                        CheckTProg.Add(s);
                    }
                }
            }
            GeneraReporteProgramado(p1, p2, p4, p5, i, p6, p7);
        }


        bool EnRangoRadProg;
        List<string> RadProg;
        string NFprog2 = "";
        string CCprog2 = "";
        string Tprog2 = "";
        string CPprog2 = "";
        string Eprog2 = "";
        public void FiltraCheck2Prog()
        {
            if (reproAux[10].Equals("Reporte general"))
            {
                for (int i = 0; i < CheckCEProg.Count; i++)
                {
                    if (CheckCEProg[i].Equals(LeeProg["ClaseExtension"].ToString()))
                    {
                        i = CheckCEProg.Count;
                        for (int n = 0; n < CheckCLProg.Count; n++)
                        {
                            if (CheckCLProg[n].Equals(LeeProg["ClaseLlamada"].ToString()))
                            {
                                n = CheckCLProg.Count;
                                for (int m = 0; m < CheckCCProg.Count; m++)
                                {
                                    if (CheckCCProg[m].Equals(LeeProg["CentroDeCosto"].ToString()))
                                    {
                                        m = CheckCCProg.Count;
                                        for (int s = 0; s < CheckTProg.Count; s++)
                                        {
                                            if (CheckTProg[s].Equals(LeeProg["TTroncal"].ToString()))
                                            {
                                                FiltroParametro = true;
                                                if (reproAux[16].Equals("Checked"))
                                                {
                                                    using (ConexionProg = new MySqlConnection(conexion))
                                                    {
                                                        ConexionProg.Open();
                                                        query = "select * from parametros where parametro = 'Llamadas extensas general'";
                                                        ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                                        LeeProg2 = ComandoProg2.ExecuteReader();
                                                        LeeProg2.Read();
                                                        if ((Convert.ToInt32(LeeProg["DuracionLlamadaAproximada"].ToString())) < (Convert.ToInt32(LeeProg2["seleccion"].ToString())))
                                                        {
                                                            FiltroParametro = false;
                                                        }
                                                        LeeProg2.Close();
                                                        ConexionProg.Close();
                                                    }
                                                }
                                                if (reproAux[17].Equals("Checked"))
                                                {
                                                    using (ConexionProg = new MySqlConnection(conexion))
                                                    {
                                                        ConexionProg.Open();
                                                        query = "select * from parametros where parametro = 'Llamadas con valor general'";
                                                        ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                                        LeeProg2 = ComandoProg2.ExecuteReader();
                                                        LeeProg2.Read();
                                                        if ((Convert.ToInt32(LeeProg["ValorTotal"].ToString())) < (Convert.ToInt32(LeeProg2["seleccion"].ToString())))
                                                        {
                                                            FiltroParametro = false;
                                                        }
                                                        LeeProg2.Close();
                                                        ConexionProg.Close();
                                                    }

                                                }

                                                if (FiltroParametro == true)
                                                {
                                                    s = CheckTProg.Count;
                                                    llamadasFil[0] = LeeProg["FFechaFinalLlamada"].ToString();
                                                    llamadasFil[1] = LeeProg["HHoraFinalLlamada"].ToString();
                                                    llamadasFil[2] = LeeProg["NNumeroMarcado"].ToString();
                                                    llamadasFil[3] = LeeProg["Destino"].ToString();
                                                    llamadasFil[4] = LeeProg["ClaseLlamada"].ToString();
                                                    llamadasFil[5] = LeeProg["DuracionLlamadaAproximada"].ToString();
                                                    llamadasFil[6] = LeeProg["ValorLlamadaTarifa"].ToString();
                                                    llamadasFil[7] = LeeProg["RecargoServicioValor"].ToString();
                                                    llamadasFil[8] = LeeProg["ValorIVA"].ToString();
                                                    llamadasFil[9] = LeeProg["ValorTotal"].ToString();
                                                    llamadasFil[10] = LeeProg["EExtension"].ToString();
                                                    try
                                                    {
                                                        using (ConexionProg = new MySqlConnection(conexion))
                                                        {
                                                            ConexionProg.Open();
                                                            query = "select * from extensiones where Nume_Extension = '" + llamadasFil[10] + "'";
                                                            ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                                            LeeProg2 = ComandoProg2.ExecuteReader();
                                                            LeeProg2.Read();
                                                            llamadasFil[11] = LeeProg["Nomb_Extension"].ToString();
                                                            LeeProg2.Close();
                                                            ConexionProg.Close();
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        llamadasFil[11] = "Extension desconocida: " + llamadasFil[10];
                                                    }
                                                    LlamadasFiltradas.Add(llamadasFil);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (reproAux[10].Equals("Reporte especifico"))
            {
                for (int s = 0; s < CheckCL2Prog.Count; s++)
                {
                    if (CheckCL2Prog[s].Equals(LeeProg["ClaseLlamada"].ToString()))
                    {
                        FiltroParametro = true;
                        if (reproAux[16].Equals("Checked"))
                        {
                            using (ConexionProg = new MySqlConnection(conexion))
                            {
                                ConexionProg.Open();
                                query = "select * from parametros where parametro = 'Llamadas extensas especificos'";
                                ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                LeeProg2 = ComandoProg2.ExecuteReader();
                                LeeProg2.Read();
                                if ((Convert.ToInt32(LeeProg["DuracionLlamadaAproximada"].ToString())) < (Convert.ToInt32(LeeProg2["seleccion"].ToString())))
                                {
                                    FiltroParametro = false;
                                }
                                LeeProg2.Close();
                                ConexionProg.Close();
                            }
                        }
                        if (reproAux[17].Equals("Checked"))
                        {
                            using (ConexionProg = new MySqlConnection(conexion))
                            {
                                ConexionProg.Open();
                                query = "select * from parametros where parametro = 'Llamadas con valor especificos'";
                                ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                LeeProg2 = ComandoProg2.ExecuteReader();
                                LeeProg2.Read();
                                if ((Convert.ToInt32(LeeProg["ValorTotal"].ToString())) < (Convert.ToInt32(LeeProg2["seleccion"].ToString())))
                                {
                                    FiltroParametro = false;
                                }
                                LeeProg2.Close();
                                ConexionProg.Close();
                            }

                        }

                        NFprog2 = "";
                        CCprog2 = "";
                        Tprog2 = "";
                        CPprog2 = "";
                        Eprog2 = "";
                        if (FiltroParametro == true)
                        {
                            using (ConexionProg = new MySqlConnection(conexion))
                            {
                                ConexionProg.Open();

                                if (reproAux[18].Equals("Extensiones")) { query = "select * from extensiones"; }
                                else if (reproAux[18].Equals("Centros de costo")) { query = "select * from centros_costo"; }
                                else if (reproAux[18].Equals("Troncales")) { query = "select * from troncales"; }
                                else if (reproAux[18].Equals("Códigos personales")) { query = "select * from codigos_personales"; }
                                else if (reproAux[18].Equals("Número de folio")) { query = "select * from extensiones"; }
                                ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                LeeProg2 = ComandoProg2.ExecuteReader();

                                RadProg = new List<string>();
                                while (LeeProg2.Read())
                                {
                                    EnRangoRadProg = false;

                                    NFprog2 = "";
                                    CCprog2 = "";
                                    Tprog2 = "";
                                    CPprog2 = "";
                                    Eprog2 = "";

                                    if (reproAux[18].Equals("Extensiones")) { Eprog2 = LeeProg2["Nume_Extension"].ToString(); }
                                    else if (reproAux[18].Equals("Centros de costo")) { CCprog2 = LeeProg2["Codi_Centro"].ToString(); }
                                    else if (reproAux[18].Equals("Troncales")) { Tprog2 = LeeProg2["Line_Troncal"].ToString(); }
                                    else if (reproAux[18].Equals("Códigos personales")) { CPprog2 = LeeProg2["Codi_Personal"].ToString(); }
                                    else if (reproAux[18].Equals("Número de folio")) { NFprog2 = LeeProg2["Nume_Folio"].ToString(); }

                                    if (RadProg.Count != 0)
                                    {
                                        foreach (string l in RadProg)
                                        {
                                            if (reproAux[18].Equals("Extensiones")) { if (l.Equals(Eprog2)) { EnRangoRadProg = true; } }
                                            else if (reproAux[18].Equals("Centros de costo")) { if (l.Equals(CCprog2)) { EnRangoRadProg = true; } }
                                            else if (reproAux[18].Equals("Troncales")) { if (l.Equals(Tprog2)) { EnRangoRadProg = true; } }
                                            else if (reproAux[18].Equals("Códigos personales")) { if (l.Equals(CPprog2)) { EnRangoRadProg = true; } }
                                            else if (reproAux[18].Equals("Número de folio")) { if (l.Equals(NFprog2)) { EnRangoRadProg = true; } }

                                        }
                                    }

                                    if (EnRangoRadProg == false)
                                    {
                                        if (reproAux[18].Equals("Extensiones")) { RadProg.Add(Eprog2); }
                                        else if (reproAux[18].Equals("Centros de costo")) { RadProg.Add(CCprog2); }
                                        else if (reproAux[18].Equals("Troncales")) { RadProg.Add(Tprog2); }
                                        else if (reproAux[18].Equals("Códigos personales")) { RadProg.Add(CPprog2); }
                                        else if (reproAux[18].Equals("Número de folio")) { RadProg.Add(NFprog2); }
                                    }
                                }
                                LeeProg2.Close();
                                RangoradProg = new List<string>();
                                EnRangoRadProg = false;
                                for (int i = 0; i < RadProg.Count; i++)
                                {
                                    if (RadProg[i].Equals(reproAux[19].Split(' ')[0]))
                                    {
                                        EnRangoRadProg = true;
                                    }
                                    if (EnRangoRadProg == true)
                                    {
                                        RangoradProg.Add(RadProg[i]);
                                    }
                                    if (RadProg[i].Equals(reproAux[20].Split(' ')[0]))
                                    {
                                        EnRangoRadProg = false;
                                        i = RadProg.Count;
                                    }
                                }
                                ConexionProg.Close();
                            }
                            EncontradoRad = false;
                            if (reproAux[18].Equals("Extensiones"))
                            {
                                foreach (string n in RangoradProg)
                                {
                                    if (n.Equals(LeeProg["EExtension"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (reproAux[18].Equals("Centros de costo"))
                            {
                                foreach (string n in RangoradProg)
                                {
                                    if (n.Equals(LeeProg["CentroDeCosto"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (reproAux[18].Equals("Troncales"))
                            {
                                foreach (string n in RangoradProg)
                                {
                                    if (n.Equals(LeeProg["TTroncal"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (reproAux[18].Equals("Número de folio"))
                            {
                                foreach (string n in RangoradProg)
                                {
                                    if (n.Equals(LeeProg["NumeFolio"].ToString()))
                                    {
                                        EncontradoRad = true;
                                    }
                                }
                            }
                            else if (reproAux[18].Equals("Códigos personales"))
                            {
                                foreach (string n in RangoradProg)
                                {
                                    using (ConexionProg = new MySqlConnection(conexion))
                                    {
                                        ConexionProg.Open();
                                        query = "select * from codigos_personales where Nomb_Cod_Personal = '" + n + "'";
                                        ComandoProg2 = new MySqlCommand(query, ConexionProg);
                                        LeeProg2 = ComandoProg2.ExecuteReader();
                                        while (LeeProg2.Read())
                                        {
                                            if (LeeProg2["Codi_Personal"].ToString().Equals(LeeProg["PCodigoPersonal"].ToString()))
                                            {
                                                s = CheckCL2Prog.Count;
                                                llamadasFil[0] = LeeProg["FFechaFinalLlamada"].ToString();
                                                llamadasFil[1] = LeeProg["HHoraFinalLlamada"].ToString();
                                                llamadasFil[2] = LeeProg["NNumeroMarcado"].ToString();
                                                llamadasFil[3] = LeeProg["Destino"].ToString();
                                                llamadasFil[4] = LeeProg["ClaseLlamada"].ToString();
                                                llamadasFil[5] = LeeProg["DuracionLlamadaAproximada"].ToString();
                                                llamadasFil[6] = LeeProg["ValorLlamadaTarifa"].ToString();
                                                llamadasFil[7] = LeeProg["RecargoServicioValor"].ToString();
                                                llamadasFil[8] = LeeProg["ValorIVA"].ToString();
                                                llamadasFil[9] = LeeProg["ValorTotal"].ToString();
                                                llamadasFil[10] = LeeProg["EExtension"].ToString();
                                                llamadasFil[11] = "--";
                                                llamadasFil[12] = LeeProg["CentroDeCosto"].ToString();
                                                llamadasFil[13] = LeeProg["TTroncal"].ToString();
                                                llamadasFil[14] = n;
                                                llamadasFil[15] = LeeProg["NumeFolio"].ToString();
                                                LlamadasFiltradas.Add(llamadasFil);
                                            }
                                        }
                                        LeeProg2.Close();
                                        ConexionProg.Close();
                                    }
                                }
                                EncontradoRad = false;
                            }
                            if (EncontradoRad == true)
                            {
                                s = CheckCL2Prog.Count;
                                llamadasFil[0] = LeeProg["FFechaFinalLlamada"].ToString();
                                llamadasFil[1] = LeeProg["HHoraFinalLlamada"].ToString();
                                llamadasFil[2] = LeeProg["NNumeroMarcado"].ToString();
                                llamadasFil[3] = LeeProg["Destino"].ToString();
                                llamadasFil[4] = LeeProg["ClaseLlamada"].ToString();
                                llamadasFil[5] = LeeProg["DuracionLlamadaAproximada"].ToString();
                                llamadasFil[6] = LeeProg["ValorLlamadaTarifa"].ToString();
                                llamadasFil[7] = LeeProg["RecargoServicioValor"].ToString();
                                llamadasFil[8] = LeeProg["ValorIVA"].ToString();
                                llamadasFil[9] = LeeProg["ValorTotal"].ToString();
                                llamadasFil[10] = LeeProg["EExtension"].ToString();
                                llamadasFil[11] = "--";
                                llamadasFil[12] = LeeProg["CentroDeCosto"].ToString();
                                llamadasFil[13] = LeeProg["TTroncal"].ToString();
                                llamadasFil[14] = LeeProg["PCodigoPersonal"].ToString();
                                llamadasFil[15] = LeeProg["NumeFolio"].ToString();
                                LlamadasFiltradas.Add(llamadasFil);

                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region filtro fecha

        string LineaCortaReporte;
        public void GeneraReporteProgramado(string tipo, string FechaFinalPro, string PDFname, List<string> ReproAux, int ps, string FechaInicialPro, string HoraFinalPro)
        {
            if (string.IsNullOrEmpty(FormatoFechaFinal))
            {
                Envia("04");
            }
            if (string.IsNullOrEmpty(FormatoHoraFinal))
            {
                Envia("06");
            }
            if (!string.IsNullOrEmpty(FormatoFechaFinal) && !string.IsNullOrEmpty(FormatoHoraFinal))
            {
                if (string.IsNullOrEmpty(PosDia) || string.IsNullOrEmpty(PosHora) || string.IsNullOrEmpty(PosMinutos))
                {
                    LeePos();
                }
                if (SePuede() == true)
                {
                    using (ConexionProg = new MySqlConnection(conexion))
                    {
                        ConexionProg.Open();
                        query = "SHOW TABLES";
                        ComandoProg = new MySqlCommand(query, ConexionProg);
                        LeeProg = ComandoProg.ExecuteReader();
                        TablasNumeros = new List<string>();
                        while (LeeProg.Read())
                        {
                            NombreTabla = "";
                            EsNumerico = false;
                            NombreTabla = LeeProg.GetValue(0).ToString();
                            EsNumerico = int.TryParse(NombreTabla.Split(' ')[0], out Out1);
                            if (EsNumerico == true)
                            {
                                TablasNumeros.Add(NombreTabla);
                            }
                        }
                        LeeProg.Close();
                        ConexionProg.Close();
                    }
                    if (TablasNumeros.Count >= 1)
                    {
                        Out1 = Convert.ToInt32(FechaInicialPro.Split('/')[0]);
                        Out2 = Convert.ToInt32(FechaFinalPro.Split('/')[0]);
                        Out3 = Convert.ToInt32(FechaInicialPro.Split('/')[1]);
                        Out4 = Convert.ToInt32(FechaFinalPro.Split('/')[1]);
                        Filtrados = new List<string>();
                        foreach (string S in TablasNumeros)
                        {
                            EnRango = true;
                            if (Convert.ToInt32(S.Split(' ')[0]) >= Out1 && Convert.ToInt32(S.Split(' ')[0]) <= Out2)
                            {
                                if (Convert.ToInt32(S.Split(' ')[0]) == Out1)
                                {
                                    if (Convert.ToInt32(S.Split(' ')[1]) < Out3)
                                    {
                                        EnRango = false;
                                    }
                                }
                                if (Convert.ToInt32(S.Split(' ')[0]) == Out2)
                                {
                                    if (Convert.ToInt32(S.Split(' ')[1]) > Out4)
                                    {
                                        EnRango = false;
                                    }
                                }
                            }
                            else
                            {
                                EnRango = false;
                            }

                            if (EnRango == true)
                            {
                                Filtrados.Add(S);
                            }
                        }
                        if (Filtrados.Count >= 1)
                        {
                            TablasNumeros = Filtrados;
                            if (TablasNumeros.Count == 1)
                            {
                                using (ConexionProg = new MySqlConnection(conexion))
                                {
                                    ConexionProg.Open();
                                    Out1 = Convert.ToInt32(FechaInicialPro.Split('/')[2]);
                                    Out2 = Convert.ToInt32(FechaFinalPro.Split('/')[2]);
                                    Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                    Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                    HoraI = "00:00";
                                    if (ReproAux[1].Equals("diario2")) { HoraF = HoraFinalPro; } else { HoraF = "23:59"; }
                                    LlamadasFiltradas = new List<string[]>();
                                    query = "select * from `" + TablasNumeros[0] + "`";
                                    ComandoProg = new MySqlCommand(query, ConexionProg);
                                    LeeProg = ComandoProg.ExecuteReader();
                                    while (LeeProg.Read())
                                    {
                                        Application.DoEvents();
                                        Minutos = 0;
                                        if (LeeProg["Errores"].ToString().Equals("-"))
                                        {
                                            if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) >= Out1 && Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                            {
                                                if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out1 && Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                {
                                                    Minutos = (Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                    if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])) && Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2Prog();
                                                    }
                                                }
                                                else if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                {
                                                    Minutos = (Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                    if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2Prog();
                                                    }
                                                }
                                                else if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                {

                                                    Minutos = (Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                    if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2Prog();
                                                    }
                                                }
                                                else
                                                {
                                                    llamadasFil = new string[16];
                                                    FiltraCheck2Prog();
                                                }
                                            }
                                        }

                                    }
                                    LeeProg.Close();
                                    ConexionProg.Close();
                                }
                                if (tipo.Equals("Reporte general"))
                                {
                                    MuestraReporteProg(Directory.GetCurrentDirectory() + @"\" + PDFname, FechaFinalPro, ReproAux, ps, HoraF);
                                }
                                else if (tipo.Equals("Reporte especifico"))
                                {
                                    MuestraReporte2Prog(Directory.GetCurrentDirectory() + @"\" + PDFname, FechaFinalPro, ReproAux, ps, HoraF);
                                }
                            }
                            else
                            {
                                LlamadasFiltradas = new List<string[]>();
                                Out1 = Convert.ToInt32(FechaInicialPro.Split('/')[2]);
                                Out2 = Convert.ToInt32(FechaFinalPro.Split('/')[2]);
                                Out3 = Convert.ToInt32(PosDia.Split('-')[0]);
                                Out4 = Convert.ToInt32(PosDia.Split('-')[1]);
                                HoraI = "00:00";
                                if (ReproAux[1].Equals("diario2")) { HoraF = HoraFinalPro; } else { HoraF = "23:59"; }
                                MesI = Convert.ToInt32(FechaInicialPro.Split('/')[1]);
                                MesF = Convert.ToInt32(FechaFinalPro.Split('/')[2]);
                                for (int i = 0; i < TablasNumeros.Count; i++)
                                {
                                    if (i == 0)
                                    {
                                        using (ConexionProg = new MySqlConnection(conexion))
                                        {
                                            ConexionProg.Open();
                                            query = "select * from `" + TablasNumeros[i] + "`";
                                            ComandoProg = new MySqlCommand(query, ConexionProg);
                                            LeeProg = ComandoProg.ExecuteReader();
                                            while (LeeProg.Read())
                                            {
                                                Application.DoEvents();
                                                if (LeeProg["Errores"].ToString().Equals("-"))
                                                {
                                                    if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker8.Value.ToString("MM")))
                                                    {
                                                        if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) >= Out1)
                                                        {
                                                            if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out1)
                                                            {
                                                                Minutos = (Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                if (Minutos >= ((Convert.ToInt32(HoraI.Split(':')[0]) * 60) + Convert.ToInt32(HoraI.Split(':')[1])))
                                                                {
                                                                    llamadasFil = new string[16];
                                                                    FiltraCheck2Prog();
                                                                }
                                                            }
                                                            else
                                                            {
                                                                llamadasFil = new string[16];
                                                                FiltraCheck2Prog();
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2Prog();
                                                    }
                                                }
                                            }
                                            ConexionProg.Close();
                                        }
                                    }
                                    else if (i == TablasNumeros.Count - 1)
                                    {
                                        using (ConexionProg = new MySqlConnection(conexion))
                                        {
                                            ConexionProg.Open();
                                            query = "select * from `" + TablasNumeros[i] + "`";
                                            ComandoProg = new MySqlCommand(query, ConexionProg);
                                            LeeProg = ComandoProg.ExecuteReader();
                                            while (LeeProg.Read())
                                            {
                                                Application.DoEvents();
                                                if (LeeProg["Errores"].ToString().Equals("-"))
                                                {
                                                    if (Convert.ToInt32(TablasNumeros[i].Split(' ')[1]) == Convert.ToInt32(dateTimePicker7.Value.ToString("MM")))
                                                    {
                                                        if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) <= Out2)
                                                        {
                                                            if (Convert.ToInt32(LeeProg[FechaRow].ToString().Substring(Out3, Out4)) == Out2)
                                                            {
                                                                Minutos = (Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosHora.Split('-')[0]), Convert.ToInt32(PosHora.Split('-')[1]))) * 60) + Convert.ToInt32(LeeProg[HoraRow].ToString().Substring(Convert.ToInt32(PosMinutos.Split('-')[0]), Convert.ToInt32(PosMinutos.Split('-')[1])));
                                                                if (Minutos <= ((Convert.ToInt32(HoraF.Split(':')[0]) * 60) + Convert.ToInt32(HoraF.Split(':')[1])))
                                                                {
                                                                    llamadasFil = new string[16];
                                                                    FiltraCheck2Prog();
                                                                }
                                                            }
                                                            else
                                                            {
                                                                llamadasFil = new string[16];
                                                                FiltraCheck2Prog();
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        llamadasFil = new string[16];
                                                        FiltraCheck2Prog();
                                                    }
                                                }
                                            }
                                            ConexionProg.Close();
                                        }
                                    }
                                    else
                                    {
                                        using (ConexionProg = new MySqlConnection(conexion))
                                        {
                                            ConexionProg.Open();
                                            query = "select * from " + TablasNumeros[i] + "`";
                                            ComandoProg = new MySqlCommand(query, ConexionProg);
                                            LeeProg = ComandoProg.ExecuteReader();
                                            while (LeeProg.Read())
                                            {
                                                Application.DoEvents();
                                                if (LeeProg["Errores"].ToString().Equals("-"))
                                                {
                                                    llamadasFil = new string[16];
                                                    FiltraCheck2Prog();
                                                }

                                            }
                                            ConexionProg.Close();
                                        }
                                    }
                                }

                                if (tipo.Equals("Reporte general"))
                                {
                                    MuestraReporteProg(Directory.GetCurrentDirectory() + @"\" + PDFname, FechaFinalPro, ReproAux, ps , HoraF);
                                }
                                else if (tipo.Equals("Reporte especifico"))
                                {
                                    MuestraReporte2Prog(Directory.GetCurrentDirectory() + @"\" + PDFname, FechaFinalPro, ReproAux, ps, HoraF);
                                }
                            }
                        }
                        else
                        {
                            if (ReproAux[1].Equals("diario2"))
                            {
                                ReproGuarda[ps][2] = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
                            }
                            else
                            {
                                ReproGuarda[ps][2] = DateTime.Now.ToString("yyyy/MM/dd");
                            }
                            File.Delete("repro.txt");
                            using (StreamWriter escritor = new StreamWriter("repro.txt"))
                            {
                                foreach (List<string> c in ReproGuarda)
                                {
                                    foreach (string n in c)
                                    {
                                        escritor.WriteLine(n);
                                    }
                                    escritor.WriteLine("--");
                                }
                                escritor.Close();
                            }
                        }
                    }
                    else
                    {
                        if (ReproAux[1].Equals("diario2"))
                        {
                            ReproGuarda[ps][2] = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
                        }
                        else
                        {
                            ReproGuarda[ps][2] = DateTime.Now.ToString("yyyy/MM/dd");
                        }
                        File.Delete("repro.txt");
                        using (StreamWriter escritor = new StreamWriter("repro.txt"))
                        {
                            foreach (List<string> c in ReproGuarda)
                            {
                                foreach (string n in c)
                                {
                                    escritor.WriteLine(n);
                                }
                                escritor.WriteLine("--");
                            }
                            escritor.Close();
                        }
                    }
                }

            }
        }

        #endregion
        
        #region MuestraReporte

        public void MuestraReporteProg(string np, string FFPro, List<string> ReproAux, int i, string HHFin)
        {
            using (FileStream stream = new FileStream(np, FileMode.Create))
            {
                pdfDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                Fuente = new iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7, iTextSharp.text.Font.NORMAL);

                Filtrados = new List<string>();
                HeadProg = "";
                try
                {
                    using (ConexionProg = new MySqlConnection(conexion))
                    {
                        ConexionProg.Open();
                        query = "select * from parametros where parametro = 'Reportes Hotel'";
                        ComandoProg = new MySqlCommand(query, ConexionProg);
                        LeeProg = ComandoProg.ExecuteReader();
                        LeeProg.Read();
                        HeadProg = LeeProg["seleccion"].ToString();
                        LeeProg.Close();
                        ConexionProg.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problema al conectarse con la base de datos\n\n" + ex.ToString());
                }
                HeadProg += "\n\nREPORTE GENERAL DE LLAMADAS" + "\nDesde: " + ReproAux[2] + " a las: " + "00:00" + " Hasta: " + FFPro + " a las: " + HHFin + "\n\n Extensiones: ";

                TodosProg = 0;
                using (ConexionProg = new MySqlConnection(conexion))
                {
                    ConexionProg.Open();
                    query = "select * from clase_extensiones";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    while (LeeProg.Read())
                    {
                        TodosProg++;
                    }
                    LeeProg.Close();
                    ConexionProg.Close();
                }

                if (CheckCEProg.Count == TodosProg)
                {
                    HeadProg += "Todos";
                }
                else
                {
                    foreach (string s in CheckCEProg)
                    {
                        HeadProg += s + ", ";
                    }
                }
                HeadProg += "\nCentros de costo: ";

                TodosProg = 0;
                using (ConexionProg = new MySqlConnection(conexion))
                {
                    ConexionProg.Open();
                    query = "select * from centros_costo";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    while (LeeProg.Read())
                    {
                        TodosProg++;
                    }
                    LeeProg.Close();
                    ConexionProg.Close();
                }

                if (CheckCCProg.Count == TodosProg)
                {
                    HeadProg += "Todos";
                }
                else
                {
                    foreach (string s in CheckCCProg)
                    {
                        HeadProg += s + ", ";
                    }
                }
                HeadProg += "\nClases de llamada: ";

                TodosProg = 0;
                using (ConexionProg = new MySqlConnection(conexion))
                {
                    ConexionProg.Open();
                    query = "select * from clase_llamadas";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    while (LeeProg.Read())
                    {
                        TodosProg++;
                    }
                    LeeProg.Close();
                    ConexionProg.Close();
                }

                if (CheckCLProg.Count == TodosProg)
                {
                    HeadProg += "Todos";
                }
                else
                {
                    foreach (string s in CheckCLProg)
                    {
                        HeadProg += s + ", ";
                    }
                }
                HeadProg += "\nTroncales: ";

                TodosProg = 0;
                using (ConexionProg = new MySqlConnection(conexion))
                {
                    ConexionProg.Open();
                    query = "select * from troncales";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    while (LeeProg.Read())
                    {
                        TodosProg++;
                    }
                    LeeProg.Close();
                    ConexionProg.Close();
                }

                if (CheckTProg.Count == TodosProg)
                {
                    HeadProg += "Todos";
                }
                else
                {
                    foreach (string s in CheckTProg)
                    {
                        HeadProg += s + ", ";
                    }
                }
                if (ReproAux[15].Equals("Checked"))
                {
                    HeadProg += "\n\nREPORTE RESUMIDO\n\n";
                }
                else
                {
                    HeadProg += "\n\nREPORTE DETALLADO\n\n";
                }
                pdfDoc.Add(new Paragraph(HeadProg, Fuente));
                EXT = new List<List<string[]>>();
                ext = new List<string[]>();

                for (int a = 1; a < LlamadasFiltradas.Count; a++)
                {
                    for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                    {
                        if (Convert.ToInt32(LlamadasFiltradas[b - 1][10]) > Convert.ToInt32(LlamadasFiltradas[b][10]))
                        {
                            t = LlamadasFiltradas[b - 1];
                            LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                            LlamadasFiltradas[b] = t;
                        }
                    }
                }
                foreach (string[] s in LlamadasFiltradas)
                {
                    Repetido = false;
                    if (EXT.Count > 0)
                    {
                        foreach (List<string[]> l in EXT)
                        {
                            foreach (string[] c in l)
                            {
                                if (c[10].Equals(s[10]))
                                {
                                    Repetido = true;
                                }
                            }
                        }
                        if (Repetido == false)
                        {
                            foreach (string[] n in LlamadasFiltradas)
                            {
                                if (n[10].Equals(s[10]))
                                {
                                    ext.Add(n);
                                }
                            }
                            if (ext.Count > 0)
                            {
                                EXT.Add(ext);
                                ext = new List<string[]>();
                            }
                        }
                    }
                    else
                    {
                        foreach (string[] n in LlamadasFiltradas)
                        {
                            if (n[10].Equals(s[10]))
                            {
                                ext.Add(n);
                            }
                        }
                        if (ext.Count > 0)
                        {
                            EXT.Add(ext);
                            ext = new List<string[]>();
                        }
                    }
                }

                LOCDur = 0; LOCTot = 0; LOCCant = 0;
                DDNDur = 0; DDNTot = 0; DDNCant = 0;
                CELDur = 0; CELTot = 0; CELCant = 0;
                TOLDur = 0; TOLTot = 0; TOLCant = 0;
                DDIDur = 0; DDITot = 0; DDICant = 0;
                ENTDur = 0; ENTTot = 0; ENTCant = 0;
                EXCDur = 0; EXCTot = 0; EXCCant = 0;
                INTDur = 0; INTTot = 0; INTCant = 0;
                INVDur = 0; INVTot = 0; INVCant = 0;
                ITHDur = 0; ITHTot = 0; ITHCant = 0;
                SATDur = 0; SATTot = 0; SATCant = 0;
                TotalValores = 0;
                TotalDuracion = 0;
                TotalCantidad = 0;
                if (ReproAux[15].Equals("Unchecked"))
                {
                    foreach (List<string[]> s in EXT)
                    {
                        LabProg = "";
                        LabProg = "EXT: " + (s[0][10]) + "    ";
                        try
                        {
                            using (ConexionProg = new MySqlConnection(conexion))
                            {
                                ConexionProg.Open();
                                query = "select * from extensiones where Nume_Extension = ?e";
                                ComandoProg = new MySqlCommand(query, ConexionProg);
                                ComandoProg.Parameters.AddWithValue("?e", (s[0][10]));
                                LeeProg = ComandoProg.ExecuteReader();
                                LeeProg.Read();
                                LabProg += LeeProg["Nomb_Extension"].ToString();
                                CodiCentro = LeeProg["Codi_Centro"].ToString();
                                LabProg += "           CENTRO: " + CodiCentro + " ";
                                LeeProg.Close();
                                ConexionProg.Close();
                                try
                                {
                                    using (ConexionProg = new MySqlConnection(conexion))
                                    {
                                        ConexionProg.Open();
                                        query = "select * from centros_costo where Codi_Centro = ?e";
                                        ComandoProg = new MySqlCommand(query, ConexionProg);
                                        ComandoProg.Parameters.AddWithValue("?e", CodiCentro);
                                        LeeProg = ComandoProg.ExecuteReader();
                                        LeeProg.Read();
                                        LabProg += LeeProg["Nomb_Centro"].ToString();
                                        LeeProg.Close();
                                        ConexionProg.Close();
                                    }
                                }
                                catch
                                {
                                    LabProg += "Nombre de centro desconocido";
                                }
                            }
                        }
                        catch
                        {
                            LabProg += "Extension desconocida";
                            LabProg += "Centro de costo desconocido";
                        }
                        pdfDoc.Add(new Paragraph(LabProg, Fuente));
                        DataLlamadas = new DataGridView();
                        DataLlamadas.ColumnCount = 10;

                        DataLlamadas.Columns[0].HeaderText = "FECHA";
                        DataLlamadas.Columns[0].Width = 57;
                        DataLlamadas.Columns[1].HeaderText = "HORA";
                        DataLlamadas.Columns[1].Width = 57;
                        DataLlamadas.Columns[2].HeaderText = "NUM.MARCADO";
                        DataLlamadas.Columns[2].Width = 110;
                        DataLlamadas.Columns[3].HeaderText = "DESTINO";
                        DataLlamadas.Columns[3].Width = 97;
                        DataLlamadas.Columns[4].HeaderText = "Cl.Llam";
                        DataLlamadas.Columns[4].Width = 67;
                        DataLlamadas.Columns[5].HeaderText = "DUR";
                        DataLlamadas.Columns[5].Width = 52;
                        DataLlamadas.Columns[6].HeaderText = "Vr.neto";
                        DataLlamadas.Columns[6].Width = 67;
                        DataLlamadas.Columns[7].HeaderText = "Vr.Recargo";
                        DataLlamadas.Columns[7].Width = 67;
                        DataLlamadas.Columns[8].HeaderText = "Vr.IVA";
                        DataLlamadas.Columns[8].Width = 67;
                        DataLlamadas.Columns[9].HeaderText = "Vr.Total";
                        DataLlamadas.Columns[9].Width = 73;


                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);
                            DataLlamadas.Rows.Add(n);
                        }
                        RowTotal = new string[10] { "TOTAL:", " ", " ", " ", " ", DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                        AnchoPDF = new float[10] { 8.77f, 7.48f, 20, 20, 8.77f, 5.60f, 9f, 9f, 9f, 9f };
                        pdfTable.SetWidths(AnchoPDF);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SetWidths(AnchoPDF);
                        foreach (DataGridViewColumn column in DataLlamadas.Columns)
                        {
                            cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                            pdfTable.AddCell(cell);
                        }
                        foreach (DataGridViewRow row in DataLlamadas.Rows)
                        {
                            foreach (DataGridViewCell celda in row.Cells)
                            {
                                if (celda.Value == null)
                                {
                                    celda.Value = " ";
                                }
                                cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                pdfTable.AddCell(cell);
                            }
                        }
                        pdfDoc.Add(pdfTable);
                        pdfDoc.Add(new Paragraph("\n\n"));
                    }
                }
                else
                {
                    DataLlamadas = new DataGridView();
                    DataLlamadas.ColumnCount = 8;


                    DataLlamadas.Columns[0].HeaderText = "EXTENSIÓN";
                    DataLlamadas.Columns[0].Width = 210;
                    DataLlamadas.Columns[1].HeaderText = "C.COSTO";
                    DataLlamadas.Columns[1].Width = 62;
                    DataLlamadas.Columns[2].HeaderText = "CANTIDAD";
                    DataLlamadas.Columns[2].Width = 65;
                    DataLlamadas.Columns[3].HeaderText = "DURACIÓN";
                    DataLlamadas.Columns[3].Width = 68;
                    DataLlamadas.Columns[4].HeaderText = "Vr.Neto";
                    DataLlamadas.Columns[4].Width = 72;
                    DataLlamadas.Columns[5].HeaderText = "Vr.Recargo";
                    DataLlamadas.Columns[5].Width = 72;
                    DataLlamadas.Columns[6].HeaderText = "Vr.IVA";
                    DataLlamadas.Columns[6].Width = 82;
                    DataLlamadas.Columns[7].HeaderText = "Vr.Total";
                    DataLlamadas.Columns[7].Width = 82;

                    CantRes = 0;
                    DurRes = 0;
                    VrNetoRes = 0;
                    VrRecargoRes = 0;
                    VrIVARes = 0;
                    VrTotalRes = 0;

                    foreach (List<string[]> s in EXT)
                    {
                        LabRes = "";
                        CentroCostoRes = "";
                        LabRes = "EXT: " + (s[0][10]) + "    ";
                        try
                        {
                            using (ConexionProg = new MySqlConnection(conexion))
                            {
                                ConexionProg.Open();
                                query = "select * from extensiones where Nume_Extension = ?e";
                                ComandoProg = new MySqlCommand(query, ConexionProg);
                                ComandoProg.Parameters.AddWithValue("?e", (s[0][10]));
                                LeeProg = ComandoProg.ExecuteReader();
                                LeeProg.Read();
                                LabRes += LeeProg["Nomb_Extension"].ToString();
                                CentroCostoRes = LeeProg["Codi_Centro"].ToString();
                                LeeProg.Close();
                                ConexionProg.Close();
                            }
                        }
                        catch
                        {
                            LabRes += "Extension desconocida";
                            CentroCostoRes = "Centro de costo desconocido";
                        }

                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);

                        }
                        RowTotal = new string[8] { LabRes, CentroCostoRes, s.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        DataLlamadas.Size = new Size(485, IncrementoGen += 10);
                        DurRes += DurGen;
                        VrNetoRes += VrNetoGen;
                        VrRecargoRes += VrRecargoGen;
                        VrIVARes += VrIvaGen;
                        VrTotalRes += VrTotalGen;
                        CantRes += s.Count;
                    }
                    RowTotal = new string[8] { "TOTAL", " ", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                    DataLlamadas.Rows.Add(RowTotal);
                    pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    AnchoPDF = new float[8] { 29f, 20f, 8f, 8f, 10f, 10f, 10f, 10f };
                    pdfTable.SetWidths(AnchoPDF);
                    pdfTable.WidthPercentage = 100;
                    pdfTable.SetWidths(AnchoPDF);
                    foreach (DataGridViewColumn column in DataLlamadas.Columns)
                    {
                        cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                        pdfTable.AddCell(cell);
                    }
                    foreach (DataGridViewRow row in DataLlamadas.Rows)
                    {
                        AnchoPDFpos = 0;
                        foreach (DataGridViewCell celda in row.Cells)
                        {
                            if (celda.Value == null)
                            {
                                celda.Value = " ";
                            }
                            cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                            if (AnchoPDFpos == 0)
                            {
                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                            }
                            else
                            {
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                            }
                            pdfTable.AddCell(cell);
                            AnchoPDFpos++;
                        }
                    }
                    pdfDoc.Add(pdfTable);
                    pdfDoc.Add(new Paragraph("\n\n"));
                }

                LabProg = "";
                LabProg = "TOTAL:";
                pdfDoc.Add(new Paragraph(LabProg, Fuente));

                DataLlamadas = new DataGridView();
                DataLlamadas.ColumnCount = 4;

                DataLlamadas.Columns[0].HeaderText = "Cl.Llam";
                DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                DataLlamadas.Columns[3].HeaderText = "Vr.Total";

                foreach (string n in CheckCLProg)
                {
                    if (n.Equals("LOC"))
                    {
                        RowTotal = new string[4] { n, LOCCant.ToString(), LOCDur.ToString(), LOCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDN"))
                    {
                        RowTotal = new string[4] { n, DDNCant.ToString(), DDNDur.ToString(), DDNTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("CEL"))
                    {
                        RowTotal = new string[4] { n, CELCant.ToString(), CELDur.ToString(), CELTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("TOL"))
                    {
                        RowTotal = new string[4] { n, TOLCant.ToString(), TOLDur.ToString(), TOLTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDI"))
                    {
                        RowTotal = new string[4] { n, DDICant.ToString(), DDIDur.ToString(), DDITot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }

                    else if (n.Equals("ENT"))
                    {
                        RowTotal = new string[4] { n, ENTCant.ToString(), ENTDur.ToString(), ENTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("EXC"))
                    {
                        RowTotal = new string[4] { n, EXCCant.ToString(), EXCDur.ToString(), EXCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INT"))
                    {
                        RowTotal = new string[4] { n, INTCant.ToString(), INTDur.ToString(), INTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INV"))
                    {
                        RowTotal = new string[4] { n, INVCant.ToString(), INVDur.ToString(), INVTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("ITH"))
                    {
                        RowTotal = new string[4] { n, ITHCant.ToString(), ITHDur.ToString(), ITHTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("SAT"))
                    {
                        RowTotal = new string[4] { n, SATCant.ToString(), SATDur.ToString(), SATTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                }
                TotalValores = LOCTot + DDNTot + CELTot + TOLTot + DDITot + ENTTot + EXCTot + INTTot + INVTot + ITHTot + SATTot;
                TotalDuracion = LOCDur + DDNDur + CELDur + TOLDur + DDIDur + ENTDur + EXCDur + INTDur + INVDur + ITHDur + SATDur;
                TotalCantidad = LOCCant + DDNCant + CELCant + TOLCant + DDICant + ENTCant + EXCCant + INTCant + INVCant + ITHCant + SATCant;
                RowTotal = new string[4] { "TOTAL:", TotalCantidad.ToString(), TotalDuracion.ToString(), TotalValores.ToString() };
                DataLlamadas.Rows.Add(RowTotal);
                pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                pdfTable.DefaultCell.PaddingBottom = 3;
                pdfTable.DefaultCell.PaddingTop = 3;
                pdfTable.WidthPercentage = 30;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
                foreach (DataGridViewColumn column in DataLlamadas.Columns)
                {
                    cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                    pdfTable.AddCell(cell);
                }
                foreach (DataGridViewRow row in DataLlamadas.Rows)
                {
                    foreach (DataGridViewCell celda in row.Cells)
                    {
                        if (celda.Value == null)
                        {
                            celda.Value = " ";
                        }
                        cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                        pdfTable.AddCell(cell);
                    }
                }
                pdfDoc.Add(pdfTable);
                pdfDoc.Add(new Paragraph("\n\n"));

                pdfDoc.Close();
                stream.Close();

                Correo = "";
                Contraseña = "";
                using (ConexionProg = new MySqlConnection(ExpoDatos.conexion))
                {
                    ConexionProg.Open();
                    query = "select * from parametros where parametro = 'Correo envio ExpoDatos'";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    LeeProg.Read();
                    Correo = LeeProg["seleccion"].ToString().Split(',')[0];
                    Contraseña = LeeProg["seleccion"].ToString().Split(',')[1];
                    ConexionProg.Close();
                }


                client = new SmtpClient("", 0);
                credentials = new NetworkCredential("", "");
                client = new SmtpClient("smtp.gmail.com", 587);
                client.EnableSsl = true;
                credentials = new NetworkCredential(Correo, Contraseña);
                client.Credentials = credentials;
                client.Timeout = 50000;
                MensajeProg = new MailMessage();

                if (!ReproAux[6].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[6])); }
                if (!ReproAux[7].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[7])); }
                if (!ReproAux[8].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[8])); }
                if (!ReproAux[9].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[9])); }

                MensajeProg.Subject = ReproAux[0].Substring(0, ReproAux[0].Length - 10);
                MensajeProg.From = new MailAddress(Correo);

                MensajeProg.Attachments.Add(new Attachment(Path.GetFileName(np)));
                try
                {
                    client.Send(MensajeProg);
                    MensajeProg.Dispose();
                    client.Dispose();
                }
                catch
                {
                    MensajeProg.Attachments.Clear();
                    MensajeProg.Dispose();
                    client.Dispose();
                    File.Delete(np);
                }
                File.Delete(np);
                if (ReproAux[1].Equals("diario2"))
                {
                    ReproGuarda[i][2] = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
                }
                else
                {
                    ReproGuarda[i][2] = DateTime.Now.ToString("yyyy/MM/dd");
                }
                File.Delete("repro.txt");
                using (StreamWriter escritor = new StreamWriter("repro.txt"))
                {
                    foreach (List<string> c in ReproGuarda)
                    {
                        foreach (string n in c)
                        {
                            escritor.WriteLine(n);
                        }
                        escritor.WriteLine("--");
                    }
                    escritor.Close();
                }
            }
        }
        
        #endregion

        #region MuestraReporte2

        public void MuestraReporte2Prog(string np, string FFPro, List<string> ReproAux, int i, string HHpro)
        {
            using (FileStream stream = new FileStream(np, FileMode.Create))
            {
                pdfDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                Fuente = new iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7, iTextSharp.text.Font.NORMAL);

                Filtrados = new List<string>();
                HeadProg = "";
                try
                {
                    using (ConexionProg = new MySqlConnection(conexion))
                    {
                        ConexionProg.Open();
                        query = "select * from parametros where parametro = 'Reportes Hotel'";
                        ComandoProg = new MySqlCommand(query, ConexionProg);
                        LeeProg = ComandoProg.ExecuteReader();
                        LeeProg.Read();
                        HeadProg = LeeProg["seleccion"].ToString();
                        LeeProg.Close();
                        ConexionProg.Close();
                    }
                }
                catch
                {
                    HeadProg = "Error al obtener el hotel";
                }
                HeadProg += "\n\nREPORTE ESPECÍFICO DE LLAMADAS" + "\nDesde: " + ReproAux[2] + " a las: " + "00:00" + " Hasta: " + FFPro + " a las: " + HHpro + "\n\n Clase de llamadas: ";

                TodosProg = 0;
                using (ConexionProg = new MySqlConnection(conexion))
                {
                    ConexionProg.Open();
                    query = "select * from clase_llamadas";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    while (LeeProg.Read())
                    {
                        TodosProg++;
                    }
                    LeeProg.Close();
                    ConexionProg.Close();
                }

                if (CheckCL2Prog.Count == TodosProg)
                {
                    HeadProg += "Todos";
                }
                else
                {
                    foreach (string s in CheckCL2Prog)
                    {
                        HeadProg += s + ", ";
                    }
                }
                HeadProg += "\n\nREPORTE: " + ReproAux[18];
                HeadProg += "\nDesde: " + ReproAux[19] + " Hasta: " + ReproAux[20] + "\n";
                if (ReproAux[15].Equals("Checked"))
                {
                    HeadProg += "\nREPORTE RESUMIDO\n\n";
                }
                else
                {
                    HeadProg += "\nREPORTE DETALLADO\n\n";
                }
                pdfDoc.Add(new Paragraph(HeadProg, Fuente));
                EXT = new List<List<string[]>>();
                ext = new List<string[]>();

                if (ReproAux[18].Equals("Extensiones"))
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][10]) > Convert.ToInt32(LlamadasFiltradas[b][10]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                else if (ReproAux[18].Equals("Centros de costo"))
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][12]) > Convert.ToInt32(LlamadasFiltradas[b][12]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                else if (ReproAux[18].Equals("Troncales"))
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][13]) > Convert.ToInt32(LlamadasFiltradas[b][13]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                else if (ReproAux[18].Equals("Número de folio"))
                {
                    for (int a = 1; a < LlamadasFiltradas.Count; a++)
                    {
                        for (int b = LlamadasFiltradas.Count - 1; b >= a; b--)
                        {
                            if (Convert.ToInt32(LlamadasFiltradas[b - 1][15]) > Convert.ToInt32(LlamadasFiltradas[b][15]))
                            {
                                t = LlamadasFiltradas[b - 1];
                                LlamadasFiltradas[b - 1] = LlamadasFiltradas[b];
                                LlamadasFiltradas[b] = t;
                            }
                        }
                    }
                }
                foreach (string[] s in LlamadasFiltradas)
                {
                    Application.DoEvents();
                    Repetido = false;
                    if (EXT.Count > 0)
                    {
                        foreach (List<string[]> l in EXT)
                        {
                            foreach (string[] c in l)
                            {
                                if (ReproAux[18].Equals("Extensiones")) { if (c[10].Equals(s[10])) { Repetido = true; } }
                                else if (ReproAux[18].Equals("Centros de costo")) { if (c[12].Equals(s[12])) { Repetido = true; } }
                                else if (ReproAux[18].Equals("Troncales")) { if (c[13].Equals(s[13])) { Repetido = true; } }
                                else if (ReproAux[18].Equals("Códigos personales")) { if (c[14].Equals(s[14])) { Repetido = true; } }
                                else if (ReproAux[18].Equals("Número de folio")) { if (c[15].Equals(s[15])) { Repetido = true; } }
                            }
                        }
                        if (Repetido == false)
                        {
                            foreach (string[] n in LlamadasFiltradas)
                            {
                                if (ReproAux[18].Equals("Extensiones")) { if (n[10].Equals(s[10])) { ext.Add(n); } }
                                else if (ReproAux[18].Equals("Centros de costo")) { if (n[12].Equals(s[12])) { ext.Add(n); } }
                                else if (ReproAux[18].Equals("Troncales")) { if (n[13].Equals(s[13])) { ext.Add(n); } }
                                else if (ReproAux[18].Equals("Códigos personales")) { if (n[14].Equals(s[14])) { ext.Add(n); } }
                                else if (ReproAux[18].Equals("Número de folio")) { if (n[15].Equals(s[15])) { ext.Add(n); } }
                            }
                            if (ext.Count > 0)
                            {
                                EXT.Add(ext);
                                ext = new List<string[]>();
                            }
                        }
                    }
                    else
                    {
                        foreach (string[] n in LlamadasFiltradas)
                        {
                            if (ReproAux[18].Equals("Extensiones")) { if (n[10].Equals(s[10])) { ext.Add(n); } }
                            else if (ReproAux[18].Equals("Centros de costo")) { if (n[12].Equals(s[12])) { ext.Add(n); } }
                            else if (ReproAux[18].Equals("Troncales")) { if (n[13].Equals(s[13])) { ext.Add(n); } }
                            else if (ReproAux[18].Equals("Códigos personales")) { if (n[14].Equals(s[14])) { ext.Add(n); } }
                            else if (ReproAux[18].Equals("Número de folio")) { if (n[15].Equals(s[15])) { ext.Add(n); } }
                        }
                        if (ext.Count > 0)
                        {
                            EXT.Add(ext);
                            ext = new List<string[]>();
                        }
                    }
                }

                LOCDur = 0; LOCTot = 0; LOCCant = 0;
                DDNDur = 0; DDNTot = 0; DDNCant = 0;
                CELDur = 0; CELTot = 0; CELCant = 0;
                TOLDur = 0; TOLTot = 0; TOLCant = 0;
                DDIDur = 0; DDITot = 0; DDICant = 0;
                ENTDur = 0; ENTTot = 0; ENTCant = 0;
                EXCDur = 0; EXCTot = 0; EXCCant = 0;
                INTDur = 0; INTTot = 0; INTCant = 0;
                INVDur = 0; INVTot = 0; INVCant = 0;
                ITHDur = 0; ITHTot = 0; ITHCant = 0;
                SATDur = 0; SATTot = 0; SATCant = 0;
                TotalValores = 0;
                TotalDuracion = 0;
                TotalCantidad = 0;

                if (ReproAux[15].Equals("Unchecked"))
                {
                    if (!ReproAux[18].Equals("Centros de costo"))
                    {
                        foreach (List<string[]> s in EXT)
                        {
                            Application.DoEvents();
                            LabProg = "";
                            if (ReproAux[18].Equals("Extensiones"))
                            {
                                LabProg = "EXT: " + (s[0][10]) + "    ";
                                try
                                {
                                    using (ConexionProg = new MySqlConnection(conexion))
                                    {
                                        ConexionProg.Open();
                                        query = "select * from extensiones where Nume_Extension = ?e";
                                        ComandoProg = new MySqlCommand(query, ConexionProg);
                                        ComandoProg.Parameters.AddWithValue("?e", (s[0][10]));
                                        LeeProg = ComandoProg.ExecuteReader();
                                        LeeProg.Read();
                                        LabProg += LeeProg["Nomb_Extension"].ToString();
                                        LeeProg.Close();
                                        ConexionProg.Close();
                                    }
                                }
                                catch
                                {
                                    LabProg += "Extension desconocida";
                                }
                            }
                            else if (ReproAux[18].Equals("Troncales")) { LabProg = "TRONCAL: " + (s[0][13]); }
                            else if (ReproAux[18].Equals("Códigos personales")) { LabProg = "CÓDIGO: " + (s[0][14]); }
                            else if (ReproAux[18].Equals("Número de folio")) { LabProg = "FOLIO: " + (s[0][15]); }

                            pdfDoc.Add(new Paragraph(LabProg, Fuente));
                            DataLlamadas = new DataGridView();
                            DataLlamadas.ColumnCount = 10;

                            DataLlamadas.Columns[0].HeaderText = "FECHA";
                            DataLlamadas.Columns[0].Width = 57;
                            DataLlamadas.Columns[1].HeaderText = "HORA";
                            DataLlamadas.Columns[1].Width = 57;
                            DataLlamadas.Columns[2].HeaderText = "NUM.MARCADO";
                            DataLlamadas.Columns[2].Width = 110;
                            DataLlamadas.Columns[3].HeaderText = "DESTINO";
                            DataLlamadas.Columns[3].Width = 97;
                            DataLlamadas.Columns[4].HeaderText = "Cl.Llam";
                            DataLlamadas.Columns[4].Width = 67;
                            DataLlamadas.Columns[5].HeaderText = "DUR";
                            DataLlamadas.Columns[5].Width = 52;
                            DataLlamadas.Columns[6].HeaderText = "Vr.neto";
                            DataLlamadas.Columns[6].Width = 67;
                            DataLlamadas.Columns[7].HeaderText = "Vr.Recargo";
                            DataLlamadas.Columns[7].Width = 67;
                            DataLlamadas.Columns[8].HeaderText = "Vr.IVA";
                            DataLlamadas.Columns[8].Width = 67;
                            DataLlamadas.Columns[9].HeaderText = "Vr.Total";
                            DataLlamadas.Columns[9].Width = 73;


                            DurGen = 0;
                            VrNetoGen = 0;
                            VrRecargoGen = 0;
                            VrIvaGen = 0;
                            VrTotalGen = 0;

                            foreach (string[] n in s)
                            {
                                Application.DoEvents();
                                if (n[4].Equals("LOC"))
                                {
                                    LOCDur += Convert.ToInt32(n[5]);
                                    LOCTot += Convert.ToInt32(n[9]);
                                    LOCCant++;
                                }
                                else if (n[4].Equals("DDN"))
                                {
                                    DDNDur += Convert.ToInt32(n[5]);
                                    DDNTot += Convert.ToInt32(n[9]);
                                    DDNCant++;
                                }
                                else if (n[4].Equals("CEL"))
                                {
                                    CELDur += Convert.ToInt32(n[5]);
                                    CELTot += Convert.ToInt32(n[9]);
                                    CELCant++;
                                }
                                else if (n[4].Equals("TOL"))
                                {
                                    TOLDur += Convert.ToInt32(n[5]);
                                    TOLTot += Convert.ToInt32(n[9]);
                                    TOLCant++;
                                }
                                else if (n[4].Equals("DDI"))
                                {
                                    DDIDur += Convert.ToInt32(n[5]);
                                    DDITot += Convert.ToInt32(n[9]);
                                    DDICant++;
                                }
                                else if (n[4].Equals("ENT"))
                                {
                                    ENTDur += Convert.ToInt32(n[5]);
                                    ENTTot += Convert.ToInt32(n[9]);
                                    ENTCant++;
                                }
                                else if (n[4].Equals("EXC"))
                                {
                                    EXCDur += Convert.ToInt32(n[5]);
                                    EXCTot += Convert.ToInt32(n[9]);
                                    EXCCant++;
                                }
                                else if (n[4].Equals("INT"))
                                {
                                    INTDur += Convert.ToInt32(n[5]);
                                    INTTot += Convert.ToInt32(n[9]);
                                    INTCant++;
                                }
                                else if (n[4].Equals("INV"))
                                {
                                    INVDur += Convert.ToInt32(n[5]);
                                    INVTot += Convert.ToInt32(n[9]);
                                    INVCant++;
                                }
                                else if (n[4].Equals("ITH"))
                                {
                                    ITHDur += Convert.ToInt32(n[5]);
                                    ITHTot += Convert.ToInt32(n[9]);
                                    ITHCant++;
                                }
                                else if (n[4].Equals("SAT"))
                                {
                                    SATDur += Convert.ToInt32(n[5]);
                                    SATTot += Convert.ToInt32(n[9]);
                                    SATCant++;
                                }

                                DurGen += Convert.ToInt32(n[5]);
                                VrNetoGen += Convert.ToInt32(n[6]);
                                VrRecargoGen += Convert.ToInt32(n[7]);
                                VrIvaGen += Convert.ToInt32(n[8]);
                                VrTotalGen += Convert.ToInt32(n[9]);
                                DataLlamadas.Rows.Add(n);
                            }
                            RowTotal = new string[10] { "TOTAL:", " ", " ", " ", " ", DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                            DataLlamadas.Rows.Add(RowTotal);

                            pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                            AnchoPDF = new float[10] { 8.77f, 7.48f, 20, 20, 8.77f, 5.60f, 9f, 9f, 9f, 9f };
                            pdfTable.SetWidths(AnchoPDF);
                            pdfTable.WidthPercentage = 100;
                            pdfTable.SetWidths(AnchoPDF);
                            foreach (DataGridViewColumn column in DataLlamadas.Columns)
                            {
                                cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                                pdfTable.AddCell(cell);
                            }
                            foreach (DataGridViewRow row in DataLlamadas.Rows)
                            {
                                foreach (DataGridViewCell celda in row.Cells)
                                {
                                    if (celda.Value == null)
                                    {
                                        celda.Value = " ";
                                    }
                                    cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                    pdfTable.AddCell(cell);
                                }
                            }
                            pdfDoc.Add(pdfTable);
                            pdfDoc.Add(new Paragraph("\n\n"));
                        }
                    }
                    else
                    {
                        foreach (List<string[]> s in EXT)
                        {
                            Application.DoEvents();
                            LabProg = "";
                            LabProg = "CENTRO: " + (s[0][12]) + "    ";
                            try
                            {
                                using (ConexionProg = new MySqlConnection(conexion))
                                {
                                    ConexionProg.Open();
                                    query = "select * from centros_costo where Codi_Centro = ?e";
                                    ComandoProg = new MySqlCommand(query, ConexionProg);
                                    ComandoProg.Parameters.AddWithValue("?e", (s[0][12]));
                                    LeeProg = ComandoProg.ExecuteReader();
                                    LeeProg.Read();
                                    LabProg += LeeProg["Nomb_Centro"].ToString();
                                    LeeProg.Close();
                                    ConexionProg.Close();
                                }
                            }
                            catch
                            {
                                LabProg += "Centro de costo desconocida";
                            }
                            pdfDoc.Add(new Paragraph(LabProg, Fuente));
                            DataLlamadas = new DataGridView();
                            DataLlamadas.ColumnCount = 7;

                            DataLlamadas.Columns[0].HeaderText = "EXTENSIÓN";
                            DataLlamadas.Columns[0].Width = 242;
                            DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                            DataLlamadas.Columns[1].Width = 70;
                            DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                            DataLlamadas.Columns[2].Width = 73;
                            DataLlamadas.Columns[3].HeaderText = "Vr.Neto";
                            DataLlamadas.Columns[3].Width = 77;
                            DataLlamadas.Columns[4].HeaderText = "Vr.Recargo";
                            DataLlamadas.Columns[4].Width = 77;
                            DataLlamadas.Columns[5].HeaderText = "Vr.IVA";
                            DataLlamadas.Columns[5].Width = 87;
                            DataLlamadas.Columns[6].HeaderText = "Vr.Total";
                            DataLlamadas.Columns[6].Width = 87;

                            EXT2 = new List<List<string[]>>();
                            foreach (string[] n in s)
                            {
                                Application.DoEvents();
                                if (n[4].Equals("LOC"))
                                {
                                    LOCDur += Convert.ToInt32(n[5]);
                                    LOCTot += Convert.ToInt32(n[9]);
                                    LOCCant++;
                                }
                                else if (n[4].Equals("DDN"))
                                {
                                    DDNDur += Convert.ToInt32(n[5]);
                                    DDNTot += Convert.ToInt32(n[9]);
                                    DDNCant++;
                                }
                                else if (n[4].Equals("CEL"))
                                {
                                    CELDur += Convert.ToInt32(n[5]);
                                    CELTot += Convert.ToInt32(n[9]);
                                    CELCant++;
                                }
                                else if (n[4].Equals("TOL"))
                                {
                                    TOLDur += Convert.ToInt32(n[5]);
                                    TOLTot += Convert.ToInt32(n[9]);
                                    TOLCant++;
                                }
                                else if (n[4].Equals("DDI"))
                                {
                                    DDIDur += Convert.ToInt32(n[5]);
                                    DDITot += Convert.ToInt32(n[9]);
                                    DDICant++;
                                }
                                else if (n[4].Equals("ENT"))
                                {
                                    ENTDur += Convert.ToInt32(n[5]);
                                    ENTTot += Convert.ToInt32(n[9]);
                                    ENTCant++;
                                }
                                else if (n[4].Equals("EXC"))
                                {
                                    EXCDur += Convert.ToInt32(n[5]);
                                    EXCTot += Convert.ToInt32(n[9]);
                                    EXCCant++;
                                }
                                else if (n[4].Equals("INT"))
                                {
                                    INTDur += Convert.ToInt32(n[5]);
                                    INTTot += Convert.ToInt32(n[9]);
                                    INTCant++;
                                }
                                else if (n[4].Equals("INV"))
                                {
                                    INVDur += Convert.ToInt32(n[5]);
                                    INVTot += Convert.ToInt32(n[9]);
                                    INVCant++;
                                }
                                else if (n[4].Equals("ITH"))
                                {
                                    ITHDur += Convert.ToInt32(n[5]);
                                    ITHTot += Convert.ToInt32(n[9]);
                                    ITHCant++;
                                }
                                else if (n[4].Equals("SAT"))
                                {
                                    SATDur += Convert.ToInt32(n[5]);
                                    SATTot += Convert.ToInt32(n[9]);
                                    SATCant++;
                                }

                                Repetido = false;
                                if (EXT2.Count != 0)
                                {
                                    foreach (List<string[]> g in EXT2)
                                    {
                                        if (g[0][10].Equals(n[10]))
                                        {
                                            Repetido = true;
                                        }
                                    }
                                }
                                if (Repetido == false)
                                {
                                    ext = new List<string[]>();
                                    foreach (string[] l in s)
                                    {
                                        if (n[10].Equals(l[10]))
                                        {
                                            ext.Add(l);
                                        }
                                    }
                                    EXT2.Add(ext);
                                }

                            }

                            DurRes = 0;
                            VrNetoRes = 0;
                            VrRecargoRes = 0;
                            VrIVARes = 0;
                            VrTotalRes = 0;
                            CantRes = 0;

                            foreach (List<string[]> n in EXT2)
                            {
                                Application.DoEvents();
                                DurGen = 0;
                                VrNetoGen = 0;
                                VrRecargoGen = 0;
                                VrIvaGen = 0;
                                VrTotalGen = 0;

                                foreach (string[] r in n)
                                {
                                    Application.DoEvents();
                                    DurGen += Convert.ToInt32(r[5]);
                                    VrNetoGen += Convert.ToInt32(r[6]);
                                    VrRecargoGen += Convert.ToInt32(r[7]);
                                    VrIvaGen += Convert.ToInt32(r[8]);
                                    VrTotalGen += Convert.ToInt32(r[9]);
                                }
                                try
                                {
                                    using (ConexionProg = new MySqlConnection(conexion))
                                    {
                                        ConexionProg.Open();
                                        query = "select * from extensiones where Nume_Extension = ?e";
                                        ComandoProg = new MySqlCommand(query, ConexionProg);
                                        ComandoProg.Parameters.AddWithValue("?e", (n[0][10]));
                                        LeeProg = ComandoProg.ExecuteReader();
                                        LeeProg.Read();
                                        RowTotal = new string[7] { "EXT: " + n[0][10] + " " + LeeProg["Nomb_Extension"].ToString(), n.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                                        LeeProg.Close();
                                        ConexionProg.Close();
                                    }
                                }
                                catch
                                {
                                    RowTotal = new string[7] { "EXT: " + n[0][10] + "Desconocida", n.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                                }
                                DataLlamadas.Rows.Add(RowTotal);
                                DurRes += DurGen;
                                VrNetoRes += VrNetoGen;
                                VrRecargoRes += VrRecargoGen;
                                VrIVARes += VrIvaGen;
                                VrTotalRes += VrTotalGen;
                                CantRes += n.Count;
                            }
                            RowTotal = new string[7] { "TOTAL", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                            DataLlamadas.Rows.Add(RowTotal);

                            pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                            AnchoPDF = new float[7] { 39f, 16f, 10f, 10f, 10f, 10f, 10f };
                            pdfTable.SetWidths(AnchoPDF);
                            pdfTable.WidthPercentage = 100;
                            pdfTable.SetWidths(AnchoPDF);
                            foreach (DataGridViewColumn column in DataLlamadas.Columns)
                            {
                                cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                                pdfTable.AddCell(cell);
                            }
                            foreach (DataGridViewRow row in DataLlamadas.Rows)
                            {
                                AnchoPDFpos = 0;
                                foreach (DataGridViewCell celda in row.Cells)
                                {
                                    if (celda.Value == null)
                                    {
                                        celda.Value = " ";
                                    }
                                    cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                                    if (AnchoPDFpos == 0)
                                    {
                                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                    }
                                    else
                                    {
                                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                    }
                                    pdfTable.AddCell(cell);
                                    AnchoPDFpos++;
                                }
                            }
                            pdfDoc.Add(pdfTable);
                            pdfDoc.Add(new Paragraph("\n\n"));
                        }
                    }
                }
                else
                {
                    DataLlamadas = new DataGridView();
                    DataLlamadas.ColumnCount = 7;

                    DataLlamadas.Columns[0].HeaderText = ReproAux[18].ToLower();
                    DataLlamadas.Columns[0].Width = 242;
                    DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                    DataLlamadas.Columns[1].Width = 70;
                    DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                    DataLlamadas.Columns[2].Width = 73;
                    DataLlamadas.Columns[3].HeaderText = "Vr.Neto";
                    DataLlamadas.Columns[3].Width = 77;
                    DataLlamadas.Columns[4].HeaderText = "Vr.Recargo";
                    DataLlamadas.Columns[4].Width = 77;
                    DataLlamadas.Columns[5].HeaderText = "Vr.IVA";
                    DataLlamadas.Columns[5].Width = 87;
                    DataLlamadas.Columns[6].HeaderText = "Vr.Total";
                    DataLlamadas.Columns[6].Width = 87;

                    CantRes = 0;
                    DurRes = 0;
                    VrNetoRes = 0;
                    VrRecargoRes = 0;
                    VrIVARes = 0;
                    VrTotalRes = 0;

                    foreach (List<string[]> s in EXT)
                    {
                        Application.DoEvents();
                        LabRes = "";
                        if (radioButton1.Checked)
                        {
                            LabRes = "EXT: " + (s[0][10]) + "    ";
                            try
                            {
                                using (ConexionProg = new MySqlConnection(conexion))
                                {
                                    ConexionProg.Open();
                                    query = "select * from extensiones where Nume_Extension = ?e";
                                    ComandoProg = new MySqlCommand(query, ConexionProg);
                                    ComandoProg.Parameters.AddWithValue("?e", (s[0][10]));
                                    LeeProg = ComandoProg.ExecuteReader();
                                    LeeProg.Read();
                                    LabRes += LeeProg["Nomb_Extension"].ToString();
                                    LeeProg.Close();
                                    ConexionProg.Close();
                                }
                            }
                            catch
                            {
                                LabRes += "Extension desconocida";
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            LabRes = "CENTRO: " + (s[0][12]) + "    ";
                            try
                            {
                                using (ConexionProg = new MySqlConnection(conexion))
                                {
                                    ConexionProg.Open();
                                    query = "select * from centros_costo where Codi_Centro = ?e";
                                    ComandoProg = new MySqlCommand(query, ConexionProg);
                                    ComandoProg.Parameters.AddWithValue("?e", (s[0][12]));
                                    LeeProg = ComandoProg.ExecuteReader();
                                    LeeProg.Read();
                                    LabRes += LeeProg["Nomb_Centro"].ToString();
                                    LeeProg.Close();
                                    ConexionProg.Close();
                                }
                            }
                            catch
                            {
                                LabRes += "Centro de costo desconocida";
                            }
                        }

                        else if (ReproAux[18].Equals("Troncales")) { LabRes = "TRONCAL: " + (s[0][13]); }
                        else if (ReproAux[18].Equals("Códigos personales")) { LabRes = "CÓDIGO: " + (s[0][14]); }
                        else if (ReproAux[18].Equals("Número de folio")) { LabRes = "FOLIO: " + (s[0][15]); }

                        DurGen = 0;
                        VrNetoGen = 0;
                        VrRecargoGen = 0;
                        VrIvaGen = 0;
                        VrTotalGen = 0;

                        foreach (string[] n in s)
                        {
                            Application.DoEvents();
                            if (n[4].Equals("LOC"))
                            {
                                LOCDur += Convert.ToInt32(n[5]);
                                LOCTot += Convert.ToInt32(n[9]);
                                LOCCant++;
                            }
                            else if (n[4].Equals("DDN"))
                            {
                                DDNDur += Convert.ToInt32(n[5]);
                                DDNTot += Convert.ToInt32(n[9]);
                                DDNCant++;
                            }
                            else if (n[4].Equals("CEL"))
                            {
                                CELDur += Convert.ToInt32(n[5]);
                                CELTot += Convert.ToInt32(n[9]);
                                CELCant++;
                            }
                            else if (n[4].Equals("TOL"))
                            {
                                TOLDur += Convert.ToInt32(n[5]);
                                TOLTot += Convert.ToInt32(n[9]);
                                TOLCant++;
                            }
                            else if (n[4].Equals("DDI"))
                            {
                                DDIDur += Convert.ToInt32(n[5]);
                                DDITot += Convert.ToInt32(n[9]);
                                DDICant++;
                            }
                            else if (n[4].Equals("ENT"))
                            {
                                ENTDur += Convert.ToInt32(n[5]);
                                ENTTot += Convert.ToInt32(n[9]);
                                ENTCant++;
                            }
                            else if (n[4].Equals("EXC"))
                            {
                                EXCDur += Convert.ToInt32(n[5]);
                                EXCTot += Convert.ToInt32(n[9]);
                                EXCCant++;
                            }
                            else if (n[4].Equals("INT"))
                            {
                                INTDur += Convert.ToInt32(n[5]);
                                INTTot += Convert.ToInt32(n[9]);
                                INTCant++;
                            }
                            else if (n[4].Equals("INV"))
                            {
                                INVDur += Convert.ToInt32(n[5]);
                                INVTot += Convert.ToInt32(n[9]);
                                INVCant++;
                            }
                            else if (n[4].Equals("ITH"))
                            {
                                ITHDur += Convert.ToInt32(n[5]);
                                ITHTot += Convert.ToInt32(n[9]);
                                ITHCant++;
                            }
                            else if (n[4].Equals("SAT"))
                            {
                                SATDur += Convert.ToInt32(n[5]);
                                SATTot += Convert.ToInt32(n[9]);
                                SATCant++;
                            }

                            DurGen += Convert.ToInt32(n[5]);
                            VrNetoGen += Convert.ToInt32(n[6]);
                            VrRecargoGen += Convert.ToInt32(n[7]);
                            VrIvaGen += Convert.ToInt32(n[8]);
                            VrTotalGen += Convert.ToInt32(n[9]);

                        }
                        RowTotal = new string[7] { LabRes, s.Count.ToString(), DurGen.ToString(), VrNetoGen.ToString(), VrRecargoGen.ToString(), VrIvaGen.ToString(), VrTotalGen.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                        DurRes += DurGen;
                        VrNetoRes += VrNetoGen;
                        VrRecargoRes += VrRecargoGen;
                        VrIVARes += VrIvaGen;
                        VrTotalRes += VrTotalGen;
                        CantRes += s.Count;
                    }
                    RowTotal = new string[7] { "TOTAL", CantRes.ToString(), DurRes.ToString(), VrNetoRes.ToString(), VrRecargoRes.ToString(), VrIVARes.ToString(), VrTotalRes.ToString() };
                    DataLlamadas.Rows.Add(RowTotal);

                    pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    AnchoPDF = new float[7] { 39f, 16f, 10f, 10f, 10f, 10f, 10f };
                    pdfTable.SetWidths(AnchoPDF);
                    pdfTable.WidthPercentage = 100;
                    pdfTable.SetWidths(AnchoPDF);
                    foreach (DataGridViewColumn column in DataLlamadas.Columns)
                    {
                        cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                        pdfTable.AddCell(cell);
                    }
                    foreach (DataGridViewRow row in DataLlamadas.Rows)
                    {
                        AnchoPDFpos = 0;
                        foreach (DataGridViewCell celda in row.Cells)
                        {
                            if (celda.Value == null)
                            {
                                celda.Value = " ";
                            }
                            cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                            if (AnchoPDFpos == 0)
                            {
                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                            }
                            else
                            {
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                            }
                            pdfTable.AddCell(cell);
                            AnchoPDFpos++;
                        }
                    }
                    pdfDoc.Add(pdfTable);
                    pdfDoc.Add(new Paragraph("\n\n"));
                }

                LabProg = "";
                LabProg = "TOTAL:";
                pdfDoc.Add(new Paragraph(LabProg, Fuente));

                DataLlamadas = new DataGridView();
                DataLlamadas.ColumnCount = 4;

                DataLlamadas.Columns[0].HeaderText = "Cl.Llam";
                DataLlamadas.Columns[1].HeaderText = "CANTIDAD";
                DataLlamadas.Columns[2].HeaderText = "DURACIÓN";
                DataLlamadas.Columns[3].HeaderText = "Vr.Total";

                foreach (string n in CheckCL2Prog)
                {
                    if (n.Equals("LOC"))
                    {
                        RowTotal = new string[4] { n, LOCCant.ToString(), LOCDur.ToString(), LOCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDN"))
                    {
                        RowTotal = new string[4] { n, DDNCant.ToString(), DDNDur.ToString(), DDNTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("CEL"))
                    {
                        RowTotal = new string[4] { n, CELCant.ToString(), CELDur.ToString(), CELTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("TOL"))
                    {
                        RowTotal = new string[4] { n, TOLCant.ToString(), TOLDur.ToString(), TOLTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("DDI"))
                    {
                        RowTotal = new string[4] { n, DDICant.ToString(), DDIDur.ToString(), DDITot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }

                    else if (n.Equals("ENT"))
                    {
                        RowTotal = new string[4] { n, ENTCant.ToString(), ENTDur.ToString(), ENTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("EXC"))
                    {
                        RowTotal = new string[4] { n, EXCCant.ToString(), EXCDur.ToString(), EXCTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INT"))
                    {
                        RowTotal = new string[4] { n, INTCant.ToString(), INTDur.ToString(), INTTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("INV"))
                    {
                        RowTotal = new string[4] { n, INVCant.ToString(), INVDur.ToString(), INVTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("ITH"))
                    {
                        RowTotal = new string[4] { n, ITHCant.ToString(), ITHDur.ToString(), ITHTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                    else if (n.Equals("SAT"))
                    {
                        RowTotal = new string[4] { n, SATCant.ToString(), SATDur.ToString(), SATTot.ToString() };
                        DataLlamadas.Rows.Add(RowTotal);
                    }
                }
                TotalValores = LOCTot + DDNTot + CELTot + TOLTot + DDITot + ENTTot + EXCTot + INTTot + INVTot + ITHTot + SATTot;
                TotalDuracion = LOCDur + DDNDur + CELDur + TOLDur + DDIDur + ENTDur + EXCDur + INTDur + INVDur + ITHDur + SATDur;
                TotalCantidad = LOCCant + DDNCant + CELCant + TOLCant + DDICant + ENTCant + EXCCant + INTCant + INVCant + ITHCant + SATCant;
                RowTotal = new string[4] { "TOTAL:", TotalCantidad.ToString(), TotalDuracion.ToString(), TotalValores.ToString() };
                DataLlamadas.Rows.Add(RowTotal);
                pdfTable = new PdfPTable(DataLlamadas.ColumnCount);
                pdfTable.DefaultCell.PaddingBottom = 3;
                pdfTable.DefaultCell.PaddingTop = 3;
                pdfTable.WidthPercentage = 30;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
                foreach (DataGridViewColumn column in DataLlamadas.Columns)
                {
                    cell = new PdfPCell(new Phrase(column.HeaderText, Fuente));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.BackgroundColor = new iTextSharp.text.Color(255, 217, 102);
                    pdfTable.AddCell(cell);
                }
                foreach (DataGridViewRow row in DataLlamadas.Rows)
                {
                    foreach (DataGridViewCell celda in row.Cells)
                    {
                        if (celda.Value == null)
                        {
                            celda.Value = " ";
                        }
                        cell = new PdfPCell(new Phrase(celda.Value.ToString(), Fuente));
                        pdfTable.AddCell(cell);
                    }
                }
                pdfDoc.Add(pdfTable);
                pdfDoc.Add(new Paragraph("\n\n"));

                pdfDoc.Close();
                stream.Close();

                Correo = "";
                Contraseña = "";
                using (ConexionProg = new MySqlConnection(ExpoDatos.conexion))
                {
                    ConexionProg.Open();
                    query = "select * from parametros where parametro = 'Correo envio ExpoDatos'";
                    ComandoProg = new MySqlCommand(query, ConexionProg);
                    LeeProg = ComandoProg.ExecuteReader();
                    LeeProg.Read();
                    Correo = LeeProg["seleccion"].ToString().Split(',')[0];
                    Contraseña = LeeProg["seleccion"].ToString().Split(',')[1];
                    ConexionProg.Close();
                }


                client = new SmtpClient("", 0);
                credentials = new NetworkCredential("", "");
                client = new SmtpClient("smtp.gmail.com", 587);
                client.EnableSsl = true;
                credentials = new NetworkCredential(Correo, Contraseña);
                client.Credentials = credentials;
                client.Timeout = 50000;
                MensajeProg = new MailMessage();

                if (!ReproAux[6].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[6])); }
                if (!ReproAux[7].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[7])); }
                if (!ReproAux[8].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[8])); }
                if (!ReproAux[9].Equals("-")) { MensajeProg.To.Add(new MailAddress(ReproAux[9])); }

                MensajeProg.Subject = ReproAux[0].Substring(0, ReproAux[0].Length-10);
                MensajeProg.From = new MailAddress(Correo);

                MensajeProg.Attachments.Add(new Attachment(Path.GetFileName(np)));
                try
                {
                    client.Send(MensajeProg);
                    MensajeProg.Dispose();
                    client.Dispose();
                }
                catch
                {
                    MensajeProg.Attachments.Clear();
                    MensajeProg.Dispose();
                    client.Dispose();
                    File.Delete(np);
                }

                File.Delete(np);
                if (ReproAux[1].Equals("diario2"))
                {
                    ReproGuarda[i][2] = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
                }
                else
                {
                    ReproGuarda[i][2] = DateTime.Now.ToString("yyyy/MM/dd");
                }
                File.Delete("repro.txt");
                using (StreamWriter escritor = new StreamWriter("repro.txt"))
                {
                    foreach (List<string> c in ReproGuarda)
                    {
                        foreach (string n in c)
                        {
                            escritor.WriteLine(n);
                        }
                        escritor.WriteLine("--");
                    }
                    escritor.Close();
                }
            }
        }

        #endregion
        
        #endregion

    }
}
