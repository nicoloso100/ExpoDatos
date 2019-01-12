using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace ExpoDatos
{
    public partial class Envio_correo : Form
    {
        MySqlConnection Conexion;
        MySqlDataReader lee;
        MySqlCommand comando;
        string query;
        string PDF;

        public Envio_correo(string Tabla, string pdf)
        {
            InitializeComponent();
            PDF = pdf;
            radioButton1.Checked = true;
            label126.Text = "Envío a " + Tabla;
            CargaList(Tabla);
        }
        List<string> Correos;
        bool Repetido;
        public void CargaList(string Tabla)
        {
            Correos = new List<string>();
            if (Tabla.Equals("Extensiones") || Tabla.Equals("Número de folio"))
            {
                try
                {
                    using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                    {
                        Conexion.Open();
                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            if(Correos.Count == 0)
                            {
                                listBox1.Items.Add(lee["Corr_Extension"].ToString());
                                Correos.Add(lee["Corr_Extension"].ToString());
                            }
                            else
                            {
                                Repetido = false;
                                foreach (string s in Correos)
                                {
                                    if (lee["Corr_Extension"].ToString().Equals(s))
                                    {
                                        Repetido = true;
                                    }
                                }
                                if(Repetido == false)
                                {
                                    listBox1.Items.Add(lee["Corr_Extension"].ToString());
                                    Correos.Add(lee["Corr_Extension"].ToString());
                                }
                            }
                        }
                        Conexion.Close();
                    }
                }
                catch
                {
                    MessageBox.Show("Ha ocurrio un error al cargar los correos de la tabla " + Tabla);
                }
            }
            else if(Tabla.Equals("Centros de costo"))
            {
                try
                {
                    using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                    {
                        Conexion.Open();
                        query = "select * from centros_costo";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            if (Correos.Count == 0)
                            {
                                listBox1.Items.Add(lee["Corr_Centro"].ToString());
                                Correos.Add(lee["Corr_Centro"].ToString());
                            }
                            else
                            {
                                Repetido = false;
                                foreach (string s in Correos)
                                {
                                    if (lee["Corr_Centro"].ToString().Equals(s))
                                    {
                                        Repetido = true;
                                    }
                                }
                                if (Repetido == false)
                                {
                                    listBox1.Items.Add(lee["Corr_Centro"].ToString());
                                    Correos.Add(lee["Corr_Centro"].ToString());
                                }
                            }
                        }
                        Conexion.Close();
                    }
                }
                catch
                {
                    MessageBox.Show("Ha ocurrio un error al cargar los correos de la tabla " + Tabla);
                }
            }
            else if (Tabla.Equals("Troncales"))
            {
                listBox1.Items.Add("No se ha encontrado correos en la tabla: " + Tabla);
                button1.Enabled = false;
            }
            else if(Tabla.Equals("Códigos personales"))
            {
                try
                {
                    using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                    {
                        Conexion.Open();
                        query = "select * from codigos_personales";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            if (Correos.Count == 0)
                            {
                                listBox1.Items.Add(lee["Corr_Codper"].ToString());
                                Correos.Add(lee["Corr_Codper"].ToString());
                            }
                            else
                            {
                                Repetido = false;
                                foreach (string s in Correos)
                                {
                                    if (lee["Corr_Codper"].ToString().Equals(s))
                                    {
                                        Repetido = true;
                                    }
                                }
                                if (Repetido == false)
                                {
                                    listBox1.Items.Add(lee["Corr_Codper"].ToString());
                                    Correos.Add(lee["Corr_Codper"].ToString());
                                }
                            }
                        }
                        Conexion.Close();
                    }
                }
                catch
                {
                    MessageBox.Show("Ha ocurrio un error al cargar los correos de la tabla " + Tabla);
                }
            }
            else if (Tabla.Equals("Generales"))
            {
                try
                {
                    using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                    {
                        Conexion.Open();
                        query = "select * from extensiones";
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            if (Correos.Count == 0)
                            {
                                listBox1.Items.Add(lee["Corr_Extension"].ToString());
                                Correos.Add(lee["Corr_Extension"].ToString());
                            }
                            else
                            {
                                Repetido = false;
                                foreach (string s in Correos)
                                {
                                    if (lee["Corr_Extension"].ToString().Equals(s))
                                    {
                                        Repetido = true;
                                    }
                                }
                                if (Repetido == false)
                                {
                                    listBox1.Items.Add(lee["Corr_Extension"].ToString());
                                    Correos.Add(lee["Corr_Extension"].ToString());
                                }
                            }
                        }
                        Conexion.Close();
                    }
                }
                catch
                {
                    MessageBox.Show("Ha ocurrio un error al cargar los correos de la tabla extensiones");
                }
            }
            else
            {
                listBox1.Items.Add("No se ha encontrado correos en la tabla: " + Tabla);
                button1.Enabled = false;
            }
        }
        string Correo;
        string Contraseña;
        public void Envia()
        {
            try
            {
                Correo = "";
                Contraseña = "";
                using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                {
                    Conexion.Open();
                    query = "select * from parametros where parametro = 'Correo envio ExpoDatos'";
                    comando = new MySqlCommand(query, Conexion);
                    lee = comando.ExecuteReader();
                    lee.Read();
                    Correo = lee["seleccion"].ToString().Split(',')[0];
                    Contraseña = lee["seleccion"].ToString().Split(',')[1];
                    Conexion.Close();
                }
                try
                {
                    SmtpClient client = new SmtpClient("", 0);
                    NetworkCredential credentials = new NetworkCredential("", "");
                    client = new SmtpClient("smtp.gmail.com", 587);
                    client.EnableSsl = true;
                    credentials = new NetworkCredential(Correo, Contraseña);
                    client.Credentials = credentials;
                    client.Timeout = 10000;
                    MailMessage Mensaje = new MailMessage();
                    if (radioButton1.Checked)
                    {
                        foreach (string s in listBox1.Items)
                        {
                            Mensaje.To.Add(new MailAddress(s));
                        }
                    }
                    else if (radioButton2.Checked)
                    {
                        if (!string.IsNullOrEmpty(textBox1.Text)) { Mensaje.To.Add(new MailAddress(textBox1.Text)); }
                        if (!string.IsNullOrEmpty(textBox2.Text)) { Mensaje.To.Add(new MailAddress(textBox2.Text)); }
                        if (!string.IsNullOrEmpty(textBox3.Text)) { Mensaje.To.Add(new MailAddress(textBox3.Text)); }
                        if (!string.IsNullOrEmpty(textBox4.Text)) { Mensaje.To.Add(new MailAddress(textBox4.Text)); }
                        if (!string.IsNullOrEmpty(textBox5.Text)) { Mensaje.To.Add(new MailAddress(textBox5.Text)); }
                    }
                    Mensaje.Subject = "Reporte " + DateTime.Now.ToString();
                    Mensaje.From = new MailAddress(Correo);
                    try
                    {
                        Mensaje.Attachments.Add(new Attachment(Path.GetFileName(PDF)));
                        client.Send(Mensaje);
                        MessageBox.Show("El reporte se ha enviado correctamente");
                        Mensaje.Dispose();
                        client.Dispose();

                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ha ocurrido un error al enviar el reporte\n\n" + ex.ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ha ocurrido un error al configurar el envío del correo\n\n" + ex.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al consultar el correo de envío de reportes en la base de datos\n\n" + ex.ToString());
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                groupBox1.Enabled = true;
            }
            else
            {
                groupBox1.Enabled = false;
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                groupBox2.Enabled = true;
            }
            else
            {
                groupBox2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Envia();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Envia();
        }
    }
}
