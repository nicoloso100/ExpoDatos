using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExpoDatos
{
    public partial class Login : Form
    {

        public static bool Logeado = false;
        public static string usuario = "";
        MySqlConnection Conexion;
        MySqlCommand comando;
        MySqlDataReader lee;
        public static List<string> Permisos = new List<string>();

        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (Conexion = new MySqlConnection(ExpoDatos.conexion))
                {
                    Conexion.Open();

                    string query = "select * from usuarios";
                    comando = new MySqlCommand(query, Conexion);
                    lee = comando.ExecuteReader();
                    bool Correcto = false;
                    string usr = "";
                    while (lee.Read())
                    {
                        if (textBox1.Text.Equals(lee["id"].ToString()))
                        {
                            if (textBox2.Text.Equals(lee["ingreso"].ToString()))
                            {
                                Correcto = true;
                                usr = lee["perfil"].ToString();
                                usuario = lee["NombreUsuario"].ToString();
                            }
                        }
                    }
                    lee.Close();
                    if (Correcto == true)
                    {
                        query = "select * from perfiles where perfil = ?c";
                        comando = new MySqlCommand(query, Conexion);
                        comando.Parameters.AddWithValue("?c", usr);
                        lee = comando.ExecuteReader();
                        lee.Read();
                        usr = lee["DetaPerfil"].ToString();
                        lee.Close();

                        query = "select * from " + "perfil" + usr.ToLower().Split(' ')[0];
                        comando = new MySqlCommand(query, Conexion);
                        lee = comando.ExecuteReader();
                        while (lee.Read())
                        {
                            Permisos.Add(lee["codigo"].ToString() + lee["SiNo"].ToString());
                        }
                        Logeado = true;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Usuario o contraseña incorrectos!!");
                    }
                    Conexion.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ocurrió un error al leer la base de datos");
            }
        }
    }
}
