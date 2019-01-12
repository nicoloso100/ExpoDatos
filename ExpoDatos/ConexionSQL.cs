using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExpoDatos
{
    public partial class ConexionSQL : Form
    {
        //salir
        public static bool Exit = true;

        //DIrectorio
        public static string Direccion = Directory.GetCurrentDirectory();
        public static string Carpeta = @Direccion;
        public static string SubCarpeta = System.IO.Path.Combine(Carpeta, "Configuracion");
        public static string Archivo = System.IO.Path.Combine(SubCarpeta, "Configuracion.txt");

        //SQL
        
        string conexion;

        public ConexionSQL()
        {
            InitializeComponent();
            textBox6.Text = Carpeta;
            this.FormClosed += ConexionSQL_FormClosed;
        }

        private void ConexionSQL_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Exit == true)
            {
                Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
                if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text)
                    && !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrEmpty(textBox6.Text))
                {
                    ExpoDatos.conexion = "server=" + textBox1.Text + ";port=" + textBox2.Text + ";database=" + textBox5.Text + ";Uid=" + textBox3.Text + ";pwd=" + textBox4.Text;
                    conexion = ExpoDatos.conexion;
                    
                    if(PruebaConexion() == true)
                    {
                        if (!Directory.Exists(SubCarpeta))
                        {
                            Directory.CreateDirectory(SubCarpeta);
                        }
                        using (StreamWriter escritor = new StreamWriter(Archivo))
                        {
                            escritor.WriteLine(ExpoDatos.conexion);
                            escritor.Close();
                        }
                        MessageBox.Show("La configuración se ha establecido exitosamente");
                        Exit = false;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ha ocurrido un error al conectarse con SQL!");
                    }
                    
                }
                else
                {
                    MessageBox.Show("Hay un campo vacío!");
                }
            }
            catch(Exception n)
            {
                MessageBox.Show("Ha ocurrido un error al escribir en el archivo Configuracion.txt\n\n" + n.ToString());
            }
        }

        public bool PruebaConexion()
        {
            try
            {
                using (MySqlConnection Conexion = new MySqlConnection(conexion))
                {
                    Conexion.Open();
                    return (true);
                }
                
            }
            catch
            {
                return (false);
            }
            
        }
    }
}
