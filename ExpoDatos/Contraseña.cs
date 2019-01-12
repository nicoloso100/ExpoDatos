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
    public partial class Contraseña : Form
    {

        public static bool Sale = false;

        public Contraseña()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals("RD0315"))
            {
                ExpoDatos.Salir = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("Contraseña incorrecta!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExpoDatos.Salir = false;
            this.Close();
        }
    }
}
