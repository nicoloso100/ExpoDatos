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
    public partial class Mensaje : Form
    {
        public static bool Seleccionado = false;
        public static string Seleccion = "";
        public Mensaje()
        {
            InitializeComponent();
            this.FormClosing += Mensaje_FormClosing;
        }

        private void Mensaje_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(Seleccionado == false)
            {
                MessageBox.Show("Escoja una opción!");
                e.Cancel = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Seleccion = "0";
            Seleccionado = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Seleccion = "1";
            Seleccionado = true;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Seleccion = "2";
            Seleccionado = true;
            this.Close();
        }
    }
}
