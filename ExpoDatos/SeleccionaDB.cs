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
    public partial class SeleccionaDB : Form
    {

        bool MensajeClose = true;

        public SeleccionaDB()
        {
            InitializeComponent();
            this.FormClosing += SeleccionaDB_FormClosing;
        }

        private void SeleccionaDB_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(MensajeClose == true)
            {
                MessageBox.Show("Operacion cancelada, se seleccionará la base de datos por defecto(" + ExpoDatos.Config[0].Split(';')[2].Split('=')[1] + ")");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                ExpoDatos.Config[0] = ExpoDatos.Config[0].Split(';')[0] + ";" + ExpoDatos.Config[0].Split(';')[1] + ";" + ExpoDatos.Config[0].Split(';')[2].Split('=')[0] + "=" + comboBox1.Text + ";" + ExpoDatos.Config[0].Split(';')[3] + ";" + ExpoDatos.Config[0].Split(';')[4];
                MensajeClose = false;
                this.Close();
            }
            else
            {
                MessageBox.Show("Seleccione una base de datos!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
