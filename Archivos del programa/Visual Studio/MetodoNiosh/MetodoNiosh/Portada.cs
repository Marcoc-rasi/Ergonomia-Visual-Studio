using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MetodoNiosh
{
    public partial class Portada : Form
    {
        public Portada()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnSiguiente_Click(object sender, EventArgs e)
        {
            HojaDeTrabajoNiosh formularioMetodoNiosh = HojaDeTrabajoNiosh.ventanaUnica();
            formularioMetodoNiosh.Show();
            formularioMetodoNiosh.BringToFront();
        }

        public static Portada instancia = null;
        public static Portada ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new Portada();
                return instancia;
            }
            return instancia;
        }
    }
}
