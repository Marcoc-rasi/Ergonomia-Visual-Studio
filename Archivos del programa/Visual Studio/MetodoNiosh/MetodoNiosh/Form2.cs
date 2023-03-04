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
    public partial class TablaFrecuenciaForm : Form
    {
        public TablaFrecuenciaForm()
        {
            InitializeComponent();
        }
        private static TablaFrecuenciaForm instancia = null;
        public static TablaFrecuenciaForm ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new TablaFrecuenciaForm();
                return instancia;
            }
            return instancia;
        }

        private void MandarValor_Click(object sender, EventArgs e)
        {
            if (txtBValorFMOrigen.Text.Length == 0)
            {
                MessageBox.Show("Fm Origen vacio Introduce un valor para continuar");
            }
            else
            {
                if (float.TryParse(txtBValorFMOrigen.Text, out ValoresTablas.FMOrigen) == true)
                {
                    MessageBox.Show("El valor de FM Origen se guardo Correctamente refresca los datos para verlo");
                }
                else
                {
                    MessageBox.Show("Introduce un numero correcto en FM Origen");
                }
            }

            if (txtBValorFMDestino.Text.Length == 0)
            {
                MessageBox.Show("Fm Destino vacio Introduce un valor para continuar");
            }
            else
            {
                if (float.TryParse(txtBValorFMDestino.Text, out ValoresTablas.FMDestino) == true)
                {
                    MessageBox.Show("El valor de FM Destino se guardo Correctamente refresca los datos para verlo");
                }
                else
                {
                    MessageBox.Show("Introduce un numero correcto en FM Destino");
                }
            }
        }

        private void TablaFrecuenciaForm_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
