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
    public partial class TablaAcoplamiento : Form
    {
        private static TablaAcoplamiento instancia = null;
        public static TablaAcoplamiento ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new TablaAcoplamiento();
                return instancia;
            }
            return instancia;
        }
        private void TablaAcoplamiento_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
        private void GuardarValor_Click(object sender, EventArgs e)
        {
            if (txtBValorCMOrigen.Text.Length == 0)
            {
                MessageBox.Show("CM Origen vacio Introduce un valor para continuar");
            }
            else
            {
                if (float.TryParse(txtBValorCMOrigen.Text, out ValoresTablas.CMOrigen) == true)
                {
                    MessageBox.Show("El valor de CM Origen se guardo Correctamente refresca los datos para verlo");
                }
                else
                {
                    MessageBox.Show("Introduce un numero correcto en CM Origen");
                }
            }

            if (txtBValorCMDestino.Text.Length == 0)
            {
                MessageBox.Show("Cm Destino vacio Introduce un valor para continuar");
            }
            else
            {
                if (float.TryParse(txtBValorCMDestino.Text, out ValoresTablas.CMDestino) == true)
                {
                    MessageBox.Show("El valor de CM Destino se guardo Correctamente refresca los datos para verlo");
                }
                else
                {
                    MessageBox.Show("Introduce un numero correcto en CM Destino");
                }
            }
        }

        
        public TablaAcoplamiento()
        {
            InitializeComponent();
        }
    }
}
