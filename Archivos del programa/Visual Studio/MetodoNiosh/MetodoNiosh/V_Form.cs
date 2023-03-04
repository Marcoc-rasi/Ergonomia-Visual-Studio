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
    public partial class V_Form : Form
    {
        public V_Form()
        {
            InitializeComponent();
        }

        private static V_Form instancia = null;
        public static V_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new V_Form();
                return instancia;
            }
            return instancia;
        }
        private void V_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
