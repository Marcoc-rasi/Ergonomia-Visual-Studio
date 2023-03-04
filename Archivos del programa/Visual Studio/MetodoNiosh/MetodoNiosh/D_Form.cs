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
    public partial class D_Form : Form
    {
        public D_Form()
        {
            InitializeComponent();
        }

        private static D_Form instancia = null;
        public static D_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new D_Form();
                return instancia;
            }
            return instancia;
        }
        private void D_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
