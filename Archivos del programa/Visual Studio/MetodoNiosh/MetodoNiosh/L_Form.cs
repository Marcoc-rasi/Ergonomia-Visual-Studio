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
    public partial class L_Form : Form
    {
        public L_Form()
        {
            InitializeComponent();
        }

        private static L_Form instancia = null;
        public static L_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new L_Form();
                return instancia;
            }
            return instancia;
        }
        private void L_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
