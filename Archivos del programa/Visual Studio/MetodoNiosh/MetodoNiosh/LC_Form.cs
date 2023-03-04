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
    public partial class LC_Form : Form
    {
        public LC_Form()
        {
            InitializeComponent();
        }

        private static LC_Form instancia = null;
        public static LC_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new LC_Form();
                return instancia;
            }
            return instancia;
        }
        private void LC_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
