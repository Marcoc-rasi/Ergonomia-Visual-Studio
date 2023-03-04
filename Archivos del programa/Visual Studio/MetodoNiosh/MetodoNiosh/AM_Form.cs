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
    public partial class AM_Form : Form
    {
        public AM_Form()
        {
            InitializeComponent();
        }

        private static AM_Form instancia = null;
        public static AM_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new AM_Form();
                return instancia;
            }
            return instancia;
        }
        private void AM_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
