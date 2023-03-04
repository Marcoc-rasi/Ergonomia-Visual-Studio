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
    public partial class DM_Form : Form
    {
        public DM_Form()
        {
            InitializeComponent();
        }

        private static DM_Form instancia = null;
        public static DM_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new DM_Form();
                return instancia;
            }
            return instancia;
        }
        private void DM_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
