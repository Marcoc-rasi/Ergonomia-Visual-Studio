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
    public partial class H_Form : Form
    {
        public H_Form()
        {
            InitializeComponent();
        }

        private static H_Form instancia = null;
        public static H_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new H_Form();
                return instancia;
            }
            return instancia;
        }

        private void H_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
