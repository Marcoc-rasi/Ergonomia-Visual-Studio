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
    public partial class A_Form : Form
    {
        public A_Form()
        {
            InitializeComponent();
        }

        private static A_Form instancia = null;
        public static A_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new A_Form();
                return instancia;
            }
            return instancia;
        }
        private void A_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
