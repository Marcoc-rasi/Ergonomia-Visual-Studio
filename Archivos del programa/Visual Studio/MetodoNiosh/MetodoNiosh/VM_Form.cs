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
    public partial class VM_Form : Form
    {
        public VM_Form()
        {
            InitializeComponent();
        }

        private static VM_Form instancia = null;
        public static VM_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new VM_Form();
                return instancia;
            }
            return instancia;
        }
        private void VM_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
