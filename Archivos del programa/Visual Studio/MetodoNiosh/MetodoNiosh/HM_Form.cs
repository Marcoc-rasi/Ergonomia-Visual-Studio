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
    public partial class HM_Form : Form
    {
        public HM_Form()
        {
            InitializeComponent();
        }

        private static HM_Form instancia = null;
        public static HM_Form ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new HM_Form();
                return instancia;
            }
            return instancia;
        }
        private void HM_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }
    }
}
