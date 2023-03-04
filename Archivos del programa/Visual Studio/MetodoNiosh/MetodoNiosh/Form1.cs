using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetLight;
namespace MetodoNiosh
{
    public partial class HojaDeTrabajoNiosh : Form
    {
        public HojaDeTrabajoNiosh()
        {
            InitializeComponent();
        }

        public static HojaDeTrabajoNiosh instancia = null;
        public static HojaDeTrabajoNiosh ventanaUnica()
        {
            if (instancia == null)
            {
                instancia = new HojaDeTrabajoNiosh();
                return instancia;
            }
            return instancia;
        }

        private void FM_Click(object sender, EventArgs e)
        {
            TablaFrecuenciaForm formularioFrecuencia = TablaFrecuenciaForm.ventanaUnica();
            formularioFrecuencia.Show();
            formularioFrecuencia.BringToFront();
        }

        private void btnRefrescar_Click(object sender, EventArgs e)
        {
            txtFMOrigen.Text = ValoresTablas.FMOrigen.ToString();
            txtCMOrigen.Text = ValoresTablas.CMOrigen.ToString();
            txtFMDestino.Text = ValoresTablas.FMDestino.ToString();
            txtCMDestino.Text = ValoresTablas.CMDestino.ToString();
        }
        
        private void CM_Click(object sender, EventArgs e)
        {
            TablaAcoplamiento formularioFrecuencia = TablaAcoplamiento.ventanaUnica();
            formularioFrecuencia.Show();
            formularioFrecuencia.BringToFront();
        }
      
        private void chkBSistemaInternacional_Click(object sender, EventArgs e)
        {
            if (chkBSistemaIngles.Checked == true)
            {
                MessageBox.Show("Desmarca la casilla de Sistema Ingles para continuar");
                chkBSistemaInternacional.Checked = false;
            }
            else
            {
                txtBPesoObjetoSistema.Text = "Kg";
                txtBUbicacionManoSistema.Text = "cm";
                txtBDistanciaVerticalSistema.Text = "cm";
                comboBoxRWLOrigen.Text = "";
                comboBoxRWLDestino.Text = "";
                txtBOrigenRWL.Text = "";
                txtBDestinoRWL.Text = "";
                //txtBRWLOrigenSistema.Text = "Kg";
                //txtBRWLDestinoSistema.Text = "Kg";
            }
        }

        private void chkBSistemaIngles_Click(object sender, EventArgs e)
        {
            if (chkBSistemaInternacional.Checked == true)
            {
                MessageBox.Show("Desmarca la casilla de Sistema Internacional para continuar");
                chkBSistemaIngles.Checked = false;
            }
            else
            {
                txtBPesoObjetoSistema.Text = "Lbs";
                txtBUbicacionManoSistema.Text = "in";
                txtBDistanciaVerticalSistema.Text = "in";
                comboBoxRWLOrigen.Text = "";
                comboBoxRWLDestino.Text = "";
                txtBOrigenRWL.Text = "";
                txtBDestinoRWL.Text = "";
                //txtBRWLOrigenSistema.Text = "Lbs";
                //txtBRWLDestinoSistema.Text = "Lbs";
            }
        }

        private void chkBLProm_Click(object sender, EventArgs e)
        {
            if (chkBLMax.Checked == true)
            {
                MessageBox.Show("Desmarca la casilla de L (máx.) para continuar");
                chkBLProm.Checked = false;
            }
        }

        private void chkBLMax_Click(object sender, EventArgs e)
        {
            if (chkBLProm.Checked == true)
            {
                MessageBox.Show("Desmarca la casilla de L (Prom.) para continuar");
                chkBLMax.Checked = false;
            }
        }

        void LimparConversion()
        {
            ///Hoja de trabajo paso 1
            ///Datos
            txtBPesoObjetoLProm.Clear();
            txtBPesoObjetoLMax.Clear();
            txtBUbicacionManoOrigenH.Clear();
            txtBUbicacionManoOrigenV.Clear();
            txtBUbicacionManoDestinoH.Clear();
            txtBUbicacionManoDestinoV.Clear();
            txtBDistanciaVerticalD.Clear();
            ///

            ///Hoja de trabajo paso 2
            ///Datos
            txtBOrigenLC.Clear();
            txtBDestinoLC.Clear();
            txtBOrigenHM.Clear();
            txtBDestinoHM.Clear();
            txtBOrigenVM.Clear();
            txtBDestinoVM.Clear();
            txtBOrigenDM.Clear();
            txtBDestinoDM.Clear();
            txtBOrigenAM.Clear();
            txtBDestinoAM.Clear();
            txtBOrigenRWL.Clear();
            txtBDestinoRWL.Clear();
            ///

            ///Hoja de trabajo paso 3
            ///Datos
            txtBIndiceElevacionOrigenPesoObjetoL.Clear();
            txtBIndiceElevacionOrigenRWL.Clear();
            txtBIndiceElevacionDestinoPesoObjetoL.Clear();
            txtBIndiceElevacionDestinoRWL.Clear();
            txtBIndiceElevacionOrigen.Clear();
            txtBIndiceElevacionDestino.Clear();
            ///
        }
        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            ///Hoja de trabajo descripcion del caso
            ///Datos
            txtBDepartamento.Clear();
            txtBTituloProfesional.Clear();
            txtBNombreAnalista.Clear();
            txtBFecha.Clear();
            txtBDescripcionDelTrabajo.Clear();
            ///

            ///Hoja de trabajo paso 1
            ///Sistema
            txtBPesoObjetoSistema.Clear();
            txtBUbicacionManoSistema.Clear();
            txtBDistanciaVerticalSistema.Clear();
            ///Datos
            txtBPesoObjetoLProm.Clear();
            txtBPesoObjetoLMax.Clear();
            txtBUbicacionManoOrigenH.Clear();
            txtBUbicacionManoOrigenV.Clear();
            txtBUbicacionManoDestinoH.Clear();
            txtBUbicacionManoDestinoV.Clear();
            txtBDistanciaVerticalD.Clear();
            txtBAnguloAsimetricoOrigenA.Clear();
            txtBAnguloAsimetricoDestinoA.Clear();
            txtBFrecuencia.Clear();
            txtBDuracion.Clear();
            txtBAcoplamientoObjetos.Clear();
            ///

            ///Hoja de trabajo paso 2
            ///Sistema
            //txtBRWLOrigenSistema.Clear();
            comboBoxRWLOrigen.Text = "";
            comboBoxRWLDestino.Text = "";
            //txtBRWLDestinoSistema.Clear();
            ///Datos
            txtBOrigenLC.Clear();
            txtBDestinoLC.Clear();
            txtBOrigenHM.Clear();
            txtBDestinoHM.Clear();
            txtBOrigenVM.Clear();
            txtBDestinoVM.Clear();
            txtBOrigenDM.Clear();
            txtBDestinoDM.Clear();
            txtBOrigenAM.Clear();
            txtBDestinoAM.Clear();
            txtFMOrigen.Clear();
            txtFMDestino.Clear();
            txtCMOrigen.Clear();
            txtCMDestino.Clear();
            txtBOrigenRWL.Clear();
            txtBDestinoRWL.Clear();
            ///

            ///Hoja de trabajo paso 3
            ///Datos
            txtBIndiceElevacionOrigenPesoObjetoL.Clear();
            txtBIndiceElevacionOrigenRWL.Clear();
            txtBIndiceElevacionDestinoPesoObjetoL.Clear();
            txtBIndiceElevacionDestinoRWL.Clear();
            txtBIndiceElevacionOrigen.Clear();
            txtBIndiceElevacionDestino.Clear();
            ///

            ///Limpiar CheckBox
            chkBSistemaInternacional.Checked = false;
            chkBSistemaIngles.Checked = false;
            chkBLMax.Checked = false;
            chkBLProm.Checked = false;
            ///

        }

        private void btnCalcularResultado_Click(object sender, EventArgs e)
        {
            if(BuscarCasillaVacia() == 1 && BuscarCasillaConLetra() == 1)
            {
                if (chkBSistemaIngles.Checked == false && chkBSistemaInternacional.Checked == false)
                {
                    MessageBox.Show("Selecciona un Sistema metrico para continuar");
                }
                else
                {
                    if (chkBLMax.Checked == false && chkBLProm.Checked == false)
                    {

                        MessageBox.Show("Selecciona una L en indice de elevacion para continuar");
                    }
                    else 
                    {  
                        if (chkBSistemaIngles.Checked == true)
                        {
                            //LC
                            txtBOrigenLC.Text = "51";
                            ValoresTablas.LCOrigen = 51;
                            txtBDestinoLC.Text = "51";
                            ValoresTablas.LCDestino = 51;
                            //

                            //HM
                            ValoresTablas.HMOrigen = (10 / ValoresTablas.UbicacionManoOrigenH);
                            txtBOrigenHM.Text = ValoresTablas.HMOrigen.ToString();

                            ValoresTablas.HMDestino = (10 / ValoresTablas.UbicacionManoDestinoH);
                            txtBDestinoHM.Text = ValoresTablas.HMDestino.ToString();
                            //

                            //VM
                            ValoresTablas.VMOrigen = (float)(1 - (0.0075 * Math.Abs(ValoresTablas.UbicacionManoOrigenV - 30)));
                            txtBOrigenVM.Text = ValoresTablas.VMOrigen.ToString();

                            ValoresTablas.VMDestino = (float)(1 - (0.0075 * Math.Abs(ValoresTablas.UbicacionManoDestinoV - 30)));
                            txtBDestinoVM.Text = ValoresTablas.VMDestino.ToString();
                            //

                            //DM
                            ValoresTablas.DMOrigen = (float)(0.82 + (1.8 / ValoresTablas.DistanciaVerticalD));
                            txtBOrigenDM.Text = ValoresTablas.DMOrigen.ToString();

                            ValoresTablas.DMDestino = (float)(0.82 + (1.8 / ValoresTablas.DistanciaVerticalD));
                            txtBDestinoDM.Text = ValoresTablas.DMDestino.ToString();
                            //

                            //AM
                            ValoresTablas.AMOrigen = (float)(1 - (0.032 * ValoresTablas.AnguloAsimetricoOrigenA));
                            txtBOrigenAM.Text = ValoresTablas.AMOrigen.ToString();

                            ValoresTablas.AMDestino = (float)(1 - (0.032 * ValoresTablas.AnguloAsimetricoDestinoA));
                            txtBDestinoAM.Text = ValoresTablas.AMDestino.ToString();

                            //RWL 
                            ValoresTablas.OrigenRWL = ValoresTablas.LCOrigen * ValoresTablas.HMOrigen
                                * ValoresTablas.VMOrigen * ValoresTablas.DMOrigen * ValoresTablas.AMOrigen
                                * ValoresTablas.FMOrigen * ValoresTablas.CMOrigen;
                            //txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();

                            ValoresTablas.DestinoRWL = ValoresTablas.LCDestino * ValoresTablas.HMDestino
                                * ValoresTablas.VMDestino * ValoresTablas.DMDestino * ValoresTablas.AMDestino
                                * ValoresTablas.FMDestino * ValoresTablas.CMDestino;
                            //txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                            if (comboBoxRWLOrigen.Text.ToString() == "")
                            {
                                MessageBox.Show("Selecciona las unidades para calcular el RWL Origen\nen el paso 2");
                            }
                            if (comboBoxRWLDestino.Text.ToString() == "")
                            {
                                MessageBox.Show("Selecciona las unidades para calcular el RWL Destino\nen el paso 2");
                            }
                            if (comboBoxRWLOrigen.Text.ToString() == "Mili-Libras")
                            {
                                ValoresTablas.OrigenRWL = ValoresTablas.OrigenRWL * 1000;
                                txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                            }
                            if (comboBoxRWLOrigen.Text.ToString() == "Libras")
                            {
                                txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                            }
                            if (comboBoxRWLOrigen.Text.ToString() == "Kilo-Libras")
                            {
                                ValoresTablas.OrigenRWL = ValoresTablas.OrigenRWL / 1000;
                                txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                            }

                            if (comboBoxRWLDestino.Text.ToString() == "Mili-Libras")
                            {
                                ValoresTablas.DestinoRWL = ValoresTablas.DestinoRWL * 1000;
                                txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                            }
                            if (comboBoxRWLDestino.Text.ToString() == "Libras")
                            {
                                txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                            }
                            if (comboBoxRWLDestino.Text.ToString() == "Kilo-Libras")
                            {
                                ValoresTablas.DestinoRWL = ValoresTablas.DestinoRWL / 1000;
                                txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                            }
                            //

                            //Indice de elevacion
                            if (chkBLProm.Checked == true)
                            {
                                ValoresTablas.IndiceElevacionOrigen = ValoresTablas.PesoObjetoLProm / ValoresTablas.OrigenRWL;
                                txtBIndiceElevacionOrigenPesoObjetoL.Text = ValoresTablas.PesoObjetoLProm.ToString();
                                txtBIndiceElevacionOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                txtBIndiceElevacionOrigen.Text = ValoresTablas.IndiceElevacionOrigen.ToString();

                                ValoresTablas.IndiceElevacionDestino = ValoresTablas.PesoObjetoLProm / ValoresTablas.DestinoRWL;
                                txtBIndiceElevacionDestinoPesoObjetoL.Text = ValoresTablas.PesoObjetoLProm.ToString();
                                txtBIndiceElevacionDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                txtBIndiceElevacionDestino.Text = ValoresTablas.IndiceElevacionDestino.ToString();
                            }
                            if (chkBLMax.Checked == true)
                            {
                                ValoresTablas.IndiceElevacionOrigen = ValoresTablas.PesoObjetoLMax / ValoresTablas.OrigenRWL;
                                txtBIndiceElevacionOrigenPesoObjetoL.Text = ValoresTablas.PesoObjetoLMax.ToString();
                                txtBIndiceElevacionOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                txtBIndiceElevacionOrigen.Text = ValoresTablas.IndiceElevacionOrigen.ToString();

                                ValoresTablas.IndiceElevacionDestino = ValoresTablas.PesoObjetoLMax / ValoresTablas.DestinoRWL;
                                txtBIndiceElevacionDestinoPesoObjetoL.Text = ValoresTablas.PesoObjetoLMax.ToString();
                                txtBIndiceElevacionDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                txtBIndiceElevacionDestino.Text = ValoresTablas.IndiceElevacionDestino.ToString();
                            }
                            //
                        }
                        else
                        {
                            if (chkBSistemaInternacional.Checked == true)
                            {
                                //LC
                                txtBOrigenLC.Text = "23";
                                ValoresTablas.LCOrigen = 23;
                                txtBDestinoLC.Text = "23";
                                ValoresTablas.LCDestino = 23;
                                //

                                //HM
                                ValoresTablas.HMOrigen = (25 / ValoresTablas.UbicacionManoOrigenH);
                                txtBOrigenHM.Text = ValoresTablas.HMOrigen.ToString();

                                ValoresTablas.HMDestino = (25 / ValoresTablas.UbicacionManoDestinoH);
                                txtBDestinoHM.Text = ValoresTablas.HMDestino.ToString();
                                //

                                //VM
                                ValoresTablas.VMOrigen = (float)(1 - (0.003 * Math.Abs(ValoresTablas.UbicacionManoOrigenV - 75)));
                                txtBOrigenVM.Text = ValoresTablas.VMOrigen.ToString();

                                ValoresTablas.VMDestino = (float)(1 - (0.003 * Math.Abs(ValoresTablas.UbicacionManoDestinoV - 75)));
                                txtBDestinoVM.Text = ValoresTablas.VMDestino.ToString();
                                //

                                //DM
                                ValoresTablas.DMOrigen = (float)(0.82 + (4.5 / ValoresTablas.DistanciaVerticalD));
                                txtBOrigenDM.Text = ValoresTablas.DMOrigen.ToString();

                                ValoresTablas.DMDestino = (float)(0.82 + (4.5 / ValoresTablas.DistanciaVerticalD));
                                txtBDestinoDM.Text = ValoresTablas.DMDestino.ToString();
                                //

                                //AM
                                ValoresTablas.AMOrigen = (float)(1 - (0.032 * ValoresTablas.AnguloAsimetricoOrigenA));
                                txtBOrigenAM.Text = ValoresTablas.AMOrigen.ToString();

                                ValoresTablas.AMDestino = (float)(1 - (0.032 * ValoresTablas.AnguloAsimetricoDestinoA));
                                txtBDestinoAM.Text = ValoresTablas.AMDestino.ToString();

                                //RWL 
                                
                                ValoresTablas.OrigenRWL = ValoresTablas.LCOrigen * ValoresTablas.HMOrigen
                                    * ValoresTablas.VMOrigen * ValoresTablas.DMOrigen * ValoresTablas.AMOrigen
                                    * ValoresTablas.FMOrigen * ValoresTablas.CMOrigen;
                               // txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();

                                ValoresTablas.DestinoRWL = ValoresTablas.LCDestino * ValoresTablas.HMDestino
                                    * ValoresTablas.VMDestino * ValoresTablas.DMDestino * ValoresTablas.AMDestino
                                    * ValoresTablas.FMDestino * ValoresTablas.CMDestino;
                                // txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                if (comboBoxRWLOrigen.Text.ToString() == "")
                                {
                                    MessageBox.Show("Selecciona las unidades para calcular el RWL Origen\nen el paso 2");
                                }
                                if (comboBoxRWLDestino.Text.ToString() == "")
                                {
                                    MessageBox.Show("Selecciona las unidades para calcular el RWL Destino\nen el paso 2");
                                }
                                if (comboBoxRWLOrigen.Text.ToString() == "Gramos")
                                {
                                    ValoresTablas.OrigenRWL = ValoresTablas.OrigenRWL * 1000;
                                    txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                }
                                if (comboBoxRWLOrigen.Text.ToString() == "Kilogramos")
                                {
                                    txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                }
                                if (comboBoxRWLOrigen.Text.ToString() == "Toneladas")
                                {
                                    ValoresTablas.OrigenRWL = ValoresTablas.OrigenRWL / 1000;
                                    txtBOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                }

                                if (comboBoxRWLDestino.Text.ToString() == "Gramos")
                                {
                                    ValoresTablas.DestinoRWL = ValoresTablas.DestinoRWL * 1000;
                                    txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                }
                                if (comboBoxRWLDestino.Text.ToString() == "Kilogramos")
                                {
                                    txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                }
                                if (comboBoxRWLDestino.Text.ToString() == "Toneladas")
                                {
                                    ValoresTablas.DestinoRWL = ValoresTablas.DestinoRWL / 1000;
                                    txtBDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                }
                                //

                                //Indice de elevacion
                                if (chkBLProm.Checked == true)
                                {
                                    ValoresTablas.IndiceElevacionOrigen = ValoresTablas.PesoObjetoLProm / ValoresTablas.OrigenRWL;
                                    txtBIndiceElevacionOrigenPesoObjetoL.Text = ValoresTablas.PesoObjetoLProm.ToString();
                                    txtBIndiceElevacionOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                    txtBIndiceElevacionOrigen.Text = ValoresTablas.IndiceElevacionOrigen.ToString();

                                    ValoresTablas.IndiceElevacionDestino = ValoresTablas.PesoObjetoLProm / ValoresTablas.DestinoRWL;
                                    txtBIndiceElevacionDestinoPesoObjetoL.Text = ValoresTablas.PesoObjetoLProm.ToString();
                                    txtBIndiceElevacionDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                    txtBIndiceElevacionDestino.Text = ValoresTablas.IndiceElevacionDestino.ToString();
                                }
                                if (chkBLMax.Checked == true)
                                {
                                    ValoresTablas.IndiceElevacionOrigen = ValoresTablas.PesoObjetoLMax / ValoresTablas.OrigenRWL;
                                    txtBIndiceElevacionOrigenPesoObjetoL.Text = ValoresTablas.PesoObjetoLMax.ToString();
                                    txtBIndiceElevacionOrigenRWL.Text = ValoresTablas.OrigenRWL.ToString();
                                    txtBIndiceElevacionOrigen.Text = ValoresTablas.IndiceElevacionOrigen.ToString();

                                    ValoresTablas.IndiceElevacionDestino = ValoresTablas.PesoObjetoLMax / ValoresTablas.DestinoRWL;
                                    txtBIndiceElevacionDestinoPesoObjetoL.Text = ValoresTablas.PesoObjetoLMax.ToString();
                                    txtBIndiceElevacionDestinoRWL.Text = ValoresTablas.DestinoRWL.ToString();
                                    txtBIndiceElevacionDestino.Text = ValoresTablas.IndiceElevacionDestino.ToString();
                                }
                                //
                            }
                        }
                    }
                }
            }
        }
        

        int BuscarCasillaConLetra()
        {
            if (float.TryParse(txtBPesoObjetoLProm.Text, out ValoresTablas.PesoObjetoLProm) == false)
            {
                MessageBox.Show("La casilla Peso del objeto LProm solo admite numeros");
                return 0;
            }
            else
            {
                if (float.TryParse(txtBPesoObjetoLMax.Text, out ValoresTablas.PesoObjetoLMax) == false)
                {
                    MessageBox.Show("La casilla Peso del objeto LMáx solo admite numeros");
                    return 0;
                }
                else
                {
                    if (float.TryParse(txtBUbicacionManoOrigenH.Text, out ValoresTablas.UbicacionManoOrigenH) == false)
                    {
                        MessageBox.Show("La casilla Ubicacion de la Mano Origen H solo admite numeros");
                        return 0;
                    }
                    else
                    {
                        if (float.TryParse(txtBUbicacionManoOrigenV.Text, out ValoresTablas.UbicacionManoOrigenV) == false)
                        {
                            MessageBox.Show("La casilla Ubicacion de la Mano Origen V solo admite numeros");
                            return 0;
                        }
                        else
                        {
                            if (float.TryParse(txtBUbicacionManoDestinoH.Text, out ValoresTablas.UbicacionManoDestinoH) == false)
                            {
                                MessageBox.Show("La casilla Ubicacion de la Mano Destino H solo admite numeros");
                                return 0;
                            }
                            else
                            {
                                if (float.TryParse(txtBUbicacionManoDestinoV.Text, out ValoresTablas.UbicacionManoDestinoV) == false)
                                {
                                    MessageBox.Show("La casilla Ubicacion de la Mano Destino V solo admite numeros");
                                    return 0;
                                }
                                else
                                {
                                    if (float.TryParse(txtBDistanciaVerticalD.Text, out ValoresTablas.DistanciaVerticalD) == false)
                                    {
                                        MessageBox.Show("La casilla Distancia Vertical D  solo admite numeros");
                                        return 0;
                                    }
                                    else
                                    {
                                        if (float.TryParse(txtBAnguloAsimetricoOrigenA.Text, out ValoresTablas.AnguloAsimetricoOrigenA) == false)
                                        {
                                            MessageBox.Show("La casilla Angulo Asimetrico Origen A solo admite numeros");
                                            return 0;
                                        }
                                        else
                                        {
                                            if (float.TryParse(txtBAnguloAsimetricoDestinoA.Text, out ValoresTablas.AnguloAsimetricoDestinoA) == false)
                                            {
                                                MessageBox.Show("La casilla Angulo Asimetrico Destino A solo admite numeros");
                                                return 0;
                                            }
                                            else
                                            {
                                                if (float.TryParse(txtFMOrigen.Text, out ValoresTablas.FMOrigen) == false)
                                                {
                                                    MessageBox.Show("La casilla FM Origen solo admite numeros");
                                                    return 0;
                                                }
                                                else
                                                {
                                                    if (ValoresTablas.FMOrigen == 0)
                                                    {
                                                        MessageBox.Show("La casilla FM Origen vale 0\nPresiona el boton FM para darle un valor");
                                                        return 0;
                                                    }
                                                    else
                                                    {
                                                        if (float.TryParse(txtCMOrigen.Text, out ValoresTablas.CMOrigen) == false)
                                                        {
                                                            MessageBox.Show("La casilla CM Origen solo admite numeros");
                                                            return 0;
                                                        }
                                                        else
                                                        {
                                                            if (ValoresTablas.CMOrigen == 0)
                                                            {
                                                                MessageBox.Show("La casilla CM Origen vale 0\nPresiona el boton FM para darle un valor");
                                                                return 0;
                                                            }
                                                            else
                                                            {
                                                                if (float.TryParse(txtFMDestino.Text, out ValoresTablas.FMDestino) == false)
                                                                {
                                                                    MessageBox.Show("La casilla FM Destino solo admite numeros");
                                                                    return 0;
                                                                }
                                                                else
                                                                {
                                                                    if (ValoresTablas.FMDestino == 0)
                                                                    {
                                                                        MessageBox.Show("La casilla FM Destino vale 0\nPresiona el boton FM para darle un valor");
                                                                        return 0;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (float.TryParse(txtCMDestino.Text, out ValoresTablas.CMDestino) == false)
                                                                        {
                                                                            MessageBox.Show("La casilla CM Destino solo admite numeros");
                                                                            return 0;
                                                                        }
                                                                        else
                                                                        {
                                                                            if (ValoresTablas.CMDestino == 0)
                                                                            {
                                                                                MessageBox.Show("La casilla CM Destino vale 0\nPresiona el boton FM para darle un valor");
                                                                                return 0;
                                                                            }
                                                                            else
                                                                            {
                                                                                return 1;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        int BuscarCasillaVacia()
        {
            if (txtBPesoObjetoLProm.TextLength == 0)
            {
                MessageBox.Show("La casilla Peso del objeto LProm esta vacia");
                return 0;
            }
            else
            {
                if (txtBPesoObjetoLMax.TextLength == 0)
                {
                    MessageBox.Show("La casilla Peso del objeto LMáx esta vacia");
                    return 0;
                }
                else
                {
                    if (txtBUbicacionManoOrigenH.TextLength == 0)
                    {
                        MessageBox.Show("La casilla Ubicacion de la Mano Origen H esta vacia");
                        return 0;
                    }
                    else
                    {
                        if (txtBUbicacionManoOrigenV.TextLength == 0)
                        {
                            MessageBox.Show("La casilla Ubicacion de la Mano Origen V esta vacia");
                            return 0;
                        }
                        else
                        {
                            if (txtBUbicacionManoDestinoH.TextLength == 0)
                            {
                                MessageBox.Show("La casilla Ubicacion de la Mano Destino H esta vacia");
                                return 0;
                            }
                            else
                            {
                                if (txtBUbicacionManoDestinoV.TextLength == 0)
                                {
                                    MessageBox.Show("La casilla Ubicacion de la Mano Destino V esta vacia");
                                    return 0;
                                }
                                else
                                {
                                    if (txtBDistanciaVerticalD.TextLength == 0)
                                    {
                                        MessageBox.Show("La casilla Distancia Vertical D  esta vacia");
                                        return 0;
                                    }
                                    else
                                    {
                                        if (txtBAnguloAsimetricoOrigenA.TextLength == 0)
                                        {
                                            MessageBox.Show("La casilla Angulo Asimetrico Origen A esta vacia");
                                            return 0;
                                        }
                                        else
                                        {
                                            if (txtBAnguloAsimetricoDestinoA.TextLength == 0)
                                            {
                                                MessageBox.Show("La casilla Angulo Asimetrico Destino A esta vacia");
                                                return 0;
                                            }
                                            else
                                            {
                                                if (txtBFrecuencia.TextLength == 0)
                                                {
                                                    MessageBox.Show("La casilla Frecuencia esta vacia");
                                                    return 0;
                                                }
                                                else
                                                {
                                                    if (txtBDuracion.TextLength == 0)
                                                    {
                                                        MessageBox.Show("La casilla Duracion esta vacia");
                                                        return 0;
                                                    }
                                                    else
                                                    {
                                                        if (txtBAcoplamientoObjetos.TextLength == 0)
                                                        {
                                                            MessageBox.Show("La casilla Acoplamiento Objetos esta vacia");
                                                            return 0;
                                                        }
                                                        else
                                                        {
                                                            if (txtFMOrigen.TextLength == 0)
                                                            {
                                                                MessageBox.Show("La casilla FM Origen esta vacia \npresiona el boton para darle un valor" +
                                                                    "\n si ya lo hiciste refresca los datos");
                                                                return 0;
                                                            }
                                                            else
                                                            {
                                                                if (txtCMOrigen.TextLength == 0)
                                                                {
                                                                    MessageBox.Show("La casilla CM Origen esta vacia \npresiona el boton para darle un valor" +
                                                                        "\n si ya lo hiciste refresca los datos");
                                                                    return 0;
                                                                }
                                                                else
                                                                {
                                                                    if (txtFMDestino.TextLength == 0)
                                                                    {
                                                                        MessageBox.Show("La casilla FM Destino esta vacia \npresiona el boton para darle un valor" +
                                                                            "\n si ya lo hiciste refresca los datos");
                                                                        return 0;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (txtCMDestino.TextLength == 0)
                                                                        {
                                                                            MessageBox.Show("La casilla CM Destino esta vacia \npresiona el boton para darle un valor" +
                                                                                "\n si ya lo hiciste refresca los datos");
                                                                            return 0;
                                                                        }
                                                                        else
                                                                        {
                                                                            return 1;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void btnInvertir_Click(object sender, EventArgs e)
        {
            if (chkBSistemaIngles.Checked == false && chkBSistemaInternacional.Checked == false)
            {
                MessageBox.Show("Selecciona un Sistema de medicion para hacer la conversion");
            }
            else
            {
                if (chkBSistemaInternacional.Checked == true)
                {
                    chkBSistemaInternacional.Checked = false;
                    chkBSistemaIngles.Checked = true;
                    if (BuscarCasillaVacia() == 1 && BuscarCasillaConLetra() == 1)
                    {
                        ValoresTablas.PesoObjetoLProm = (float)(ValoresTablas.PesoObjetoLProm * 2.205);
                        ValoresTablas.PesoObjetoLMax = (float)(ValoresTablas.PesoObjetoLMax * 2.205);
                        ValoresTablas.UbicacionManoOrigenH = (float)(ValoresTablas.UbicacionManoOrigenH / 2.54);
                        ValoresTablas.UbicacionManoOrigenV = (float)(ValoresTablas.UbicacionManoOrigenV / 2.54);
                        ValoresTablas.UbicacionManoDestinoH = (float)(ValoresTablas.UbicacionManoDestinoH / 2.54);
                        ValoresTablas.UbicacionManoDestinoV = (float)(ValoresTablas.UbicacionManoDestinoV / 2.54);
                        ValoresTablas.DistanciaVerticalD = (float)(ValoresTablas.DistanciaVerticalD / 2.54);
                        LimparConversion();
                        txtBPesoObjetoSistema.Text = "Lbs";
                        txtBUbicacionManoSistema.Text = "in";
                        txtBDistanciaVerticalSistema.Text = "in";
                        comboBoxRWLOrigen.Text = "";
                        comboBoxRWLDestino.Text = "";
                        //txtBRWLOrigenSistema.Text = "Lbs";
                        //txtBRWLDestinoSistema.Text = "Lbs";
                        txtBPesoObjetoLProm.Text = ValoresTablas.PesoObjetoLProm.ToString();
                        txtBPesoObjetoLMax.Text = ValoresTablas.PesoObjetoLMax.ToString();
                        txtBUbicacionManoOrigenH.Text = ValoresTablas.UbicacionManoOrigenH.ToString();
                        txtBUbicacionManoOrigenV.Text = ValoresTablas.UbicacionManoOrigenV.ToString();
                        txtBUbicacionManoDestinoH.Text = ValoresTablas.UbicacionManoDestinoH.ToString();
                        txtBUbicacionManoDestinoV.Text = ValoresTablas.UbicacionManoDestinoV.ToString();
                        txtBDistanciaVerticalD.Text = ValoresTablas.DistanciaVerticalD.ToString();
                    }
                }
                else
                {
                    if (chkBSistemaIngles.Checked == true)
                    {
                        chkBSistemaIngles.Checked = false;
                        chkBSistemaInternacional.Checked = true;
                        if (BuscarCasillaVacia() == 1 && BuscarCasillaConLetra() == 1)
                        {
                            ValoresTablas.PesoObjetoLProm = (float)(ValoresTablas.PesoObjetoLProm / 2.205);
                            ValoresTablas.PesoObjetoLMax = (float)(ValoresTablas.PesoObjetoLMax / 2.205);
                            ValoresTablas.UbicacionManoOrigenH = (float)(ValoresTablas.UbicacionManoOrigenH * 2.54);
                            ValoresTablas.UbicacionManoOrigenV = (float)(ValoresTablas.UbicacionManoOrigenV * 2.54);
                            ValoresTablas.UbicacionManoDestinoH = (float)(ValoresTablas.UbicacionManoDestinoH * 2.54);
                            ValoresTablas.UbicacionManoDestinoV = (float)(ValoresTablas.UbicacionManoDestinoV * 2.54);
                            ValoresTablas.DistanciaVerticalD = (float)(ValoresTablas.DistanciaVerticalD * 2.54);
                            LimparConversion();
                            txtBPesoObjetoSistema.Text = "Kg";
                            txtBUbicacionManoSistema.Text = "cm";
                            txtBDistanciaVerticalSistema.Text = "cm";
                            comboBoxRWLOrigen.Text = "";
                            comboBoxRWLDestino.Text = "";
                            //txtBRWLOrigenSistema.Text = "Kg";
                            //txtBRWLDestinoSistema.Text = "Kg";
                            txtBPesoObjetoLProm.Text = ValoresTablas.PesoObjetoLProm.ToString();
                            txtBPesoObjetoLMax.Text = ValoresTablas.PesoObjetoLMax.ToString();
                            txtBUbicacionManoOrigenH.Text = ValoresTablas.UbicacionManoOrigenH.ToString();
                            txtBUbicacionManoOrigenV.Text = ValoresTablas.UbicacionManoOrigenV.ToString();
                            txtBUbicacionManoDestinoH.Text = ValoresTablas.UbicacionManoDestinoH.ToString();
                            txtBUbicacionManoDestinoV.Text = ValoresTablas.UbicacionManoDestinoV.ToString();
                            txtBDistanciaVerticalD.Text = ValoresTablas.DistanciaVerticalD.ToString();
                        }
                    }
                }
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Image Files (*.bmp, *.jpeg)|*.bmp;*.jpeg";
            saveFileDialog1.Title = "Guardar Imagen Metodo niosh";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Bitmap bm = new Bitmap(groupBox1.Width, groupBox1.Height);
                groupBox1.DrawToBitmap(bm, new Rectangle(0, 0, bm.Width, bm.Height));
                bm.Save(saveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                saveFileDialog1.FileName = "";
            }
            else
            {
                MessageBox.Show("Hubo un error a la hora de guardar pulsa otra vez el boton");
            }
        }

        private void btnAM_Click(object sender, EventArgs e)
        {
            AM_Form formularioAM = AM_Form.ventanaUnica();
            formularioAM.Show();
            formularioAM.BringToFront();
        }

        private void btnDM_Click(object sender, EventArgs e)
        {
            DM_Form formularioDM = DM_Form.ventanaUnica();
            formularioDM.Show();
            formularioDM.BringToFront();
        }

        private void btnVM_Click(object sender, EventArgs e)
        {
            VM_Form formularioVM = VM_Form.ventanaUnica();
            formularioVM.Show();
            formularioVM.BringToFront();
        }

        private void btnHM_Click(object sender, EventArgs e)
        {
            HM_Form formularioHM = HM_Form.ventanaUnica();
            formularioHM.Show();
            formularioHM.BringToFront();
        }

        private void btnLC_Click(object sender, EventArgs e)
        {
            LC_Form formularioLC = LC_Form.ventanaUnica();
            formularioLC.Show();
            formularioLC.BringToFront();
        }

        private void btnA_Click(object sender, EventArgs e)
        {
            A_Form formularioA = A_Form.ventanaUnica();
            formularioA.Show();
            formularioA.BringToFront();
        }

        private void btnD_Click(object sender, EventArgs e)
        {
            D_Form formularioD = D_Form.ventanaUnica();
            formularioD.Show();
            formularioD.BringToFront();
        }

        private void btnV_Click(object sender, EventArgs e)
        {
            V_Form formularioV = V_Form.ventanaUnica();
            formularioV.Show();
            formularioV.BringToFront();
        }

        private void btnH_Click(object sender, EventArgs e)
        {
            H_Form formularioH = H_Form.ventanaUnica();
            formularioH.Show();
            formularioH.BringToFront();
        }

        private void btnL_Click(object sender, EventArgs e)
        {
            L_Form formularioL = L_Form.ventanaUnica();
            formularioL.Show();
            formularioL.BringToFront();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Portada formularioPortada = Portada.ventanaUnica();
            formularioPortada.Show();
            this.BringToFront();
        }

        private void btnH_Click_1(object sender, EventArgs e)
        {
            H_Form formularioH = H_Form.ventanaUnica();
            formularioH.Show();
            formularioH.BringToFront();
        }

        private void HojaDeTrabajoNiosh_FormClosing(object sender, FormClosingEventArgs e)
        {
            instancia = null;
        }

        private void comboBoxRWLOrigen_Click(object sender, EventArgs e)
        {
            if (chkBSistemaIngles.Checked == true)
            {
                comboBoxRWLOrigen.Items.Clear();
                comboBoxRWLOrigen.Items.Add("Mili-Libras");
                comboBoxRWLOrigen.Items.Add("Libras");
                comboBoxRWLOrigen.Items.Add("Kilo-Libras");

                comboBoxRWLDestino.Items.Clear();
                comboBoxRWLDestino.Items.Add("Mili-Libras");
                comboBoxRWLDestino.Items.Add("Libras");
                comboBoxRWLDestino.Items.Add("Kilo-Libras");
                
            }
            else
            {
                if(chkBSistemaInternacional.Checked == true)
                {
                    comboBoxRWLOrigen.Items.Clear();
                    comboBoxRWLOrigen.Items.Add("Gramos");
                    comboBoxRWLOrigen.Items.Add("Kilogramos");
                    comboBoxRWLOrigen.Items.Add("Toneladas");

                    comboBoxRWLDestino.Items.Clear();
                    comboBoxRWLDestino.Items.Add("Gramos");
                    comboBoxRWLDestino.Items.Add("Kilogramos");
                    comboBoxRWLDestino.Items.Add("Toneladas");
                }
            }
        }

        private void botonGuardarExcel_Click(object sender, EventArgs e)
        {

            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog1.Title = "Guardar Excel Metodo niosh";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SLDocument MetodoNioshResultadosExcel = new SLDocument();

                DataTable datos_A_Guardar = new DataTable();

                ///Paso 1
                datos_A_Guardar.Columns.Add("PesoObjetoLMax", typeof(float));
                datos_A_Guardar.Columns.Add("UbicacionManoOrigenH", typeof(float));
                datos_A_Guardar.Columns.Add("UbicacionManoOrigenV", typeof(float));
                datos_A_Guardar.Columns.Add("UbicacionManoDestinoH", typeof(float));
                datos_A_Guardar.Columns.Add("UbicacionManoDestinoV", typeof(float));
                datos_A_Guardar.Columns.Add("DistanciaVerticalD", typeof(float));
                datos_A_Guardar.Columns.Add("AnguloAsimetricoOrigenA", typeof(float));
                datos_A_Guardar.Columns.Add("AnguloAsimetricoDestinoA", typeof(float));
                datos_A_Guardar.Columns.Add("Frecuencia", typeof(string));
                datos_A_Guardar.Columns.Add("Duracion", typeof(string));
                datos_A_Guardar.Columns.Add("AcoplamientoObjetos", typeof(string));
                //Paso 2
                datos_A_Guardar.Columns.Add("LCOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("LCDestino", typeof(float));
                datos_A_Guardar.Columns.Add("HMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("HMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("VMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("VMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("DMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("DMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("AMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("AMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("CMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("FMOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("CMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("FMDestino", typeof(float));
                datos_A_Guardar.Columns.Add("OrigenRWL", typeof(float));
                datos_A_Guardar.Columns.Add("DestinoRWL", typeof(float));
                //Paso 3
                datos_A_Guardar.Columns.Add("IndiceElevacionOrigen", typeof(float));
                datos_A_Guardar.Columns.Add("IndiceElevacionDestino", typeof(float));

                DataRow Renglon = datos_A_Guardar.NewRow();
                //Paso 1
                Renglon["PesoObjetoLMax"] = ValoresTablas.PesoObjetoLMax;
                Renglon["UbicacionManoOrigenH"] = ValoresTablas.UbicacionManoOrigenH;
                Renglon["UbicacionManoOrigenV"] = ValoresTablas.UbicacionManoOrigenV;
                Renglon["UbicacionManoDestinoH"] = ValoresTablas.UbicacionManoDestinoH;
                Renglon["UbicacionManoDestinoV"] = ValoresTablas.UbicacionManoDestinoV;
                Renglon["DistanciaVerticalD"] = ValoresTablas.DistanciaVerticalD;
                Renglon["AnguloAsimetricoOrigenA"] = ValoresTablas.AnguloAsimetricoOrigenA;
                Renglon["AnguloAsimetricoDestinoA"] = ValoresTablas.AnguloAsimetricoDestinoA;
                Renglon["Frecuencia"] = txtBFrecuencia.Text.ToString();
                Renglon["Duracion"] = txtBDuracion.Text.ToString();
                Renglon["AcoplamientoObjetos"] = txtBAcoplamientoObjetos.Text.ToString();
                //Paso 2
                Renglon["LCOrigen"] = ValoresTablas.LCOrigen;
                Renglon["LCDestino"] = ValoresTablas.LCDestino;
                Renglon["HMOrigen"] = ValoresTablas.HMOrigen;
                Renglon["HMDestino"] = ValoresTablas.HMDestino;
                Renglon["VMOrigen"] = ValoresTablas.VMOrigen;
                Renglon["VMDestino"] = ValoresTablas.VMDestino;
                Renglon["DMOrigen"] = ValoresTablas.DMOrigen;
                Renglon["DMDestino"] = ValoresTablas.DMDestino;
                Renglon["AMOrigen"] = ValoresTablas.AMOrigen;
                Renglon["AMDestino"] = ValoresTablas.AMDestino;
                Renglon["CMOrigen"] = ValoresTablas.CMOrigen;
                Renglon["FMOrigen"] = ValoresTablas.FMOrigen;
                Renglon["CMDestino"] = ValoresTablas.CMDestino;
                Renglon["FMDestino"] = ValoresTablas.FMDestino;
                Renglon["OrigenRWL"] = ValoresTablas.OrigenRWL;
                Renglon["DestinoRWL"] = ValoresTablas.DestinoRWL;
                //Paso 3
                Renglon["IndiceElevacionOrigen"] = ValoresTablas.IndiceElevacionOrigen;
                Renglon["IndiceElevacionDestino"] = ValoresTablas.IndiceElevacionDestino;

                datos_A_Guardar.Rows.Add(Renglon);
                MetodoNioshResultadosExcel.ImportDataTable(1, 1, datos_A_Guardar, true);
                MetodoNioshResultadosExcel.SaveAs(saveFileDialog1.FileName);
                saveFileDialog1.FileName = "";
            }
            else
            {
                MessageBox.Show("Hubo un error a la hora de guardar pulsa otra vez el boton");
            }
        } 
    }
}
