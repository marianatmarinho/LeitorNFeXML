using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EntradaNF.Class;

namespace EntradaNF
{
    public partial class FrmNFe : Form
    {
        #region ATRIBUTOS

        private LerXML impNotaFiscal = new LerXML();
        private List<ItensXML> listaNotaFiscalItens = new List<ItensXML>();
        private List<ParcelasXML> listaNotaFiscalParcelas = new List<ParcelasXML>();

        #endregion

        #region Construtor

        public FrmNFe()
        {
            InitializeComponent();
        }

        #endregion

        #region MÉTODOS

        private void PersonalizarGridItens()
        {
            grdItens.Columns.Add("seq", "Sequencia");
            grdItens.Columns.Add("produto", "Produto");
            grdItens.Columns.Add("descricao", "Descrição");
            grdItens.Columns.Add("unid", "Unidade");
            grdItens.Columns.Add("qtde", "Quantidade");
            grdItens.Columns.Add("vlUnit", "Valor Unitário");
            grdItens.Columns.Add("vlTotal", "Valor Total");

            grdItens.Columns["descricao"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void PreencherGridItens()
        {
            PersonalizarGridItens();
            foreach (ItensXML item in listaNotaFiscalItens)
            {
                grdItens.Rows.Add(item.sequenciaItem,
                                   item.produtoNF,
                                   item.descricaoNF,
                                   item.unidNF,
                                   item.qtdeNF,
                                   item.valorUnitario,
                                   item.valorTotal);
            }
        }

        private void PersonalizarGridParcelas()
        {
            grdParcelas.Columns.Add("seq", "Sequencia");
            grdParcelas.Columns.Add("descricao", "Descrição");
            grdParcelas.Columns.Add("data", "Data");
            grdParcelas.Columns.Add("valor", "Valor");

            grdParcelas.Columns["descricao"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void PreencherGridParcelas()
        {
            PersonalizarGridParcelas();

            foreach (ParcelasXML parcela in listaNotaFiscalParcelas)
            {
                grdParcelas.Rows.Add(parcela.ordemParcela,
                                  parcela.descricaoParcela,
                                  parcela.dataParcela,
                                  parcela.valorParcela);
            }
        }

        private void PreencherGrid()
        {
            PreencherGridItens();
            PreencherGridParcelas();
        }

        private bool CarregarDadosXMLNotaFiscal()
        {
            bool sucesso = false;
            //Carregar dados do cabeçalho da Nota Fiscal
            try
            {
                impNotaFiscal.CarregarDadosXMLNotaFiscal(txtCaminho.Text);
            }
            catch
            {
                MessageBox.Show("Erro ao carregar o XML! Verifique se ele está no formato correto.", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //Popula os campos do cabeçalho da Nota Fiscal
            txtNF.Text = impNotaFiscal.NotaFiscal;
            txtSerie.Text = impNotaFiscal.Serie;
            txtDtEmissao.Text = impNotaFiscal.DataEmissao.ToShortDateString();
            txtCNPJ.Text = impNotaFiscal.DestCNPJ;
            txtRazaoSocial.Text = impNotaFiscal.DestRSocial;
            txtNatureza.Text = impNotaFiscal.NatOp;

            try
            {
                //Carrega dados dos Itens da Nota Fiscal
                listaNotaFiscalItens = impNotaFiscal.CarregarDadosXMLNotaFiscalItens(txtCaminho.Text);
            }
            catch
            {
                MessageBox.Show("Erro ao carregar o XML! Verifique se ele está no formato correto ou se foi protocolado na Receita Federal.", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //Carrega dados das Parcelas da Nota Fiscal
            listaNotaFiscalParcelas = impNotaFiscal.CarregarDadosXMLNotaFiscalParcelas(txtCaminho.Text);

            //Carrega o Grid dos Itens da Nota Fiscal
            PreencherGrid();

            //Carregar a Aba Totais
            txtObs.Text = impNotaFiscal.InfoComplementar;


            txtICMS.Text = impNotaFiscal.ValorICMS.ToString();
            txtST.Text = impNotaFiscal.ValorST.ToString();

            txtMercadoria.Text = impNotaFiscal.ValorProdutos.ToString();
            txtFrete.Text = impNotaFiscal.ValorFrete.ToString();
            txtSeguro.Text = impNotaFiscal.ValorSeguro.ToString();
            txtDesconto.Text = impNotaFiscal.ValorDescontos.ToString();
            txtDespesas.Text = impNotaFiscal.ValorDespesasAcessorias.ToString();
            txtIPI.Text = impNotaFiscal.ValorIPI.ToString();
            txtPIS.Text = impNotaFiscal.ValorPIS.ToString();
            txtCofins.Text = impNotaFiscal.ValorCOFINS.ToString();
            txtOutros.Text = impNotaFiscal.ValorOutros.ToString();
            txtContabil.Text = impNotaFiscal.ValorNotaFiscal.ToString();

            return sucesso;
        }

   
        #endregion

        #region EVENTOS

        private void btnPesqCaminho_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Localizar arquivo XML";
                openFileDialog.Multiselect = false;

                openFileDialog.Filter = "Dados XML|*.xml;";

                DialogResult dlgResult = openFileDialog.ShowDialog();
                if ((dlgResult == DialogResult.OK) || (dlgResult == DialogResult.Yes))
                {
                    txtCaminho.Text = openFileDialog.FileName;
                }
                else
                {
                    txtCaminho.Clear();
                }
            }
            catch (Exception ex)
            {
                txtCaminho.Clear();

                MessageBox.Show("Erro ao tentar selecionar o arquivo!\n" + ex.ToString(), "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnLerXML_Click(object sender, EventArgs e)
        {

            CarregarDadosXMLNotaFiscal();
        }
        
    }
        #endregion
}
