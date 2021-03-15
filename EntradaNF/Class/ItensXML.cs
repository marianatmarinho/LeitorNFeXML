using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EntradaNF.Class
{
    class ItensXML
    {
        #region Atributos

        #region Atributos da Nota Fiscal de Entrada

        public int sequenciaItem;
        public string produtoNF;
        public string descricaoNF;
        public decimal qtdeNF;
        public string unidNF;
        public decimal valorUnitario;
        public decimal valorTotal;

        public string NCM;
        public string NCMFornec; //Classificação Fiscal do Fornecedor
        public string CFOP;
        public string EAN;

        public string origemCSTICMS;
        public string CSTICMS;
        public decimal baseICMS;
        public decimal percICMS;
        public decimal valorICMS;
        public decimal percRedICMS;

        public decimal baseST;
        public decimal percST;
        public decimal valorST;
        public decimal percRedST;

        public string CSTIPI;
        public decimal baseIPI;
        public decimal percIPI;
        public decimal valorIPI;

        public string CSTPIS;
        public decimal basePIS;
        public decimal percPIS;
        public decimal valorPIS;

        public string CSTCOFINS;
        public decimal baseCOFINS;
        public decimal percCOFINS;
        public decimal valorCOFINS;

        public decimal valorDesconto;

        public string infoAdicionalProduto;
        public string descricaoNFInfoAdicionalProduto;

        #endregion
        
        #endregion
    }
}
