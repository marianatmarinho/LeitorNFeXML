using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;

namespace EntradaNF.Class
{
    class LerXML
    {
        #region Atributos

        //Cabeçalho
        private string destCNPJ;
        private string destRSocial;
        private string emitCNPJ;
        private string notaFiscal;
        private string serie;
        private DateTime dataEmissao;
        private string natOp;
        private string autorizacao;

        //Totais
        private decimal baseICMS;
        private decimal valorICMS;
        private decimal baseST;
        private decimal valorST;
        private decimal valorProdutos;
        private decimal valorFrete;
        private decimal valorSeguro;
        private decimal valorDescontos;
        private decimal valorDespesasAcessorias;
        private decimal valorIPI;
        private decimal valorPIS;
        private decimal valorCOFINS;
        private decimal valorNotaFiscal;
        private decimal valorOutros;
        
        //Transporte
        private int modalidadeFrete;
        private string transpCNPJ;
        private string transpPlaca;
        private string transpPlacaUF;
        private decimal transpVolumeQtde;
        private string transpEspecie;
        private string transpMarca;
        private string transpVolumeNumero;
        private decimal transpPesoBruto;
        private decimal transpPesoLiquido;

        //Informações complementares
        private string infoComplementar;

        #endregion

        #region Propriedades

        public string DestCNPJ
        {
            get { return destCNPJ; }
            set { destCNPJ = value; }
        }

        public string DestRSocial
        {
            get { return destRSocial; }
            set { destRSocial = value; }
        }

        public string EmitCNPJ
        {
            get { return emitCNPJ; }
            set { emitCNPJ = value; }
        }

        public string NotaFiscal
        {
            get { return notaFiscal; }
            set { notaFiscal = value; }
        }

        public string Serie
        {
            get { return serie; }
            set { serie = value; }
        }

        public DateTime DataEmissao
        {
            get { return dataEmissao; }
            set { dataEmissao = value; }
        }

        public string NatOp
        {
            get { return natOp; }
            set { natOp = value; }
        }

        public string Autorizacao
        {
            get { return autorizacao; }
            set { autorizacao = value; }
        }

        public decimal BaseICMS
        {
            get { return baseICMS; }
            set { baseICMS = value; }
        }

        public decimal ValorICMS
        {
            get { return valorICMS; }
            set { valorICMS = value; }
        }

        public decimal BaseST
        {
            get { return baseST; }
            set { baseST = value; }
        }

        public decimal ValorST
        {
            get { return valorST; }
            set { valorST = value; }
        }

        public decimal ValorProdutos
        {
            get { return valorProdutos; }
            set { valorProdutos = value; }
        }

        public decimal ValorFrete
        {
            get { return valorFrete; }
            set { valorFrete = value; }
        }

        public decimal ValorSeguro
        {
            get { return valorSeguro; }
            set { valorSeguro = value; }
        }

        public decimal ValorDescontos
        {
            get { return valorDescontos; }
            set { valorDescontos = value; }
        }

        public decimal ValorDespesasAcessorias
        {
            get { return valorDespesasAcessorias; }
            set { valorDespesasAcessorias = value; }
        }

        public decimal ValorIPI
        {
            get { return valorIPI; }
            set { valorIPI = value; }
        }

        public decimal ValorPIS
        {
            get { return valorPIS; }
            set { valorPIS = value; }
        }

        public decimal ValorCOFINS
        {
            get { return valorCOFINS; }
            set { valorCOFINS = value; }
        }

        public decimal ValorNotaFiscal
        {
            get { return valorNotaFiscal; }
            set { valorNotaFiscal = value; }
        }

        public decimal ValorOutros
        {
            get { return valorOutros; }
            set { valorOutros = value; }
        }

        public string InfoComplementar
        {
            get { return infoComplementar; }
            set { infoComplementar = value; }
        }

        /// <summary>
        /// Modalidade do frete 0- Por conta do emitente; 1- Por conta do destinatário/remetente; 2- Por conta de terceiros; 9- Sem frete (v2.0)
        /// </summary>
        /// 
        public int ModalidadeFrete
        {
            get { return modalidadeFrete; }
            set { modalidadeFrete = value; }
        }

        public string TranspCNPJ
        {
            get { return transpCNPJ; }
            set { transpCNPJ = value; }
        }

        public string TranspPlaca
        {
            get { return transpPlaca; }
            set { transpPlaca = value; }
        }

        public string TranspPlacaUF
        {
            get { return transpPlacaUF; }
            set { transpPlacaUF = value; }
        }

        public decimal TranspVolumeQtde
        {
            get { return transpVolumeQtde; }
            set { transpVolumeQtde = value; }
        }

        public string TranspEspecie
        {
            get { return transpEspecie; }
            set { transpEspecie = value; }
        }

        public string TranspMarca
        {
            get { return transpMarca; }
            set { transpMarca = value; }
        }

        public string TranspVolumeNumero
        {
            get { return transpVolumeNumero; }
            set { transpVolumeNumero = value; }
        }

        public decimal TranspPesoBruto
        {
            get { return transpPesoBruto; }
            set { transpPesoBruto = value; }
        }

        public decimal TranspPesoLiquido
        {
            get { return transpPesoLiquido; }
            set { transpPesoLiquido = value; }
        }

        public void Limpar()
        {
            //Cabeçalho
            destCNPJ = string.Empty;
            emitCNPJ = string.Empty;
            notaFiscal = string.Empty;
            serie = string.Empty;
            natOp = string.Empty;
            autorizacao = string.Empty;

            //Totais
            baseICMS = 0;
            valorICMS = 0;
            baseST = 0;
            valorST = 0;
            valorProdutos = 0;
            valorFrete = 0;
            valorSeguro = 0;
            valorDescontos = 0;
            valorDespesasAcessorias = 0;
            valorIPI = 0;
            valorPIS = 0;
            valorCOFINS = 0;
            valorNotaFiscal = 0;
            valorOutros = 0;

            //Transporte
            modalidadeFrete = 0;
            transpCNPJ = string.Empty;
            transpPlaca = string.Empty;
            transpPlacaUF = string.Empty;
            transpVolumeQtde = 0;
            transpEspecie = string.Empty;
            transpMarca = string.Empty;
            transpVolumeNumero = string.Empty;
            transpPesoBruto = 0;
            transpPesoLiquido = 0;

            //Informações complementares
            infoComplementar = string.Empty;
        }
        #endregion

        #region Métodos

        /// <summary>
        /// Carrega os dados da Nota Fiscal a partir do arquivo XML gerado pela Receita
        /// Necessário informar o caminho de rede completo onde o arquivo XML está
        /// </summary>
        public void CarregarDadosXMLNotaFiscal(string strCaminhoXMLNotaFiscal)
        {
            Limpar();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNamespaceManager xmlTag = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlDoc.Load(strCaminhoXMLNotaFiscal);
            xmlTag.AddNamespace("nfe", xmlDoc.DocumentElement.NamespaceURI);

            XmlNodeList xnDest = xmlDoc.GetElementsByTagName("dest");
            foreach (XmlNode xnD in xnDest)
            {
                destCNPJ = xnD["CNPJ"] == null ? destCNPJ : xnD["CNPJ"].InnerText;
                destRSocial = xnD["xNome"] == null ? destRSocial : xnD["xNome"].InnerText;
            }

            XmlNodeList xnEmit = xmlDoc.GetElementsByTagName("emit");
            foreach (XmlNode xnE in xnEmit)
            {
                emitCNPJ = xnE["CNPJ"] == null ? emitCNPJ : xnE["CNPJ"].InnerText;
            }

            XmlNodeList xnIde = xmlDoc.GetElementsByTagName("ide");
            foreach (XmlNode xnId in xnIde)
            {
                notaFiscal = xnId["nNF"] == null ? notaFiscal : xnId["nNF"].InnerText;
                serie = xnId["serie"] == null ? serie : xnId["serie"].InnerText;

                if (xnId["dEmi"] != null)
                {
                    string data = xnId["dEmi"].InnerText;
                    DateTime.TryParse(data.Split('-')[2] + "/" + data.Split('-')[1] + "/" + data.Split('-')[0], out dataEmissao);
                }
                else if (xnId["dhEmi"] != null)
                {
                    string data = xnId["dhEmi"].InnerText.Split('T')[0];
                    string hora = xnId["dhEmi"].InnerText.Split('T')[1].Split('-')[0];
                    DateTime.TryParse(data.Split('-')[2] + "/" + data.Split('-')[1] + "/" + data.Split('-')[0] + " " + hora, out dataEmissao);
                }

                natOp = xnId["natOp"] == null ? natOp : xnId["natOp"].InnerText;
            }

            //Carregar dados de Informação de  Protocolos
            XmlNodeList xnInfoProt = xmlDoc.GetElementsByTagName("infProt");
            foreach (XmlNode xnIp in xnInfoProt)
            {
                autorizacao = xnIp["nProt"] == null ? autorizacao : xnIp["nProt"].InnerText;
            }

            //Carregar as inforamções de totais da nota fiscal
            XmlNodeList xnTotais = xmlDoc.GetElementsByTagName("ICMSTot");
            foreach (XmlNode xnTt in xnTotais)
            {
                baseICMS = xnTt["vBC"] == null ? baseICMS : Convert.ToDecimal(xnTt["vBC"].InnerText.Replace(".", ","));
                valorICMS = xnTt["vICMS"] == null ? valorICMS : Convert.ToDecimal(xnTt["vICMS"].InnerText.Replace(".", ","));
                baseST = xnTt["vBCST"] == null ? baseST : Convert.ToDecimal(xnTt["vBCST"].InnerText.Replace(".", ","));
                valorST = xnTt["vST"] == null ? valorST : Convert.ToDecimal(xnTt["vST"].InnerText.Replace(".", ","));

                valorProdutos = xnTt["vProd"] == null ? valorProdutos : Convert.ToDecimal(xnTt["vProd"].InnerText.Replace(".", ","));

                valorFrete = xnTt["vFrete"] == null ? valorFrete : Convert.ToDecimal(xnTt["vFrete"].InnerText.Replace(".", ","));
                valorSeguro = xnTt["vSeg"] == null ? valorSeguro : Convert.ToDecimal(xnTt["vSeg"].InnerText.Replace(".", ","));
                valorDescontos = xnTt["vDesc"] == null ? valorDescontos : Convert.ToDecimal(xnTt["vDesc"].InnerText.Replace(".", ","));

                //valorDespesasAcessorias

                valorIPI = xnTt["vIPI"] == null ? valorIPI : Convert.ToDecimal(xnTt["vIPI"].InnerText.Replace(".", ","));
                valorPIS = xnTt["vPIS"] == null ? valorPIS : Convert.ToDecimal(xnTt["vPIS"].InnerText.Replace(".", ","));
                valorCOFINS = xnTt["vCOFINS"] == null ? valorCOFINS : Convert.ToDecimal(xnTt["vCOFINS"].InnerText.Replace(".", ","));
                valorOutros = xnTt["vOutro"] == null ? valorOutros : Convert.ToDecimal(xnTt["vOutro"].InnerText.Replace(".", ","));

                valorNotaFiscal = xnTt["vNF"] == null ? valorNotaFiscal : Convert.ToDecimal(xnTt["vNF"].InnerText.Replace(".", ","));
            }

            //Carregar as informações de transporte
            XmlNodeList xnTransporte = xmlDoc.GetElementsByTagName("transp");
            foreach (XmlNode xnTransp in xnTransporte)
            {
                modalidadeFrete = xnTransp["modFrete"] == null ? modalidadeFrete : Convert.ToInt32(xnTransp["modFrete"].InnerText);
            }
            XmlNodeList xnTransporta = xmlDoc.GetElementsByTagName("transporta");
            foreach (XmlNode xnTrans in xnTransporta)
            {
                transpCNPJ = xnTrans["CNPJ"] == null ? transpCNPJ : xnTrans["CNPJ"].InnerText;
            }
            XmlNodeList xnTranspVeiculo = xmlDoc.GetElementsByTagName("veicTransp");
            foreach (XmlNode xnVei in xnTranspVeiculo)
            {
                transpPlaca = xnVei["placa"] == null ? transpPlaca : xnVei["placa"].InnerText;
                transpPlacaUF = xnVei["UF"] == null ? transpPlacaUF : xnVei["UF"].InnerText;
            }
            XmlNodeList xnTranspVolume = xmlDoc.GetElementsByTagName("vol");
            foreach (XmlNode xnVol in xnTranspVolume)
            {
                transpVolumeQtde = xnVol["qVol"] == null ? transpVolumeQtde : Convert.ToDecimal(xnVol["qVol"].InnerText.Replace(".", ","));
                transpEspecie = xnVol["esp"] == null ? transpEspecie : xnVol["esp"].InnerText;
                transpMarca = xnVol["marca"] == null ? transpMarca : xnVol["marca"].InnerText;
                transpVolumeNumero = xnVol["nVol"] == null ? transpVolumeNumero : xnVol["nVol"].InnerText;
                transpPesoBruto = xnVol["pesoB"] == null ? transpPesoBruto : Convert.ToDecimal(xnVol["pesoB"].InnerText.Replace(".", ","));
                transpPesoLiquido = xnVol["pesoL"] == null ? transpPesoLiquido : Convert.ToDecimal(xnVol["pesoL"].InnerText.Replace(".", ","));
            }

            //Informação Adicional
            XmlNodeList xnInfoAdic = xmlDoc.GetElementsByTagName("infAdic");
            foreach (XmlNode xnIa in xnInfoAdic)
            {
                infoComplementar = xnIa["infCpl"] == null ? infoComplementar : xnIa["infCpl"].InnerText;
            }
        }

        /// <summary>
        /// Carrega os dados os Itens da Nota Fiscal a partir do arquivo XML gerado pela Receita
        /// Necessário informar o caminho de rede completo onde o arquivo XML está
        /// <param name="strCaminhoXMLNotaFiscal">Obrigatório - Caminhao físico completo do XML</param>
        /// <returns>Lista de Itens da Nota Fiscal (ItensXML)</returns>
        public List<RefXML> CarregarDadosXMLNotaFiscalRefNFe(string strCaminhoXMLNotaFiscal)
        {
            List<RefXML> lista = new List<RefXML>();

            XmlDocument xmlDoc = new XmlDocument();
            XmlNamespaceManager xmlTag = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlDoc.Load(strCaminhoXMLNotaFiscal);
            xmlTag.AddNamespace("nfe", xmlDoc.DocumentElement.NamespaceURI);

            XmlNodeList xnNFref = xmlDoc.GetElementsByTagName("NFref");
            foreach (XmlNode xnNF in xnNFref)
            {
                if (xnNF["refNFe"] != null)
                {
                    RefXML refNFe = new RefXML
                    {
                        chaveAcessoRefNFe = xnNF["refNFe"].InnerText
                    };

                    lista.Add(refNFe);
                }
            }

            return lista;
        }

        /// <summary>
        /// Carrega os dados os Itens da Nota Fiscal a partir do arquivo XML gerado pela Receita
        /// Necessário informar o caminho de rede completo onde o arquivo XML está
        /// <param name="strCaminhoXMLNotaFiscal">Obrigatório - Caminhao físico completo do XML</param>
        /// <returns>Lista de Itens da Nota Fiscal (ItensXML)</returns>
        public List<ItensXML> CarregarDadosXMLNotaFiscalItens(string strCaminhoXMLNotaFiscal)
        {
            List<ItensXML> lista = new List<ItensXML>();

            XmlDocument xmlDoc = new XmlDocument();
            XmlNamespaceManager xmlTag = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlDoc.Load(strCaminhoXMLNotaFiscal);
            xmlTag.AddNamespace("nfe", xmlDoc.DocumentElement.NamespaceURI);

            XmlNodeList nfeProc = xmlDoc.GetElementsByTagName("nfeProc");
            XmlNodeList nItem = ((XmlElement)nfeProc[0]).GetElementsByTagName("det");

            int sequenciaItem = 0;
            foreach (XmlElement nodo in nItem)
            {
                XmlNodeList xnProd = nodo.GetElementsByTagName("prod");
                {
                    //Incrementa a sequencia do item
                    sequenciaItem++;

                    //Cria Item
                    ItensXML item = new ItensXML();

                    //Sequencia do item
                    item.sequenciaItem = sequenciaItem;

                    //Produto
                    if (xnProd[0] != null)
                    {
                        item.produtoNF = xnProd[0]["cProd"] == null ? item.produtoNF : xnProd[0]["cProd"].InnerText;
                        item.descricaoNF = xnProd[0]["xProd"] == null ? item.descricaoNF : xnProd[0]["xProd"].InnerText;
                        item.qtdeNF = xnProd[0]["qCom"] == null ? item.qtdeNF : Convert.ToDecimal(xnProd[0]["qCom"].InnerText.Replace(".", ","));
                        item.unidNF = xnProd[0]["uCom"] == null ? item.unidNF : xnProd[0]["uCom"].InnerText;
                        item.unidNF = (item.unidNF.Length > 3 ? item.unidNF.Substring(0, 2).ToUpper() : item.unidNF.ToUpper());
                        item.valorUnitario = xnProd[0]["vUnCom"] == null ? item.valorUnitario : Convert.ToDecimal(xnProd[0]["vUnCom"].InnerText.Replace(".", ","));
                        item.valorTotal = xnProd[0]["vProd"] == null ? item.valorTotal : Convert.ToDecimal(xnProd[0]["vProd"].InnerText.Replace(".", ","));

                        item.CFOP = xnProd[0]["CFOP"] == null ? item.CFOP : xnProd[0]["CFOP"].InnerText.Replace(".", ",");
                        item.EAN = xnProd[0]["cEAN"] == null ? item.EAN : xnProd[0]["cEAN"].InnerText;
                        item.NCM = xnProd[0]["NCM"] == null ? item.NCM : xnProd[0]["NCM"].InnerText;
                        item.NCMFornec = xnProd[0]["NCM"] == null ? item.NCM : xnProd[0]["NCM"].InnerText;
                        item.valorDesconto = xnProd[0]["vDesc"] == null ? 0 : Convert.ToDecimal(xnProd[0]["vDesc"].InnerText.Replace(".", ","));
                    }

                    //ICMS00
                    XmlNodeList xnIICMS = nodo.GetElementsByTagName("ICMS00");
                    {
                        if (xnIICMS[0] != null)
                        {
                            item.baseICMS = xnIICMS[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS[0]["vICMS"].InnerText.Replace(".", ","));
                            item.origemCSTICMS = xnIICMS[0]["orig"] == null ? item.origemCSTICMS : xnIICMS[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS[0]["CST"] == null ? item.CSTICMS : xnIICMS[0]["CST"].InnerText;
                        }
                    }

                    //ICMS10
                    XmlNodeList xnIICMS10 = nodo.GetElementsByTagName("ICMS10");
                    {
                        if (xnIICMS10[0] != null)
                        {
                            item.baseICMS = xnIICMS10[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS10[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS10[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS10[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS10[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS10[0]["vICMS"].InnerText.Replace(".", ","));

                            item.percRedST = xnIICMS10[0]["pRedBCST"] == null ? item.percRedST : Convert.ToDecimal(xnIICMS10[0]["pRedBCST"].InnerText.Replace(".", ","));
                            item.baseST = xnIICMS10[0]["vBCST"] == null ? item.baseST : Convert.ToDecimal(xnIICMS10[0]["vBCST"].InnerText.Replace(".", ","));
                            item.percST = xnIICMS10[0]["pICMSST"] == null ? item.percST : Convert.ToDecimal(xnIICMS10[0]["pICMSST"].InnerText.Replace(".", ","));
                            item.valorST = xnIICMS10[0]["vICMSST"] == null ? item.valorST : Convert.ToDecimal(xnIICMS10[0]["vICMSST"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS10[0]["orig"] == null ? item.origemCSTICMS : xnIICMS10[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS10[0]["CST"] == null ? item.CSTICMS : xnIICMS10[0]["CST"].InnerText;
                        }
                    }

                    //ICMS20
                    XmlNodeList xnIICMS20 = nodo.GetElementsByTagName("ICMS20");
                    {
                        if (xnIICMS20[0] != null)
                        {
                            item.percRedICMS = xnIICMS20[0]["pRedBC"] == null ? item.percRedICMS : Convert.ToDecimal(xnIICMS20[0]["pRedBC"].InnerText.Replace(".", ","));
                            item.baseICMS = xnIICMS20[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS20[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS20[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS20[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS20[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS20[0]["vICMS"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS20[0]["orig"] == null ? item.origemCSTICMS : xnIICMS20[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS20[0]["CST"] == null ? item.CSTICMS : xnIICMS20[0]["CST"].InnerText;
                        }
                    }

                    //ICMS30
                    XmlNodeList xnIICMS30 = nodo.GetElementsByTagName("ICMS30");
                    {
                        if (xnIICMS30[0] != null)
                        {
                            item.percRedST = xnIICMS30[0]["pRedBCST"] == null ? item.percRedST : Convert.ToDecimal(xnIICMS30[0]["pRedBCST"].InnerText.Replace(".", ","));
                            item.baseST = xnIICMS30[0]["vBCST"] == null ? item.baseST : Convert.ToDecimal(xnIICMS30[0]["vBCST"].InnerText.Replace(".", ","));
                            item.percST = xnIICMS30[0]["pICMSST"] == null ? item.percST : Convert.ToDecimal(xnIICMS30[0]["pICMSST"].InnerText.Replace(".", ","));
                            item.valorST = xnIICMS30[0]["vICMSST"] == null ? item.valorST : Convert.ToDecimal(xnIICMS30[0]["vICMSST"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS30[0]["orig"] == null ? item.origemCSTICMS : xnIICMS30[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS30[0]["CST"] == null ? item.CSTICMS : xnIICMS30[0]["CST"].InnerText;
                        }
                    }

                    //ICMS40
                    XmlNodeList xnIICMS40 = nodo.GetElementsByTagName("ICMS40");
                    {
                        if (xnIICMS40[0] != null)
                        {
                            item.valorICMS = xnIICMS40[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS40[0]["vICMS"].InnerText.Replace(".", ","));
                            item.origemCSTICMS = xnIICMS40[0]["orig"] == null ? item.origemCSTICMS : xnIICMS40[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS40[0]["CST"] == null ? item.CSTICMS : xnIICMS40[0]["CST"].InnerText;
                        }
                    }

                    //ICMS51
                    XmlNodeList xnIICMS51 = nodo.GetElementsByTagName("ICMS51");
                    {
                        if (xnIICMS51[0] != null)
                        {
                            item.percRedICMS = xnIICMS51[0]["pRedBC"] == null ? item.percRedICMS : Convert.ToDecimal(xnIICMS51[0]["pRedBC"].InnerText.Replace(".", ","));
                            item.baseICMS = xnIICMS51[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS51[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS51[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS51[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS51[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS51[0]["vICMS"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS51[0]["orig"] == null ? item.origemCSTICMS : xnIICMS51[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS51[0]["CST"] == null ? item.CSTICMS : xnIICMS51[0]["CST"].InnerText;
                        }
                    }

                    //ICMS60
                    XmlNodeList xnIICMS60 = nodo.GetElementsByTagName("ICMS60");
                    {
                        if (xnIICMS60[0] != null)
                        {
                            item.baseST = xnIICMS60[0]["vBCSTRet"] == null ? item.baseST : Convert.ToDecimal(xnIICMS60[0]["vBCSTRet"].InnerText.Replace(".", ","));
                            item.valorST = xnIICMS60[0]["vICMSSTRet"] == null ? item.valorST : Convert.ToDecimal(xnIICMS60[0]["vICMSSTRet"].InnerText.Replace(".", ","));
                            item.origemCSTICMS = xnIICMS60[0]["orig"] == null ? item.origemCSTICMS : xnIICMS60[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS60[0]["CST"] == null ? item.CSTICMS : xnIICMS60[0]["CST"].InnerText;
                        }
                    }

                    //ICMS70
                    XmlNodeList xnIICMS70 = nodo.GetElementsByTagName("ICMS70");
                    {
                        if (xnIICMS70[0] != null)
                        {
                            item.percRedICMS = xnIICMS70[0]["pRedBC"] == null ? item.percRedICMS : Convert.ToDecimal(xnIICMS70[0]["pRedBC"].InnerText.Replace(".", ","));
                            item.baseICMS = xnIICMS70[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS70[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS70[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS70[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS70[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS70[0]["vICMS"].InnerText.Replace(".", ","));

                            item.percRedST = xnIICMS70[0]["pRedBCST"] == null ? item.percRedST : Convert.ToDecimal(xnIICMS70[0]["pRedBCST"].InnerText.Replace(".", ","));
                            item.baseST = xnIICMS70[0]["vBCST"] == null ? item.baseST : Convert.ToDecimal(xnIICMS70[0]["vBCST"].InnerText.Replace(".", ","));
                            item.percST = xnIICMS70[0]["pICMSST"] == null ? item.percST : Convert.ToDecimal(xnIICMS70[0]["pICMSST"].InnerText.Replace(".", ","));
                            item.valorST = xnIICMS70[0]["vICMSST"] == null ? item.valorST : Convert.ToDecimal(xnIICMS70[0]["vICMSST"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS70[0]["orig"] == null ? item.origemCSTICMS : xnIICMS70[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS70[0]["CST"] == null ? item.CSTICMS : xnIICMS70[0]["CST"].InnerText;
                        }
                    }

                    //ICMS90
                    XmlNodeList xnIICMS90 = nodo.GetElementsByTagName("ICMS90");
                    {
                        if (xnIICMS90[0] != null)
                        {
                            item.percRedICMS = xnIICMS90[0]["pRedBC"] == null ? item.percRedICMS : Convert.ToDecimal(xnIICMS90[0]["pRedBC"].InnerText.Replace(".", ","));
                            item.baseICMS = xnIICMS90[0]["vBC"] == null ? item.baseICMS : Convert.ToDecimal(xnIICMS90[0]["vBC"].InnerText.Replace(".", ","));
                            item.percICMS = xnIICMS90[0]["pICMS"] == null ? item.percICMS : Convert.ToDecimal(xnIICMS90[0]["pICMS"].InnerText.Replace(".", ","));
                            item.valorICMS = xnIICMS90[0]["vICMS"] == null ? item.valorICMS : Convert.ToDecimal(xnIICMS90[0]["vICMS"].InnerText.Replace(".", ","));

                            item.percRedST = xnIICMS90[0]["pRedBCST"] == null ? item.percRedST : Convert.ToDecimal(xnIICMS90[0]["pRedBCST"].InnerText.Replace(".", ","));
                            item.baseST = xnIICMS90[0]["vBCST"] == null ? item.baseST : Convert.ToDecimal(xnIICMS90[0]["vBCST"].InnerText.Replace(".", ","));
                            item.percST = xnIICMS90[0]["pICMSST"] == null ? item.percST : Convert.ToDecimal(xnIICMS90[0]["pICMSST"].InnerText.Replace(".", ","));
                            item.valorST = xnIICMS90[0]["vICMSST"] == null ? item.valorST : Convert.ToDecimal(xnIICMS90[0]["vICMSST"].InnerText.Replace(".", ","));

                            item.origemCSTICMS = xnIICMS90[0]["orig"] == null ? item.origemCSTICMS : xnIICMS90[0]["orig"].InnerText;
                            item.CSTICMS = xnIICMS90[0]["CST"] == null ? item.CSTICMS : xnIICMS90[0]["CST"].InnerText;
                        }
                    }

                    //IPI
                    XmlNodeList xnIIPI = nodo.GetElementsByTagName("IPITrib");
                    if (xnIIPI[0] != null)
                    {
                        item.baseIPI = xnIIPI[0]["vBC"] == null ? item.baseIPI : Convert.ToDecimal(xnIIPI[0]["vBC"].InnerText.Replace(".", ","));
                        item.percIPI = xnIIPI[0]["pIPI"] == null ? item.percIPI : Convert.ToDecimal(xnIIPI[0]["pIPI"].InnerText.Replace(".", ","));
                        item.valorIPI = xnIIPI[0]["vIPI"] == null ? item.valorIPI : Convert.ToDecimal(xnIIPI[0]["vIPI"].InnerText.Replace(".", ","));
                        item.CSTIPI = xnIIPI[0]["CST"] == null ? item.CSTIPI : xnIIPI[0]["CST"].InnerText;
                    }

                    //PIS
                    XmlNodeList xnIPIS = nodo.GetElementsByTagName("PISAliq");
                    if (xnIPIS[0] != null)
                    {
                        item.basePIS = xnIPIS[0]["vBC"] == null ? item.basePIS : Convert.ToDecimal(xnIPIS[0]["vBC"].InnerText.Replace(".", ","));
                        item.percPIS = xnIPIS[0]["pPIS"] == null ? item.percPIS : Convert.ToDecimal(xnIPIS[0]["pPIS"].InnerText.Replace(".", ","));
                        item.valorPIS = xnIPIS[0]["vPIS"] == null ? item.valorPIS : Convert.ToDecimal(xnIPIS[0]["vPIS"].InnerText.Replace(".", ","));
                        item.CSTPIS = xnIPIS[0]["CST"] == null ? item.CSTPIS : xnIPIS[0]["CST"].InnerText;
                    }

                    //COFINS
                    XmlNodeList xnICofins = nodo.GetElementsByTagName("COFINSAliq");
                    if (xnICofins[0] != null)
                    {
                        item.baseCOFINS = xnICofins[0]["vBC"] == null ? item.baseCOFINS : Convert.ToDecimal(xnICofins[0]["vBC"].InnerText.Replace(".", ","));
                        item.percCOFINS = xnICofins[0]["pCOFINS"] == null ? item.percCOFINS : Convert.ToDecimal(xnICofins[0]["pCOFINS"].InnerText.Replace(".", ","));
                        item.valorCOFINS = xnICofins[0]["vCOFINS"] == null ? item.valorCOFINS : Convert.ToDecimal(xnICofins[0]["vCOFINS"].InnerText.Replace(".", ","));
                        item.CSTCOFINS = xnICofins[0]["CST"] == null ? item.CSTCOFINS : xnICofins[0]["CST"].InnerText;
                    }

                    //Informação Adicional Produto
                    XmlNodeList xnInfoAdicProd = nodo.GetElementsByTagName("infAdProd");
                    item.infoAdicionalProduto = xnInfoAdicProd[0] == null ? item.infoAdicionalProduto : xnInfoAdicProd[0].InnerText;

                    //Descrição Produto NF + Informação Adicional Produto
                    item.descricaoNFInfoAdicionalProduto = string.IsNullOrEmpty(item.infoAdicionalProduto) ? item.descricaoNF : item.descricaoNF + " " + item.infoAdicionalProduto;

                    lista.Add(item);
                }
            }

            return lista;
        }

        /// <summary>
        /// Carrega os dados das Duplicatas da Nota Fiscal a partir do arquivo XML gerado pela Receita
        /// Necessário informar o caminho de rede completo onde o arquivo XML está
        /// <param name="strCaminhoXMLNotaFiscal">Obrigatório - Caminhao físico completo do XML</param>
        /// <returns>Duplicatas da Nota Fiscal (ItensXML)</returns>
        public List<ParcelasXML> CarregarDadosXMLNotaFiscalParcelas(string strCaminhoXMLNotaFiscal)
        {
            List<ParcelasXML> lista = new List<ParcelasXML>();

            XmlDocument xmlDoc = new XmlDocument();
            XmlNamespaceManager xmlTag = new XmlNamespaceManager(xmlDoc.NameTable);
            xmlDoc.Load(strCaminhoXMLNotaFiscal);
            xmlTag.AddNamespace("nfe", xmlDoc.DocumentElement.NamespaceURI);

            XmlNodeList xnlCobr = xmlDoc.GetElementsByTagName("cobr");
            XmlNodeList xnlDup = xnlCobr[0] == null ? null : ((XmlElement)xnlCobr[0]).GetElementsByTagName("dup");

            // Verifica se veio o parcelamento no xml da nota
            if (xnlDup != null)
            {
                int sequenciaParcela = 0;
                foreach (XmlElement xnParc in xnlDup)
                {
                    //Incrementa a sequencia da Parcela
                    sequenciaParcela++;

                    //Cria Parcela
                    ParcelasXML parcela = new ParcelasXML();

                    //Sequencia da Parcela
                    parcela.ordemParcela = sequenciaParcela;

                    //Parcela
                    if (xnParc != null)
                    {
                        parcela.descricaoParcela = xnParc["nDup"] == null ? parcela.descricaoParcela : xnParc["nDup"].InnerText;

                        if (xnParc["dVenc"] != null)
                        {
                            string data = xnParc["dVenc"].InnerText;
                            DateTime.TryParse(data.Split('-')[2] + "/" + data.Split('-')[1] + "/" + data.Split('-')[0], out parcela.dataParcela);
                        }

                        parcela.valorParcela = xnParc["vDup"] == null ? parcela.valorParcela : Convert.ToDecimal(xnParc["vDup"].InnerText.Replace(".", ","));
                    }

                    lista.Add(parcela);
                }
            }

            return lista;
        }

        /// <summary>
        /// Carrega a lista de itens da nota fiscal a partir do DataTablePosition de produtos informada
        /// </summary>
        /// <param name="dtpProdutos">Obrigatório - DataTablePosition de itens da nota, devendo os campos possuírem o mesmo nome da classe ItensXML</param>
        /// <param name="dtpPedidoItens">Opcional - DataTablePosition de itens do pedido, devendo os campos possuírem o mesmo nome da classe ImpXMLNotaFiscalPedidoItem</param>
        /// <returns>Lista de Itens da Nota Fiscal (ItensXML)</returns>
        public List<ItensXML> CarregarDadosBancoDadosNotaFiscalItens(DataTable dtpProdutos)
        {
            List<ItensXML> lista = new List<ItensXML>();

            foreach (DataRow dr in dtpProdutos.Rows)
            {
                // Cria o item da nota fiscal
                ItensXML item = new ItensXML();
                
                // Popula os campos dos produtos da nota cadastrados
                item.sequenciaItem = dr.IsNull("sequenciaItem") ? item.sequenciaItem : Convert.ToInt32(dr["sequenciaItem"]);

                item.produtoNF = dr.IsNull("produtoNF") ? item.produtoNF : dr["produtoNF"].ToString();
                item.descricaoNF = dr.IsNull("descricaoNF") ? item.descricaoNF : dr["descricaoNF"].ToString();
                item.qtdeNF = dr.IsNull("qtdeNF") ? item.qtdeNF : Convert.ToDecimal(dr["qtdeNF"]);
                item.unidNF = dr.IsNull("unidNF") ? item.unidNF : dr["unidNF"].ToString();
                item.valorUnitario = dr.IsNull("valorUnitario") ? item.valorUnitario : Convert.ToDecimal(dr["valorUnitario"]);
                item.valorTotal = dr.IsNull("valorTotal") ? item.valorTotal : Convert.ToDecimal(dr["valorTotal"]);

                item.CFOP = dr.IsNull("CFOP") ? item.CFOP : dr["CFOP"].ToString();
                item.EAN = dr.IsNull("EAN") ? item.EAN : dr["EAN"].ToString();
                item.NCM = dr.IsNull("NCM") ? item.NCM : dr["NCM"].ToString();
                item.NCMFornec = dr.IsNull("NCMFORNEC") ? item.NCM : dr["NCMFORNEC"].ToString();

                item.percRedICMS = dr.IsNull("percRedICMS") ? item.percRedICMS : Convert.ToDecimal(dr["percRedICMS"]);
                item.baseICMS = dr.IsNull("baseICMS") ? item.baseICMS : Convert.ToDecimal(dr["baseICMS"]);
                item.percICMS = dr.IsNull("percICMS") ? item.percICMS : Convert.ToDecimal(dr["percICMS"]);
                item.valorICMS = dr.IsNull("valorICMS") ? item.valorICMS : Convert.ToDecimal(dr["valorICMS"]);

                item.percRedST = dr.IsNull("percRedST") ? item.percRedST : Convert.ToDecimal(dr["percRedST"]);
                item.baseST = dr.IsNull("baseST") ? item.baseST : Convert.ToDecimal(dr["baseST"]);
                item.percST = dr.IsNull("percST") ? item.percST : Convert.ToDecimal(dr["percST"]);
                item.valorST = dr.IsNull("valorST") ? item.valorST : Convert.ToDecimal(dr["valorST"]);

                item.origemCSTICMS = dr.IsNull("origemCSTICMS") ? item.origemCSTICMS : dr["origemCSTICMS"].ToString();
                item.CSTICMS = dr.IsNull("CSTICMS") ? item.CSTICMS : dr["CSTICMS"].ToString();

                item.baseIPI = dr.IsNull("baseIPI") ? item.baseIPI : Convert.ToDecimal(dr["baseIPI"]);
                item.percIPI = dr.IsNull("percIPI") ? item.percIPI : Convert.ToDecimal(dr["percIPI"]);
                item.valorIPI = dr.IsNull("valorIPI") ? item.valorIPI : Convert.ToDecimal(dr["valorIPI"]);

                item.basePIS = dr.IsNull("basePIS") ? item.basePIS : Convert.ToDecimal(dr["basePIS"]);
                item.percPIS = dr.IsNull("percPIS") ? item.percPIS : Convert.ToDecimal(dr["percPIS"]);
                item.valorPIS = dr.IsNull("valorPIS") ? item.valorPIS : Convert.ToDecimal(dr["valorPIS"]);

                item.baseCOFINS = dr.IsNull("baseCOFINS") ? item.baseCOFINS : Convert.ToDecimal(dr["baseCOFINS"]);
                item.percCOFINS = dr.IsNull("percCOFINS") ? item.percCOFINS : Convert.ToDecimal(dr["percCOFINS"]);
                item.valorCOFINS = dr.IsNull("valorCOFINS") ? item.valorCOFINS : Convert.ToDecimal(dr["valorCOFINS"]);

                item.valorDesconto = dr.IsNull("valorDesconto") ? item.valorDesconto : Convert.ToDecimal(dr["valorDesconto"]);

                item.descricaoNFInfoAdicionalProduto = item.descricaoNF;
                
                lista.Add(item);
            }

            return lista;
        }

        #endregion
    }
}
