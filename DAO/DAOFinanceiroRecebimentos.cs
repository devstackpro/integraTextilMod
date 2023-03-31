using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class DAOFinanceiroRecebimentos
    {
        public string Empresa { get;set; }
        public string NumDuplicata { get;set; }
        public string Parcela { get;set; }
        public string TipoTitulo { get;set; }
        public string SituacaoDuplic { get;set; }
        public decimal Emissao { get;set; }
        public decimal VencOriginal { get;set; }
        public decimal Vencimento { get;set; }
        public string CodCliente { get;set; }
        public string Cliente { get;set; }
        public string CodResponsavel { get;set; }
        public string Responsavel { get;set; }
        public DateTime CodEndosso { get;set; }
        public DateTime Endosso { get;set; }
        public DateTime PedidoVenda { get;set; }
        public string Representante { get;set; }
        public decimal PercComissao { get;set; }
        public decimal BaseComissao { get;set; }
        public decimal ValorComissao { get;set; }
        public string Portador { get;set; }
        public string NumeroBordero { get;set; }
        public string NumeroRemessa { get;set; }
        public string NrTituloBanco { get;set; }
        public string ContaCorrente { get;set; }
        public string CodCarteira { get;set; }
        public string Transacao { get;set; }
        public decimal PercDesconto { get;set; }
        public string NrSolicitacao { get;set; }
        public decimal ValorDuplicata { get;set; }
        public DateTime SaldoDuplicata { get;set; }
        public decimal Moeda { get;set; }
        public string Posicao { get;set; }
        public string LocalEmpresa { get;set; }
        public string SituacaoDuplicata { get;set; }
        public string Historico { get;set; }
        public string ComplHistorico { get;set; }
        public string NumContabil { get;set; }
        public string FormaPagto { get;set; }
        public string CodBarras { get;set; }
        public string LinhaDigitavel { get;set; }
        public string DuplicImpressa { get;set; }
        public string Previsao { get;set; }
        public string NumeroTitulo { get;set; }
        public string NotaFiscal { get;set; }
        public string Serie { get;set; }
        public string CodCancelamento { get;set; }
        public string Cancelamento { get;set; }
        public string CodigoContabil { get;set; }
        public decimal ValorMoeda { get;set; }
        public string CodUsuario { get;set; }
        public string NumeroCaixa { get;set; }
        public string NrAdiantamento { get;set; }
        public string FantasiaCliente { get;set; }
        public string TelefoneCliente { get;set; }
        public string EmailCliente { get;set; }
        public string Radm { get;set; }
        public string Administrador { get;set; }
        public decimal ComissaoAdministr { get;set; }
        public string SeqRcbto { get;set; }
        public DateTime DataRcnto { get;set; }
        public DateTime DataCredito { get;set; }
        public decimal ValorRecebido { get;set; }
        public decimal ValorJuros { get;set; }
        public decimal ValorDesconto { get;set; }
        public string HisRcbto { get;set; }
        public string HistoricoRcbto { get;set; }
        public string NumeroDocumento { get;set; }
        public string DoctoRcbto { get;set; }
        public string PorRcbto { get;set; }
        public string PortadorRcbto { get;set; }
        public string ContaCorrenteRcbto { get;set; }
        public string NumContabilRcbto { get;set; }
        public decimal Atraso { get;set; }

    }
}