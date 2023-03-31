using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class DAOFinanceiroAPagar
    {
        public string Empresa { get;set; }
        public string Duplicata { get;set; }
        public string Parcela { get;set; }
        public DateTime DataContrato { get;set; }
        public string TipoTitulo { get;set; }
        public string Documento { get;set; }
        public string Serie { get;set; }
        public string Historico { get;set; }
        public string EmpresaCobranca { get;set; }
        public string CodContabil { get;set; }
        public string CodFornecedor { get;set; }
        public string NomeFornecedor { get;set; }
        public DateTime TipoFornecedor { get;set; }
        public string Transacao { get;set; }
        public DateTime DataTransacao { get;set; }
        public string Previsao { get;set; }
        public string Portador { get;set; }
        public DateTime VencimentoOrig { get;set; }
        public DateTime Vencimento { get;set; }
        public string Posicao { get;set; }
        public string CentroCusto { get;set; }
        public string NumContabil { get;set; }
        public string OrigemDebito { get;set; }
        public string SituacaoTitulo { get;set; }
        public string SituacaoSispag { get;set; }
        public string TipoPagamento { get;set; }
        public string CodigoBarras { get;set; }
        public decimal Moeda { get;set; }
        public decimal ValorTitulo { get;set; }
        public decimal SaldoTitulo { get;set; }


    }
}
