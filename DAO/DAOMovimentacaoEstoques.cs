using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class DAOMovimentacaoEstoques
    {
    
        // string = VARCHAR
        // int = INTEGER
        // datetime = TIMESTAMP
        // decimal = DECIMAL

        public DateTime DataMovimento { get; set; }
        public string Dep { get; set; }
        public string Deposito { get; set; }
        public string Emp { get; set; }
        public string Empresa { get; set; }
        public string Cardex { get; set; }
        public string TipoVolume { get; set; }
        public string Qualidade { get; set; }
        public string Valorizacao { get; set; }
        public string Proc { get; set; }
        public string Processo { get; set; }
        public string Propriedade { get; set; }
        public string Terc { get; set; }
        public string Terceiro { get; set; }
        public string Tran { get; set; }
        public string Transacao { get; set; }
        public string Cat { get; set; }
        public string Agrupador { get; set; }
        public string Es { get; set; }
        public string TEntrada { get; set; }
        public string TCancela { get; set; }
        public string PrecoMedio { get; set; }
        public string AjusteFinan { get; set; }
        public string TipoConsumo { get; set; }
        public string TipoTransacao { get; set; }
        public string Cc { get; set; }
        public string CentroCusto { get; set; }
        public string Nivel { get; set; }
        public string Grupo { get; set; }
        public string Sub { get; set; }
        public string Cor { get; set; }
        public string Produto { get; set; }
        public string Um { get; set; }
        public string CodigoBarras { get; set; }
        public string CodigoVelho { get; set; }
        public string NomeGrupo { get; set; }
        public string NomeSub { get; set; }
        public string NomeCor { get; set; }
        public string Narrativa { get; set; }
        public string Cf { get; set; }
        public string Col { get; set; }
        public string Colecao { get; set; }
        public string Lin { get; set; }
        public string Linha { get; set; }
        public string Art { get; set; }
        public string Artigo { get; set; }
        public string Cota { get; set; }
        public string ArtigoCotas { get; set; }
        public string Ces { get; set; }
        public string ContaEstoque { get; set; }
        public string Tpg { get; set; }
        public string TipoProdutoGlobal { get; set; }
        public string TprogTpg { get; set; }
        public string NivTpg { get; set; }
        public string EstTpg { get; set; }
        public string DepTpg { get; set; }
        public string Cliente { get; set; }
        public string NomeCliente { get; set; }
        public string Marca { get; set; }
        public string NomeMarca { get; set; }
        public string TipoTecido { get; set; }
        public string Tpm { get; set; }
        public decimal Ncm { get; set; }
        public string AltP { get; set; }
        public string RotP { get; set; }
        public string AltC { get; set; }
        public string RotC { get; set; }
        public decimal ValorMedioEstoque { get; set; }
        public decimal ValorUltimaCompra { get; set; }
        public decimal Custo { get; set; }
        public decimal CustoInformado { get; set; }
        public string Lead { get; set; }
        public string FamiliaTear { get; set; }
        public string TipoProdQuimico { get; set; }
        public string Lote { get; set; }
        public string Documento { get; set; }
        public string Serie { get; set; }
        public string Projeto { get; set; }
        public string Seq { get; set; }
        public string Cnpj9 { get; set; }
        public string Cnpj4 { get; set; }
        public string Cnpj2 { get; set; }
        public string SeqFc { get; set; }
        public string SeqInc { get; set; }
        public decimal Qtde { get; set; }
        public decimal SaldoFisico { get; set; }
        public decimal SaldoFinanceiro { get; set; }
        public decimal ValorMovimentoUnitario { get; set; }
        public decimal ValorContabilUnitario { get; set; }
        public decimal ValorTotal { get; set; }
        public decimal PrecoMedioUnitario { get; set; }
        public string UsuarioSystextil { get; set; }
        public string ProcessoSystextil { get; set; }
        public DateTime DataInsercao { get; set; }
        public string TabelaOrigem { get; set; }

    }
}