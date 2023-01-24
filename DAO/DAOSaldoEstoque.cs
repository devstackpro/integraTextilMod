using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class DAOSaldoEstoque
    {
    
        // string = VARCHAR
        // int = INTEGER
        // datetime = TIMESTAMP
        // decimal = DECIMAL

        public string Emp { get; set; }
        public string Empresa { get; set; }
        public string Dep { get; set; }
        public string Deposito { get; set; }
        public int Nivel { get; set; }
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
        public string Lin { get; set; }
        public string Linha { get; set; }
        public string Art { get; set; }
        public string Cota { get; set; }
        public string ArtigoCotas { get; set; }
        public string Ces { get; set; }
        public string ContaEstoque { get; set; }
        public string Tpg { get; set; }
        public string TipoProdutoGlobal { get; set; }
        public string TprogTpg { get; set; }
        public string NivTpg { get; set; }
        public string EstTpg { get; set; }
        public string Cliente { get; set; }
        public string NomeCliente { get; set; }
        public string Marca { get; set; }
        public string NomeMarca { get; set; }
        public string TipoTecido { get; set; }
        public string Tpm { get; set; }
        public int Ncm { get; set; }
        public string AltP { get; set; }
        public string RotP { get; set; }
        public string AntC { get; set; }
        public string RotC { get; set; }
        public string ValorMedioEstoque { get; set; }
        public decimal ValorUltimaCopmpra { get; set; }
        public decimal Custo { get; set; }
        public string CustoInformado { get; set; }
        public string Lead { get; set; }
        public string FamiliaTear { get; set; }
        public string LoteTam { get; set; }
        public decimal PesoLiquido { get; set; }
        public decimal PesoRolo { get; set; }
        public decimal PesoMinRolo { get; set; }
        public string DescTamFicha { get; set; }
        public string TipoProdQuimico { get; set; }
        public string ItemAtivo { get; set; }
        public int CodigoContabil { get; set; }
        public string CodProcesso { get; set; }
        public string Lote { get; set; }
        public string LoteProduto { get; set; }
        public string SaldoAtual { get; set; }
        public string Volumes { get; set; }
        public string QtEstqInicioMes { get; set; }
        public string QtEstqFinalMes { get; set; }
        public datetime UltimaEntrada { get; set; }
        public datetime UltimaSaida { get; set; }
        public string QtSugerida { get; set; }
        public string QtEmpenhada { get; set; }
        public int CnpjFornecedor { get; set; }
        public int NotaFiscal { get; set; }
        public string PeriodoEstoque { get; set; }

    }
}