using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLComercialVendas
    {
        #region ATRIBUTOS | OBJETOS

        DALMySQL dalMySQL = new DALMySQL();

        string data = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");

        #endregion

        #region MÉTODOS

        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        public string PegarNomeArquivo(string pasta, string parteNomeArquivo)
        {
            string arquivoNome = "";

            BLLFerramentas bllFerramentas = new BLLFerramentas();

            try
            {
                // Take a snapshot of the file system.  
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(pasta);

                // This method assumes that the application has discovery permissions  
                // for all folders under the specified path.  
                IEnumerable<System.IO.FileInfo> fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);

                //Create the query  
                IEnumerable<System.IO.FileInfo> fileQuery =
                    from file in fileList
                        //where file.Extension == ".csv"
                    where file.Name.Contains(parteNomeArquivo)
                    orderby file.Name
                    select file;

                //Execute the query. This might write out a lot of files!  
                foreach (System.IO.FileInfo fi in fileQuery)
                {
                    arquivoNome = fi.Name;
                }

                if (arquivoNome.Equals(""))
                {
                    Convert.ToInt32(arquivoNome);
                }
                else
                {
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de comercial vendas encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de comercial vendas não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
            }

            return arquivoNome;

        }

        public string RenomearArquivo(string arquivoNome, string arquivoNomeFinal, string pastaOrigem, string pastaDestino)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";

            try
            {
                if (arquivoNome.Equals(""))
                {
                    Convert.ToInt32(arquivoNome);
                    retorno = "Arquivo não foi encontrado na pasta Origem";
                }
                else
                {
                    string[] arquivos = Directory.GetFiles(pastaOrigem);
                    string dirSaida = pastaDestino;

                    if (!Directory.Exists(dirSaida))
                        Directory.CreateDirectory(dirSaida);

                    for (int i = 0; i < arquivos.Length; i++)
                    {

                        if (arquivos[i].Equals(pastaOrigem + arquivoNome))
                        {
                            var files = new FileInfo(arquivos[i]);
                            files.MoveTo(Path.Combine(dirSaida, files.Name.Replace(arquivoNome, arquivoNomeFinal)));
                        }

                    }

                    retorno = "ok";
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: comercial vendas e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear comercial vendas. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOComercialVendasList LerCsv(string path)
        {
            DAOComercialVendasList daoComercialVendasList = new DAOComercialVendasList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOComercialVendas daoComercialVendas = new DAOComercialVendas();   
                campo = linha.Split(';');
                index++;

                if (index > 1)
                {
                    daoComercialVendas.Pedido = campo[0].ToString();
                    daoComercialVendas.PedidoCliente = campo[1].ToString();
                    daoComercialVendas.NumeroInterno = campo[2].ToString();
                    daoComercialVendas.Emp = campo[3].ToString();
                    daoComercialVendas.Empresa = campo[4].ToString();
                    daoComercialVendas.TipoPedido = campo[5].ToString();
                    daoComercialVendas.TipoProduto = campo[6].ToString();
                    daoComercialVendas.CriterioPedido = campo[7].ToString();
                    daoComercialVendas.CriterioQualidade = campo[8].ToString();
                    daoComercialVendas.Situacao = campo[9].ToString();
                    daoComercialVendas.BloqueiosPendentes = campo[10].ToString();
                    daoComercialVendas.BloqueiosLiberados = campo[11].ToString();
                    daoComercialVendas.Emissao = campo[12].ToString();
                    daoComercialVendas.Entrega = campo[13].ToString();
                    daoComercialVendas.Chegada = campo[14].ToString();
                    daoComercialVendas.Periodo = campo[15].ToString();
                    daoComercialVendas.Uf = campo[16].ToString();
                    daoComercialVendas.Cnpj = campo[17].ToString();
                    daoComercialVendas.NomeCliente = campo[18].ToString();
                    daoComercialVendas.Fantasia = campo[19].ToString();
                    daoComercialVendas.Geco = campo[20].ToString();
                    daoComercialVendas.GrupoEconomico = campo[21].ToString();
                    daoComercialVendas.DataValidadeLimite = campo[22].ToString();
                    daoComercialVendas.DataAlteracaoLimite = campo[23].ToString();
                    daoComercialVendas.ObsCredito = campo[24].ToString();
                    daoComercialVendas.LimCreditoConfeccao = campo[25].ToString();
                    daoComercialVendas.LimCreditoTecidos = campo[26].ToString();
                    daoComercialVendas.LimCreditoCrus = campo[27].ToString();
                    daoComercialVendas.LimCreditoFios = campo[28].ToString();
                    daoComercialVendas.DataUltimaFatura = campo[29].ToString();    
                    daoComercialVendas.PedidosAFaturar = Convert.ToDecimal[30].ToString(); 
                    daoComercialVendas.TitulosVencidos = Convert.ToDecimal[31].ToString();
                    daoComercialVendas.TitulosAVenc = Convert.ToDecimal[32].ToString();
                    daoComercialVendas.TitulosPagos = Convert.ToDecimal[33].ToString();
                    daoComercialVendas.Rep = campo[34].ToString();
                    daoComercialVendas.NomeRepresenante = campo[35].ToString();
                    daoComercialVendas.Regiao = campo[36].ToString();
                    daoComercialVendas.NomeRegiao = campo[37].ToString();
                    daoComercialVendas.Bco = campo[38].ToString();
                    daoComercialVendas.Banco = campo[39].ToString();
                    daoComercialVendas.Cid = campo[40].ToString();
                    daoComercialVendas.Cidade = campo[41].ToString();
                    daoComercialVendas.TipoComissao = campo[42].ToString();
                    daoComercialVendas.Comissao = campo[43].ToString();
                    daoComercialVendas.Class = campo[44].ToString();
                    daoComercialVendas.Classificacao = campo[45].ToString();
                    daoComercialVendas.Pgto = campo[46].ToString();
                    daoComercialVendas.CondicaoPgto = campo[47].ToString();
                    daoComercialVendas.TabelaPreco = campo[48].ToString();
                    daoComercialVendas.TipoDesconto = campo[49].ToString();
                    daoComercialVendas.Desconto1 = campo[50].ToString();
                    daoComercialVendas.Desconto2 = campo[51].ToString();
                    daoComercialVendas.Desconto3 = campo[52].ToString();
                    daoComercialVendas.DescontoItem1 = campo[53].ToString();
                    daoComercialVendas.DescontoItem2 = campo[54].ToString();
                    daoComercialVendas.DescontoItem3 = campo[55].ToString();
                    daoComercialVendas.DescontoEspecial = campo[56].ToString();
                    daoComercialVendas.DescontoExtra = campo[57].ToString();
                    daoComercialVendas.DescFinanceiro = campo[58].ToString();
                    daoComercialVendas.DescDuplicatas = campo[59].ToString();
                    daoComercialVendas.TipoFrete = campo[60].ToString();
                    daoComercialVendas.CnpjTrans = campo[61].ToString();
                    daoComercialVendas.Transportadora = campo[62].ToString();
                    daoComercialVendas.TipoRedespacho = campo[63].ToString();
                    daoComercialVendas.CnpjRedesp = campo[64].ToString();
                    daoComercialVendas.Redespacho = campo[65].ToString();
                    daoComercialVendas.ValorTotalPedido = campo[66].ToString();
                    daoComercialVendas.ValorSaldoPedido = campo[67].ToString();
                    daoComercialVendas.ValorItens = campo[68].ToString();
                    daoComercialVendas.ValorFrete = campo[69].ToString();
                    daoComercialVendas.Col = campo[70].ToString();
                    daoComercialVendas.Colecao = campo[71].ToString();
                    daoComercialVendas.Lin = campo[72].ToString();
                    daoComercialVendas.Linha = campo[73].ToString();
                    daoComercialVendas.Art = campo[74].ToString();
                    daoComercialVendas.Artigo = campo[75].ToString();
                    daoComercialVendas.Grupo = campo[76].ToString();
                    daoComercialVendas.Sub = campo[77].ToString();
                    daoComercialVendas.Cor = campo[78].ToString();
                    daoComercialVendas.Produto = campo[79].ToString();
                    daoComercialVendas.NomeGrupo = campo[80].ToString();
                    daoComercialVendas.NomeSUB = campo[81].ToString();
                    daoComercialVendas.NomeCor = campo[82].ToString();
                    daoComercialVendas.Narrativa = campo[83].ToString();
                    daoComercialVendas.Lote = campo[84].ToString();
                    daoComercialVendas.Embalagem = campo[85].ToString();
                    daoComercialVendas.Dep = campo[86].ToString();
                    daoComercialVendas.Deposito = campo[87].ToString();
                    daoComercialVendas.Vendido = Convert.ToDecimal[88].ToString();
                    daoComercialVendas.Faturado = Convert.ToDecimal[89].ToString();
                    daoComercialVendas.Solicitado = Convert.ToDecimal[90].ToString();
                    daoComercialVendas.Cancelado = Convert.ToDecimal[92].ToString();
                    daoComercialVendas.Saldo = Convert.ToDecimal[91].ToString();
                    daoComercialVendas.SitItem = campo[92].ToString();
                    daoComercialVendas.Alocado = campo[93].ToString();
                    daoComercialVendas.CodCanc = campo[94].ToString();
                    daoComercialVendas.DescontoITEM = campo[95].ToString();
                    daoComercialVendas.Unitario = campo[96].ToString();


                    daoComercialVendasList.Add(daoComercialVendas);
                    
                }
            }
            return daoComercialVendasList;
        }

        public string InserirDadosBD(DAOComercialVendasList daoComercialVendasList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspComercialVendasDeletar");

            try
            {
                DataTable dataTableComercialVendasList = ConvertToDataTable(daoComercialVendasList);
                foreach (DataRow linha in dataTableComercialVendasList.Rows)
                {
                    DAOComercialVendas daoComercialVendas = new DAOComercialVendas();

                    daoComercialVendas.Pedido = linha["Pedido"].ToString();
                    daoComercialVendas.PedidoCliente = linha["PedidoCliente"].ToString();
                    daoComercialVendas.NumeroInterno = linha["NumeroInterno"].ToString();
                    daoComercialVendas.Emp = linha["Emp"].ToString();
                    daoComercialVendas.Empresa = linha["Empresa"].ToString();
                    daoComercialVendas.TipoPedido = linha["TipoPedido"].ToString();
                    daoComercialVendas.TipoProduto = linha["TipoProduto"].ToString();
                    daoComercialVendas.CriterioPedido = linha["CriterioPedido"].ToString();
                    daoComercialVendas.CriterioQualidade = linha["CriterioQualidade"].ToString();
                    daoComercialVendas.Situacao = linha["Situacao"].ToString();
                    daoComercialVendas.BloqueiosPendentes = linha["BloqueiosPendentes"].ToString(); 
                    daoComercialVendas.BloqueiosLiberados = linha["BloqueiosLiberados"].ToString(); 
                    daoComercialVendas.Emissao = linha["Emissao"].ToString(); 
                    daoComercialVendas.Entrega = linha["Entrega"].ToString(); 
                    daoComercialVendas.Chegada = linha["Chegada"].ToString(); 
                    daoComercialVendas.Periodo = linha["Periodo"].ToString(); 
                    daoComercialVendas.Uf = linha["Uf"].ToString(); 
                    daoComercialVendas.Cnpj = linha["Cnpj"].ToString(); 
                    daoComercialVendas.NomeCliente = linha["NomeCliente"].ToString(); 
                    daoComercialVendas.Fantasia = linha["Fantasia"].ToString(); 
                    daoComercialVendas.Geco = linha["Geco"].ToString(); 
                    daoComercialVendas.GrupoEconomico = linha["GrupoEconomico"].ToString(); 
                    daoComercialVendas.DataValidadeLimite = linha["DataValidadeLimite"].ToString(); 
                    daoComercialVendas.DataAlteracaoLimite = linha["DataAlteracaoLimite"].ToString(); 
                    daoComercialVendas.ObsCredito = linha["ObsCredito"].ToString(); 
                    daoComercialVendas.LimCreditoConfeccao = linha["LimCreditoConfeccao"].ToString(); 
                    daoComercialVendas.LimCreditoTecidos = linha["LimCreditoTecidos"].ToString(); 
                    daoComercialVendas.LimCreditoCrus = linha["LimCreditoCrus"].ToString(); 
                    daoComercialVendas.LimCreditoFios = linha["LimCreditoFios"].ToString(); 
                    daoComercialVendas.DataUltimaFatura = linha["DataUltimaFatura+"].ToString();     
                    daoComercialVendas.PedidosAFaturar = Convert.ToDecimal(linha["PedidosAFaturar"].ToString()); 
                    daoComercialVendas.TitulosVencidos = Convert.ToDecimal(linha["TitulosVencidos"].ToString()); 
                    daoComercialVendas.TitulosAVenc = Convert.ToDecimal(linha["TitulosAVenc"].ToString()); 
                    daoComercialVendas.TitulosPagos = Convert.ToDecimal(linha["TitulosPagos"].ToString()); 
                    daoComercialVendas.Rep = linha["Rep"].ToString(); 
                    daoComercialVendas.NomeRepresenante = linha["NomeRepresenante"].ToString(); 
                    daoComercialVendas.Regiao = linha["Regiao"].ToString(); 
                    daoComercialVendas.NomeRegiao = linha["NomeRegiao"].ToString(); 
                    daoComercialVendas.Bco = linha["Bco"].ToString();  
                    daoComercialVendas.Banco = linha["Banco"].ToString(); 
                    daoComercialVendas.Cid = linha["Cid"].ToString(); 
                    daoComercialVendas.Cidade = linha["Cidade"].ToString(); 
                    daoComercialVendas.TipoComissao = linha["TipoComissao"].ToString(); 
                    daoComercialVendas.Comissao = linha["Comissao"].ToString(); 
                    daoComercialVendas.Class = linha["Class"].ToString(); 
                    daoComercialVendas.Classificacao = linha["Classificacao"].ToString(); 
                    daoComercialVendas.Pgto = linha["Pgto"].ToString(); 
                    daoComercialVendas.CondicaoPgto = linha["CondicaoPgto"].ToString(); 
                    daoComercialVendas.TabelaPreco = linha["TabelaPreco"].ToString(); 
                    daoComercialVendas.TipoDesconto = linha["TipoDesconto"].ToString(); 
                    daoComercialVendas.Desconto1 = linha["Desconto1"].ToString(); 
                    daoComercialVendas.Desconto2 = linha["Desconto2"].ToString(); 
                    daoComercialVendas.Desconto3 = linha["Desconto3"].ToString(); 
                    daoComercialVendas.DescontoItem1 = linha["DescontoItem1"].ToString(); 
                    daoComercialVendas.DescontoItem2 = linha["DescontoItem2"].ToString(); 
                    daoComercialVendas.DescontoItem3 = linha["DescontoItem3"].ToString(); 
                    daoComercialVendas.DescontoEspecial = linha["DescontoEspecial"].ToString(); 
                    daoComercialVendas.DescontoExtra = linha["DescontoExtra"].ToString(); 
                    daoComercialVendas.DescFinanceiro = linha["DescFinanceiro"].ToString(); 
                    daoComercialVendas.DescDuplicatas = linha["DescDuplicatas"].ToString(); 
                    daoComercialVendas.TipoFrete = linha["TipoFrete"].ToString(); 
                    daoComercialVendas.CnpjTrans = linha["CnpjTrans"].ToString(); 
                    daoComercialVendas.Transportadora = linha["Transportadora"].ToString(); 
                    daoComercialVendas.TipoRedespacho = linha["TipoRedespacho"].ToString(); 
                    daoComercialVendas.CnpjRedesp = linha["CnpjRedesp"].ToString(); 
                    daoComercialVendas.Redespacho = linha["Redespacho"].ToString(); 
                    daoComercialVendas.ValorTotalPedido = linha["ValorTotalPedido"].ToString(); 
                    daoComercialVendas.ValorSaldoPedido = linha["ValorSaldoPedido"].ToString(); 
                    daoComercialVendas.ValorItens = linha["ValorItens"].ToString(); 
                    daoComercialVendas.ValorFrete = linha["ValorFrete"].ToString(); 
                    daoComercialVendas.Col = linha["Col"].ToString(); 
                    daoComercialVendas.Colecao = linha["Colecao"].ToString(); 
                    daoComercialVendas.Lin = linha["Lin"].ToString(); 
                    daoComercialVendas.Linha = linha["Linha"].ToString(); 
                    daoComercialVendas.Art = linha["Art"].ToString(); 
                    daoComercialVendas.Artigo = linha["Artigo"].ToString(); 
                    daoComercialVendas.Grupo = linha["Grupo"].ToString(); 
                    daoComercialVendas.Sub = linha["Sub"].ToString(); 
                    daoComercialVendas.Cor = linha["Cor"].ToString(); 
                    daoComercialVendas.Produto = linha["Produto"].ToString(); 
                    daoComercialVendas.NomeGrupo = linha["NomeGrupo"].ToString(); 
                    daoComercialVendas.NomeSUB = linha["NomeSUB"].ToString(); 
                    daoComercialVendas.NomeCor = linha["NomeCor"].ToString(); 
                    daoComercialVendas.Narrativa = linha["Narrativa"].ToString(); 
                    daoComercialVendas.Lote = linha["Lote"].ToString(); 
                    daoComercialVendas.Embalagem = linha["Embalagem"].ToString(); 
                    daoComercialVendas.Dep = linha["Dep"].ToString(); 
                    daoComercialVendas.Deposito = linha["Deposito"].ToString(); 
                    daoComercialVendas.Vendido = Convert.ToDecimal(linha["Vendido"].ToString()); 
                    daoComercialVendas.Faturado = Convert.ToDecimal(linha["Faturado"].ToString()); 
                    daoComercialVendas.Solicitado = Convert.ToDecimal(linha["Solicitado"].ToString()); 
                    daoComercialVendas.Cancelado = Convert.ToDecimal(linha["Cancelado"].ToString()); 
                    daoComercialVendas.Saldo = linConvert.ToDecimal(linhaha["Saldo"].ToString()); 
                    daoComercialVendas.SitItem = linha["SitItem"].ToString(); 
                    daoComercialVendas.Alocado = Convert.ToDecimal(linha["Alocado"].ToString()); 
                    daoComercialVendas.CodCanc = linha["CodCanc"].ToString(); 
                    daoComercialVendas.DescontoITEM = linha["DescontoITEM"].ToString(); 
                    daoComercialVendas.Unitario = Convert.ToDecimal(linha["Unitario"].ToString()); 

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParamet("@Pedido", daoComercialVendas.Pedido);                                                       
                     dalMySQL.AdicionaParamet("@PedidoCliente", daoComercialVendas.PedidoCliente);                                                       
                     dalMySQL.AdicionaParamet("@NumeroInterno", daoComercialVendas.NumeroInterno);                                                       
                     dalMySQL.AdicionaParamet("@Emp", daoComercialVendas.Emp);                                                       
                     dalMySQL.AdicionaParamet("@Empresa", daoComercialVendas.Empresa);                                                       
                     dalMySQL.AdicionaParamet("@TipoPedido", daoComercialVendas.TipoPedido);                                                       
                     dalMySQL.AdicionaParamet("@TipoProduto", daoComercialVendas.TipoProduto);                                                      
                     dalMySQL.AdicionaParamet("@CriterioPedido", daoComercialVendas.CriterioPedido);                                                      
                     dalMySQL.AdicionaParamet("@CriterioQualidade", daoComercialVendas.CriterioQualidade);                                                       
                     dalMySQL.AdicionaParamet("@Situacao", daoComercialVendas.Situacao);   
                     dalMySQL.AdicionaParamet("@BloqueiosPendentes", daoComercialVendas.BloqueiosPendentes);                                                        
                     dalMySQL.AdicionaParamet("@BloqueiosLiberados", daoComercialVendas.BloqueiosLiberados);                                                        
                     dalMySQL.AdicionaParamet("@Emissao", daoComercialVendas.Emissao);                                                       
                     dalMySQL.AdicionaParamet("@Chegada", daoComercialVendas.Chegada);                                                       
                     dalMySQL.AdicionaParamet("@Periodo", daoComercialVendas.Periodo);                                                       
                     dalMySQL.AdicionaParamet("@Uf", daoComercialVendas.Uf);                                                       
                     dalMySQL.AdicionaParamet("@Cnpj", daoComercialVendas.Cnpj);                                                       
                     dalMySQL.AdicionaParamet("@NomeCliente", daoComercialVendas.NomeCliente);                                                      
                     dalMySQL.AdicionaParamet("@Fantasia", daoComercialVendas.Fantasia);                                                                                      
                     dalMySQL.AdicionaParamet("@Geco", daoComercialVendas.Geco);                                                                                      
                     dalMySQL.AdicionaParamet("@GrupoEconomico", daoComercialVendas.GrupoEconomico);                                                                                     
                     dalMySQL.AdicionaParamet("@DataValidadeLimite", daoComercialVendas.DataValidadeLimite);                                                                                    
                     dalMySQL.AdicionaParamet("@ObsCredito", daoComercialVendas.ObsCredito);                                                      
                     dalMySQL.AdicionaParamet("@DataAlteracaoLimite", daoComercialVendas.DataAlteracaoLimite);                                                                                   
                     dalMySQL.AdicionaParamet("@LimCreditoConfeccao", daoComercialVendas.LimCreditoConfeccao);                                                                                      
                     dalMySQL.AdicionaParamet("@LimCreditoTecidos", daoComercialVendas.LimCreditoTecidos);                                                                                     
                     dalMySQL.AdicionaParamet("@LimCreditoCrus", daoComercialVendas.LimCreditoCrus);                                                                                     
                     dalMySQL.AdicionaParamet("@LimCreditoFios", daoComercialVendas.LimCreditoFios);                                                                                     
                     dalMySQL.AdicionaParamet("@DataUltimaFatura", daoComercialVendas.DataUltimaFatura);                                                                                       
                     dalMySQL.AdicionaParamet("@PedidosAFaturar", daoComercialVendas.PedidosAFaturar);                                                                                 
                     dalMySQL.AdicionaParamet("@TitulosVencidos", daoComercialVendas.TitulosVencidos);                                                                                 
                     dalMySQL.AdicionaParamet("@TitulosAVenc", daoComercialVendas.TitulosAVenc);                                                                              
                     dalMySQL.AdicionaParamet("@TitulosPagos", daoComercialVendas.TitulosPagos);                                                                                                           
                     dalMySQL.AdicionaParamet("@Rep", daoComercialVendas.Rep);                                                      
                     dalMySQL.AdicionaParamet("@NomeRepresenante", daoComercialVendas.NomeRepresenante);                                                                                     
                     dalMySQL.AdicionaParamet("@Regiao", daoComercialVendas.Regiao);                                                                                      
                     dalMySQL.AdicionaParamet("@NomeRegiao", daoComercialVendas.NomeRegiao);                                                                                      
                     dalMySQL.AdicionaParamet("@Bco", daoComercialVendas.Bco);                                                                                     
                     dalMySQL.AdicionaParamet("@Banco", daoComercialVendas.Banco);                                                                                     
                     dalMySQL.AdicionaParamet("@Cid", daoComercialVendas.Cid);                                                                                      
                     dalMySQL.AdicionaParamet("@Cidade", daoComercialVendas.Cidade);                                                                                     
                     dalMySQL.AdicionaParamet("@TipoComissao", daoComercialVendas.TipoComissao);                                                                                      
                     dalMySQL.AdicionaParamet("@Comissao", daoComercialVendas.Comissao);                                                                                     
                     dalMySQL.AdicionaParamet("@Class", daoComercialVendas.Class);                                                                                     
                     dalMySQL.AdicionaParamet("@Classificacao", daoComercialVendas.Classificacao);                                                                                   
                     dalMySQL.AdicionaParamet("@Pgto", daoComercialVendas.Pgto);                                                                                     
                     dalMySQL.AdicionaParamet("@CondicaoPgto", daoComercialVendas.CondicaoPgto);                                                                                   
                     dalMySQL.AdicionaParamet("@TabelaPreco", daoComercialVendas.TabelaPreco);                                                                                  
                     dalMySQL.AdicionaParamet("@TipoDesconto", daoComercialVendas.TipoDesconto);                                                                                     
                     dalMySQL.AdicionaParamet("@Desconto2", daoComercialVendas.Desconto2);                                                                                       
                     dalMySQL.AdicionaParamet("@Desconto3", daoComercialVendas.Desconto3);                                                                                     
                     dalMySQL.AdicionaParamet("@DescontoItem1", daoComercialVendas.DescontoItem1);                                                                                   
                     dalMySQL.AdicionaParamet("@DescontoItem2", daoComercialVendas.DescontoItem2);                                                                                     
                     dalMySQL.AdicionaParamet("@DescontoItem3", daoComercialVendas.DescontoItem3);                                                                                   
                     dalMySQL.AdicionaParamet("@DescontoEspecial", daoComercialVendas.DescontoEspecial);                                                                                     
                     dalMySQL.AdicionaParamet("@DescontoExtra", daoComercialVendas.DescontoExtra);                                                                                      
                     dalMySQL.AdicionaParamet("@DescFinanceiro", daoComercialVendas.DescFinanceiro);                                                                                      
                     dalMySQL.AdicionaParamet("@DescDuplicatas", daoComercialVendas.DescDuplicatas);                                                                                      
                     dalMySQL.AdicionaParamet("@TipoFrete", daoComercialVendas.TipoFrete);                                                                                    
                     dalMySQL.AdicionaParamet("@CnpjTrans", daoComercialVendas.CnpjTrans);                                                                                     
                     dalMySQL.AdicionaParamet("@Transportadora", daoComercialVendas.Transportadora);                                                                                   
                     dalMySQL.AdicionaParamet("@TipoRedespacho", daoComercialVendas.TipoRedespacho);                                                                                     
                     dalMySQL.AdicionaParamet("@CnpjRedesp", daoComercialVendas.CnpjRedesp);                                                                                      
                     dalMySQL.AdicionaParamet("@Redespacho", daoComercialVendas.Redespacho);                                                                                     
                     dalMySQL.AdicionaParamet("@ValorTotalPedido", daoComercialVendas.ValorTotalPedido);                                                                                     
                     dalMySQL.AdicionaParamet("@ValorSaldoPedido", daoComercialVendas.ValorSaldoPedido);                                                                                      
                     dalMySQL.AdicionaParamet("@ValorItens", daoComercialVendas.ValorItens);                                                                                      
                     dalMySQL.AdicionaParamet("@ValorFrete", daoComercialVendas.ValorFrete);                                                                                      
                     dalMySQL.AdicionaParamet("@Col", daoComercialVendas.Col);                                                                                      
                     dalMySQL.AdicionaParamet("@Colecao", daoComercialVendas.Colecao);                                                                                      
                     dalMySQL.AdicionaParamet("@Lin", daoComercialVendas.Lin);                                                                                     
                     dalMySQL.AdicionaParamet("@Linha", daoComercialVendas.Linha);                                                                                      
                     dalMySQL.AdicionaParamet("@Art", daoComercialVendas.Art);                                                                                      
                     dalMySQL.AdicionaParamet("@Artigo", daoComercialVendas.Artigo);                                                                                      
                     dalMySQL.AdicionaParamet("@Grupo", daoComercialVendas.Grupo);                                                                                      
                     dalMySQL.AdicionaParamet("@Sub", daoComercialVendas.Sub);                                                                                    
                     dalMySQL.AdicionaParamet("@Cor", daoComercialVendas.Cor);                                                                                      
                     dalMySQL.AdicionaParamet("@Produto", daoComercialVendas.Produto);                                                                                      
                     dalMySQL.AdicionaParamet("@NomeGrupo", daoComercialVendas.NomeGrupo);                                                                                      
                     dalMySQL.AdicionaParamet("@NomeSUB", daoComercialVendas.NomeSUB);                                                                                      
                     dalMySQL.AdicionaParamet("@NomeCor", daoComercialVendas.NomeCor);                                                                                      
                     dalMySQL.AdicionaParamet("@Narrativa", daoComercialVendas.Narrativa);                                                                                       
                     dalMySQL.AdicionaParamet("@Lote", daoComercialVendas.Lote);                                                                                       
                     dalMySQL.AdicionaParamet("@Embalagem", daoComercialVendas.Embalagem);                                                                                      
                     dalMySQL.AdicionaParamet("@Dep", daoComercialVendas.Dep);                                                                                       
                     dalMySQL.AdicionaParamet("@Deposito", daoComercialVendas.Deposito);                                                                                      
                     dalMySQL.AdicionaParamet("@Vendido", daoComercialVendas.Vendido);                                                                         
                     dalMySQL.AdicionaParamet("@Faturado", daoComercialVendas.Faturado);                                                                          
                     dalMySQL.AdicionaParamet("@Solicitado", daoComercialVendas.Solicitado);                                                                            
                     dalMySQL.AdicionaParamet("@Cancelado", daoComercialVendas.Cancelado);                                                                           
                     dalMySQL.AdicionaParamet("@Saldo", daoComercialVendas.Saldo);                                                       
                     dalMySQL.AdicionaParamet("@SitItem", daoComercialVendas.SitItem);                                                                                      
                     dalMySQL.AdicionaParamet("@Alocado", daoComercialVendas.Alocado);                                                                         
                     dalMySQL.AdicionaParamet("@CodCanc", daoComercialVendas.CodCanc);                                                                                      
                     dalMySQL.AdicionaParamet("@DescontoITEM", daoComercialVendas.DescontoITEM);                                                                                      
                     dalMySQL.AdicionaParamet("@Unitario", daoComercialVendas.Unitario);   

                    
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: comercial vendas inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir comercial vendas. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        public string DeletarArquivos(string path)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            try
            {
                //System.IO.Directory.Delete(path, true);
                System.IO.DirectoryInfo di = new DirectoryInfo(path);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo dir in di.GetDirectories())
                {
                    dir.Delete(true);
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de comercial vendas deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de comercial vendas renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
