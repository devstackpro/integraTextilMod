using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL 
{
    public class BLLFinanceiroAPagar
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de financeiro contas a pagar encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContaspagar.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de financeiro contas a pagar não encontrado, nome: nulo. Detalhes: Classe: BLLContaspagar.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: financeiro contas a pagar movido e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear financeiro contas a pagar. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOFinanceiroAPagarList LerCsv(string path)
        {
            DAOFinanceiroAPagarList daoFinanceiroAPagarList = new DAOFinanceiroAPagarList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOFinanceiroAPagar daoFinanceiroAPagar = new DAOFinanceiroAPagar();
                campo = linha.Split(";");
                index++;

                if (index > 1)
                {
                    daoFinanceiroAPagar.Empresa  = campo[0].toString();
                    daoFinanceiroAPagar.Duplicata  = campo[1].toString();
                    daoFinanceiroAPagar.Parcela  = campo[2].toString();
                    daoFinanceiroAPagar.DataContrato  = campo[3].toString();
                    daoFinanceiroAPagar.TipoTitulo  = campo[4].toString();
                    daoFinanceiroAPagar.Documento  = campo[5].toString();
                    daoFinanceiroAPagar.Serie  = campo[6].toString();
                    daoFinanceiroAPagar.Historico  = campo[7].toString();
                    daoFinanceiroAPagar.EmpresaCobranca  = campo[8].toString();
                    daoFinanceiroAPagar.CodContabil  = campo[9].toString();
                    daoFinanceiroAPagar.CodFornecedor  = campo[10].toString();
                    daoFinanceiroAPagar.NomeFornecedor  = campo[11].toString();
                    daoFinanceiroAPagar.TipoFornecedor  = campo[12].toString();
                    daoFinanceiroAPagar.Transacao  = campo[13].toString();
                    daoFinanceiroAPagar.DataTransacao  = campo[14].toString();
                    daoFinanceiroAPagar.Previsao  = campo[15].toString();
                    daoFinanceiroAPagar.Portador  = campo[16].toString();
                    daoFinanceiroAPagar.VencimentoOrig  = campo[17].toString();
                    daoFinanceiroAPagar.Vencimento  = campo[18].toString();
                    daoFinanceiroAPagar.Posicao  = campo[19].toString();
                    daoFinanceiroAPagar.CentroCusto  = campo[20].toString();
                    daoFinanceiroAPagar.NumContabil  = campo[21].toString();
                    daoFinanceiroAPagar.OrigemDebito  = campo[22].toString();
                    daoFinanceiroAPagar.SituacaoTitulo  = campo[23].toString();
                    daoFinanceiroAPagar.SituacaoSispag  = campo[24].toString();
                    daoFinanceiroAPagar.TipoPagamento  = campo[25].toString();
                    daoFinanceiroAPagar.CodigoBarras  = campo[26].toString();
                    daoFinanceiroAPagar.Moeda  = campo[27].toString();
                    daoFinanceiroAPagar.ValorTitulo  = campo[28].toString();
                    daoFinanceiroAPagar.SaldoTitulo  = campo[29].toString();
                                        

                }
            }
            return daoFinanceiroAPagarList;
        }

        public string InserirDadosBD(DAOFinanceiroAPagarList daoFinanceiroAPagarList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspFinanceiroAPagarDeletar");

            try
            {
                DataTable dataTableFinanceiroAPagarList = ConvertToDataTable(daoFinanceiroAPagarList);
                foreach (DataRow linha in dataTableFinanceiroAPagarList.Rows)
                {
                    DAOFinanceiroAPagar daoFinanceiroAPagar = new DAOFinanceiroAPagar();

                    daofinanceiroAPagar.Empresa = linha["Empresa"].ToString();
                    daofinanceiroAPagar.Duplicata = linha["Duplicata"].ToString();
                    daofinanceiroAPagar.Parcela = linha["Parcela"].ToString();
                    daofinanceiroAPagar.DataContrato = Convert.ToDateTime(linha["DataContrato"].ToString());
                    daofinanceiroAPagar.TipoTitulo = linha["TipoTitulo"].ToString();
                    daofinanceiroAPagar.Documento = linha["Documento"].ToString();
                    daofinanceiroAPagar.Serie = linha["Serie"].ToString();
                    daofinanceiroAPagar.Historico = linha["Historico"].ToString();
                    daofinanceiroAPagar.EmpresaCobranca = linha["EmpresaCobranca"].ToString();
                    daofinanceiroAPagar.CodContabil = linha["CodContabil"].ToString();
                    daofinanceiroAPagar.CodFornecedor = linha["CodFornecedor"].ToString();
                    daofinanceiroAPagar.NomeFornecedor = linha["NomeFornecedor"].ToString();
                    daofinanceiroAPagar.TipoFornecedor = linha["TipoFornecedor"].ToString();
                    daofinanceiroAPagar.Transacao = linha["Transacao"].ToString();
                    daofinanceiroAPagar.DataTransacao = Convert.ToDateTime(linha["DataTransacao"].ToString());
                    daofinanceiroAPagar.Previsao = linha["Previsao"].ToString();
                    daofinanceiroAPagar.Portador = linha["Portador"].ToString();
                    daofinanceiroAPagar.VencimentoOrig = Convert.ToDateTime(linha["VencimentoOrig"].ToString());
                    daofinanceiroAPagar.Vencimento = Convert.ToDateTime(linha["Vencimento"].ToString());
                    daofinanceiroAPagar.Posicao = linha["Posicao"].ToString();
                    daofinanceiroAPagar.CentroCusto = linha["CentroCusto"].ToString();
                    daofinanceiroAPagar.NumContabil = linha["NumContabil"].ToString();
                    daofinanceiroAPagar.OrigemDebito = linha["OrigemDebito"].ToString();
                    daofinanceiroAPagar.SituacaoTitulo = linha["SituacaoTitulo"].ToString();
                    daofinanceiroAPagar.SituacaoSispag = linha["SituacaoSispag"].ToString();
                    daofinanceiroAPagar.TipoPagamento = linha["TipoPagamento"].ToString();
                    daofinanceiroAPagar.CodigoBarras = linha["CodigoBarras"].ToString();
                    daofinanceiroAPagar.Moeda = Convert.ToDecimal(linha["Moeda"].ToString());
                    daofinanceiroAPagar.ValorTitulo = Convert.ToDecimal(linha["ValorTitulo"].ToString());
                    daofinanceiroAPagar.SaldoTitulo = Convert.ToDecimal(linha["SaldoTitulo"].ToString());

                    
                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Empresa", daoFinanceiroAPagar.Empresa);
                    dalMySQL.AdicionaParametros("@Duplicata", daoFinanceiroAPagar.Duplicata);
                    dalMySQL.AdicionaParametros("@Parcela", daoFinanceiroAPagar.Parcela);
                    dalMySQL.AdicionaParametros("@DataContrato", daoFinanceiroAPagar.DataContrato);
                    dalMySQL.AdicionaParametros("@TipoTitulo", daoFinanceiroAPagar.TipoTitulo);
                    dalMySQL.AdicionaParametros("@Documento", daoFinanceiroAPagar.Documento);
                    dalMySQL.AdicionaParametros("@Serie", daoFinanceiroAPagar.Serie);
                    dalMySQL.AdicionaParametros("@Historico", daoFinanceiroAPagar.Historico);
                    dalMySQL.AdicionaParametros("@EmpresaCobranca", daoFinanceiroAPagar.EmpresaCobranca);
                    dalMySQL.AdicionaParametros("@CodContabil", daoFinanceiroAPagar.CodContabil);
                    dalMySQL.AdicionaParametros("@CodFornecedor", daoFinanceiroAPagar.CodFornecedor);
                    dalMySQL.AdicionaParametros("@NomeFornecedor", daoFinanceiroAPagar.NomeFornecedor);
                    dalMySQL.AdicionaParametros("@TipoFornecedor", daoFinanceiroAPagar.TipoFornecedor);
                    dalMySQL.AdicionaParametros("@Transacao", daoFinanceiroAPagar.Transacao);
                    dalMySQL.AdicionaParametros("@DataTransacao", daoFinanceiroAPagar.DataTransacao);
                    dalMySQL.AdicionaParametros("@Previsao", daoFinanceiroAPagar.Previsao);
                    dalMySQL.AdicionaParametros("@Portador", daoFinanceiroAPagar.Portador);
                    dalMySQL.AdicionaParametros("@VencimentoOrig", daoFinanceiroAPagar.VencimentoOrig);
                    dalMySQL.AdicionaParametros("@Vencimento", daoFinanceiroAPagar.Vencimento);
                    dalMySQL.AdicionaParametros("@Posicao", daoFinanceiroAPagar.Posicao);
                    dalMySQL.AdicionaParametros("@CentroCusto", daoFinanceiroAPagar.CentroCusto);
                    dalMySQL.AdicionaParametros("@NumContabil", daoFinanceiroAPagar.NumContabil);
                    dalMySQL.AdicionaParametros("@OrigemDebito", daoFinanceiroAPagar.OrigemDebito);
                    dalMySQL.AdicionaParametros("@SituacaoTitulo", daoFinanceiroAPagar.SituacaoTitulo);
                    dalMySQL.AdicionaParametros("@SituacaoSispag", daoFinanceiroAPagar.SituacaoSispag);
                    dalMySQL.AdicionaParametros("@TipoPagamento", daoFinanceiroAPagar.TipoPagamento);
                    dalMySQL.AdicionaParametros("@CodigoBarras", daoFinanceiroAPagar.CodigoBarras);
                    dalMySQL.AdicionaParametros("@Moeda", daoFinanceiroAPagar.Moeda);
                    dalMySQL.AdicionaParametros("@ValorTitulo", daoFinanceiroAPagar.ValorTitulo);
                    dalMySQL.AdicionaParametros("@SaldoTitulo", daoFinanceiroAPagar.SaldoTitulo);
                                        
                    dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspFinanceiroAPagarInserir");
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: financeiro contas a pagar inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir financeiro contas a pagar. Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de financeiro contas a pagar deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de financeiro contas a pagar renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
