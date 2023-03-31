using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL 
{
    public class BLLFinanceiroAReceber
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de financeiro contas a receber encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasReceber.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de financeiro contas a receber não encontrado, nome: nulo. Detalhes: Classe: BLLContasReceber.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: financeiro contas a receber movido e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear financeiro contas a receber. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOFinanceiroAReceberList LerCsv(string path)
        {
            DAOFinanceiroAReceberList daoFinanceiroAReceberList = new DAOFinanceiroAReceberList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOFinanceiroAReceber daoFinanceiroAReceber = new DAOFinanceiroAReceber();
                campo = linha.Split(';');
                index++;

                if (index > 1)
                {
                    daoFinanceiroAReceber.Empresa = campo[0].ToString();
                    daoFinanceiroAReceber.NumDuplicata = campo[1].ToString();
                    daoFinanceiroAReceber.Parcela = campo[2].ToString();
                    daoFinanceiroAReceber.TipoTitulo = campo[3].ToString();
                    daoFinanceiroAReceber.SituacaoDuplic = campo[4].ToString();
                    daoFinanceiroAReceber.Emissao = Convert.ToDateTime(campo[5].ToString());
                    daoFinanceiroAReceber.VencOriginal = Convert.ToDateTime(campo[6].ToString());
                    daoFinanceiroAReceber.Vencimento = Convert.ToDateTime(campo[7].ToString());
                    daoFinanceiroAReceber.CodCliente = campo[8].ToString();
                    daoFinanceiroAReceber.Cliente = campo[9].ToString();
                    daoFinanceiroAReceber.CodResponsavel = campo[10].ToString();
                    daoFinanceiroAReceber.Responsavel = campo[11].ToString();
                    daoFinanceiroAReceber.CodEndosso = campo[12].ToString();
                    daoFinanceiroAReceber.Endosso = campo[13].ToString();
                    daoFinanceiroAReceber.PedidoVenda = campo[14].ToString();
                    daoFinanceiroAReceber.Representante = campo[15].ToString();
                    daoFinanceiroAReceber.PercComissao = Convert.ToDecimal(campo[16].ToString());
                    daoFinanceiroAReceber.BaseComissao = Convert.ToDecimal(campo[17].ToString());
                    daoFinanceiroAReceber.ValorComissao = Convert.ToDecimal(campo[18].ToString());
                    daoFinanceiroAReceber.Portador = campo[19].ToString();
                    daoFinanceiroAReceber.NumeroBordero = campo[20].ToString();
                    daoFinanceiroAReceber.NumeroRemessa = campo[21].ToString();
                    daoFinanceiroAReceber.NrTituloBanco = campo[22].ToString();
                    daoFinanceiroAReceber.ContaCorrente = campo[23].ToString();
                    daoFinanceiroAReceber.CodCarteira = campo[24].ToString();
                    daoFinanceiroAReceber.Transacao = campo[25].ToString();
                    daoFinanceiroAReceber.PercDesconto = Convert.ToDecimal(campo[26].ToString());
                    daoFinanceiroAReceber.NrSolicitacao = campo[27].ToString();
                    daoFinanceiroAReceber.ValorDuplicata = Convert.ToDecimal(campo[28].ToString());
                    daoFinanceiroAReceber.SaldoDuplicata = Convert.ToDecimal(campo[29].ToString());
                    daoFinanceiroAReceber.Moeda = campo[30].ToString();
                    daoFinanceiroAReceber.Prorrogacao = campo[31].ToString();
                    daoFinanceiroAReceber.Posicao = campo[32].ToString();
                    daoFinanceiroAReceber.LocalEmpresa = campo[33].ToString();
                    daoFinanceiroAReceber.SituacaoDuplicata = campo[34].ToString();
                    daoFinanceiroAReceber.Historico = campo[35].ToString();
                    daoFinanceiroAReceber.ComplHistorico = campo[36].ToString();
                    daoFinanceiroAReceber.NumContabil = campo[37].ToString();
                    daoFinanceiroAReceber.FormaPagto = campo[38].ToString();
                    daoFinanceiroAReceber.CodBarras = campo[39].ToString();
                    daoFinanceiroAReceber.LinhaDigitavel = campo[40].ToString();
                    daoFinanceiroAReceber.DuplicImpressa = Convert.ToDecimal(campo[41].ToString());
                    daoFinanceiroAReceber.Previsao = campo[42].ToString();
                    daoFinanceiroAReceber.NumeroTitulo = campo[43].ToString();
                    daoFinanceiroAReceber.NotaFiscal = campo[44].ToString();
                    daoFinanceiroAReceber.Serie = campo[45].ToString();
                    daoFinanceiroAReceber.CodCancelamento = campo[46].ToString();
                    daoFinanceiroAReceber.Cancelamento = campo[47].ToString();
                    daoFinanceiroAReceber.CodigoContabil = campo[48].ToString();
                    daoFinanceiroAReceber.ValorMoeda = Convert.ToDecimal(campo[49].ToString());
                    daoFinanceiroAReceber.CodUsuario = campo[50].ToString();
                    daoFinanceiroAReceber.NumeroCaixa = campo[51].ToString();
                    daoFinanceiroAReceber.NrAdiantamento = campo[52].ToString();
                    daoFinanceiroAReceber.FantasiaCliente = campo[53].ToString();
                    daoFinanceiroAReceber.TelefoneCliente = campo[54].ToString();
                    daoFinanceiroAReceber.EmailCliente = campo[55].ToString();
                    daoFinanceiroAReceber.Radm = campo[56].ToString();
                    daoFinanceiroAReceber.Administrador = campo[57].ToString();
                    daoFinanceiroAReceber.ComissaoAdministr = Convert.ToDecimal(campo[58].ToString());

                    
                    daoFinanceiroAReceberList.Add(daoFinanceiroAReceber);
                }
            }
            return daoFinanceiroAReceberList;
        }

        public string InserirDadosBD(DAOFinanceiroAReceberList daoFinanceiroAReceberList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspFinanceiroAReceberDeletar");

            try
            {
                DataTable dataTableFinanceiroAReceberList = ConvertToDataTable(daoFinanceiroAReceberList);
                foreach (DataRow linha in dataTableFinanceiroAReceberList.Rows)
                {
                    DAOFinanceiroAReceber daoFinanceiroAReceber = new DAOFinanceiroAReceber();

                    daoContasReceber.Empresa = linha["Empresa"].ToString();
                    daoContasReceber.NumDuplicata = linha["NumDuplicata"].ToString();
                    daoContasReceber.Parcela = linha["Parcela"].ToString(); 
                    daoContasReceber.TipoTitulo = linha["TipoTitulo"].ToString();
                    daoContasReceber.SituacaoDuplic = linha["SituacaoDuplic"].ToString();
                    daoContasReceber.Emissao = Convert.ToDateTime(linha["Emissao"].ToString());
                    daoContasReceber.VencOriginal = Convert.ToDateTime(linha["VencOriginal"].ToString());
                    daoContasReceber.Vencimento = Convert.ToDateTime(linha["Vencimento"].ToString());
                    daoContasReceber.CodCliente = linha["CodCliente"].ToString();
                    daoContasReceber.Cliente = linha["Cliente"].ToString();
                    daoContasReceber.CodResponsavel = linha["CodResponsavel"].ToString();
                    daoContasReceber.Responsavel = linha["Responsavel"].ToString();
                    daoContasReceber.CodEndosso = linha["CodEndosso"].ToString();
                    daoContasReceber.Endosso = linha["Endosso"].ToString();
                    daoContasReceber.PedidoVenda = linha["PedidoVenda"].ToString();
                    daoContasReceber.Representante = linha["Representante"].ToString();
                    daoContasReceber.PercComissao = Convert.ToDecimal(linha["PercComissao"].ToString());
                    daoContasReceber.BaseComissao = Convert.ToDecimal(linha["BaseComissao"].ToString());
                    daoContasReceber.ValorComissao = Convert.ToDecimal(linha["ValorComissao"].ToString());
                    daoContasReceber.Portador = linha["Portador"].ToString();
                    daoContasReceber.NumeroBordero = linha["NumeroBordero"].ToString();
                    daoContasReceber.NumeroRemessa = linha["NumeroRemessa"].ToString();
                    daoContasReceber.NrTituloBanco = linha["NrTituloBanco"].ToString();
                    daoContasReceber.ContaCorrente = linha["ContaCorrente"].ToString();
                    daoContasReceber.CodCarteira = linha["CodCarteira"].ToString();
                    daoContasReceber.Transacao = linha["Transacao"].ToString();
                    daoContasReceber.PercDesconto = Convert.ToDecimal(linha["PercDesconto"].ToString());
                    daoContasReceber.NrSolicitacao = linha["NrSolicitacao"].ToString();
                    daoContasReceber.ValorDuplicata = Convert.ToDecimal(linha["ValorDuplicata"].ToString());
                    daoContasReceber.SaldoDuplicata = Convert.ToDecimal(linha["SaldoDuplicata"].ToString());
                    daoContasReceber.Moeda = linha["Moeda"].ToString();
                    daoContasReceber.Prorrogacao = linha["Prorrogacao"].ToString();
                    daoContasReceber.Posicao = linha["Posicao"].ToString();
                    daoContasReceber.LocalEmpresa = linha["LocalEmpresa"].ToString();
                    daoContasReceber.SituacaoDuplicata = linha["SituacaoDuplicata"].ToString();
                    daoContasReceber.Historico = linha["Historico"].ToString();
                    daoContasReceber.ComplHistorico = linha["ComplHistorico"].ToString();
                    daoContasReceber.NumContabil = linha["NumContabil"].ToString();
                    daoContasReceber.FormaPagto = linha["FormaPagto"].ToString();
                    daoContasReceber.CodBarras = linha["CodBarras"].ToString();
                    daoContasReceber.LinhaDigitavel = linha["LinhaDigitavel"].ToString();
                    daoContasReceber.DuplicImpressa = Convert.ToDecimal(linha["DuplicImpressa"].ToString());
                    daoContasReceber.Previsao = linha["Previsao"].ToString();
                    daoContasReceber.NumeroTitulo = linha["NumeroTitulo"].ToString();
                    daoContasReceber.NotaFiscal = linha["NotaFiscal"].ToString();
                    daoContasReceber.Serie = linha["Serie"].ToString();
                    daoContasReceber.CodCancelamento = linha["CodCancelamento"].ToString();
                    daoContasReceber.Cancelamento = linha["Cancelamento"].ToString();
                    daoContasReceber.CodigoContabil = linha["CodigoContabil"].ToString();
                    daoContasReceber.ValorMoeda = Convert.ToDecimal(linha["ValorMoeda"].ToString());
                    daoContasReceber.CodUsuario = linha["CodUsuario"].ToString();
                    daoContasReceber.NumeroCaixa = linha["NumeroCaixa"].ToString();
                    daoContasReceber.NrAdiantamento = linha["NrAdiantamento"].ToString();
                    daoContasReceber.FantasiaCliente = linha["FantasiaCliente"].ToString();
                    daoContasReceber.TelefoneCliente = linha["TelefoneCliente"].ToString();
                    daoContasReceber.EmailCliente = linha["EmailCliente"].ToString();
                    daoContasReceber.Radm = linha["Radm"].ToString();
                    daoContasReceber.Administrador = linha["Administrador"].ToString();
                    daoContasReceber.ComissaoAdministr = Convert.ToDecimal(linha["ComissaoAdministr"].ToString());

                    
                    dalMySQL.LimparParametros();

                        dalMySQL.AdicionaParametros("@Empresa", daoFinanceiroAReceber.Empresa);
                        dalMySQL.AdicionaParametros("@NumDuplicata", daoFinanceiroAReceber.NumDuplicata);
                        dalMySQL.AdicionaParametros("@Parcela", daoFinanceiroAReceber.Parcela); 
                        dalMySQL.AdicionaParametros("@TipoTitulo", daoFinanceiroAReceber.TipoTitulo);
                        dalMySQL.AdicionaParametros("@SituacaoDuplic", daoFinanceiroAReceber.SituacaoDuplic);
                        dalMySQL.AdicionaParametros("@Emissao", daoFinanceiroAReceber.Emissao);
                        dalMySQL.AdicionaParametros("@VencOriginal", daoFinanceiroAReceber.VencOriginal);
                        dalMySQL.AdicionaParametros("@Vencimento", daoFinanceiroAReceber.Vencimento);
                        dalMySQL.AdicionaParametros("@CodCliente", daoFinanceiroAReceber.CodCliente);
                        dalMySQL.AdicionaParametros("@Cliente", daoFinanceiroAReceber.Cliente);
                        dalMySQL.AdicionaParametros("@CodResponsavel", daoFinanceiroAReceber.CodResponsavel);
                        dalMySQL.AdicionaParametros("@Responsavel", daoFinanceiroAReceber.Responsavel);
                        dalMySQL.AdicionaParametros("@CodEndosso", daoFinanceiroAReceber.CodEndosso);
                        dalMySQL.AdicionaParametros("@Endosso", daoFinanceiroAReceber.Endosso);
                        dalMySQL.AdicionaParametros("@PedidoVenda", daoFinanceiroAReceber.PedidoVenda);
                        dalMySQL.AdicionaParametros("@Representante", daoFinanceiroAReceber.Representante);
                        dalMySQL.AdicionaParametros("@PercComissao", daoFinanceiroAReceber.PercComissao);
                        dalMySQL.AdicionaParametros("@BaseComissao", daoFinanceiroAReceber.BaseComissao);
                        dalMySQL.AdicionaParametros("@ValorComissao", daoFinanceiroAReceber.ValorComissao);
                        dalMySQL.AdicionaParametros("@Portador", daoFinanceiroAReceber.Portador);
                        dalMySQL.AdicionaParametros("@NumeroBordero", daoFinanceiroAReceber.NumeroBordero);
                        dalMySQL.AdicionaParametros("@NumeroRemessa", daoFinanceiroAReceber.NumeroRemessa);
                        dalMySQL.AdicionaParametros("@NrTituloBanco", daoFinanceiroAReceber.NrTituloBanco);
                        dalMySQL.AdicionaParametros("@ContaCorrente", daoFinanceiroAReceber.ContaCorrente);
                        dalMySQL.AdicionaParametros("@CodCarteira", daoFinanceiroAReceber.CodCarteira);
                        dalMySQL.AdicionaParametros("@Transacao", daoFinanceiroAReceber.Transacao);
                        dalMySQL.AdicionaParametros("@PercDesconto", daoFinanceiroAReceber.PercDesconto);
                        dalMySQL.AdicionaParametros("@NrSolicitacao", daoFinanceiroAReceber.NrSolicitacao);
                        dalMySQL.AdicionaParametros("@ValorDuplicata", daoFinanceiroAReceber.ValorDuplicata);
                        dalMySQL.AdicionaParametros("@SaldoDuplicata", daoFinanceiroAReceber.SaldoDuplicata);
                        dalMySQL.AdicionaParametros("@Moeda", daoFinanceiroAReceber.Moeda);
                        dalMySQL.AdicionaParametros("@Prorrogacao", daoFinanceiroAReceber.Prorrogacao);
                        dalMySQL.AdicionaParametros("@Posicao", daoFinanceiroAReceber.Posicao);
                        dalMySQL.AdicionaParametros("@LocalEmpresa", daoFinanceiroAReceber.LocalEmpresa);
                        dalMySQL.AdicionaParametros("@SituacaoDuplicata", daoFinanceiroAReceber.SituacaoDuplicata);
                        dalMySQL.AdicionaParametros("@Historico", daoFinanceiroAReceber.Historico);
                        dalMySQL.AdicionaParametros("@ComplHistorico", daoFinanceiroAReceber.ComplHistorico);
                        dalMySQL.AdicionaParametros("@NumContabil", daoFinanceiroAReceber.NumContabil);
                        dalMySQL.AdicionaParametros("@FormaPagto", daoFinanceiroAReceber.FormaPagto);
                        dalMySQL.AdicionaParametros("@CodBarras", daoFinanceiroAReceber.CodBarras);
                        dalMySQL.AdicionaParametros("@LinhaDigitavel", daoFinanceiroAReceber.LinhaDigitavel);
                        dalMySQL.AdicionaParametros("@DuplicImpressa", daoFinanceiroAReceber.DuplicImpressa);
                        dalMySQL.AdicionaParametros("@Previsao", daoFinanceiroAReceber.Previsao);
                        dalMySQL.AdicionaParametros("@NumeroTitulo", daoFinanceiroAReceber.NumeroTitulo);
                        dalMySQL.AdicionaParametros("@NotaFiscal", daoFinanceiroAReceber.NotaFiscal);
                        dalMySQL.AdicionaParametros("@Serie", daoFinanceiroAReceber.Serie);
                        dalMySQL.AdicionaParametros("@CodCancelamento", daoFinanceiroAReceber.CodCancelamento);
                        dalMySQL.AdicionaParametros("@Cancelamento", daoFinanceiroAReceber.Cancelamento);
                        dalMySQL.AdicionaParametros("@CodigoContabil", daoFinanceiroAReceber.CodigoContabil);
                        dalMySQL.AdicionaParametros("@ValorMoeda", daoFinanceiroAReceber.ValorMoeda);
                        dalMySQL.AdicionaParametros("@CodUsuario", daoFinanceiroAReceber.CodUsuario);
                        dalMySQL.AdicionaParametros("@NumeroCaixa", daoFinanceiroAReceber.NumeroCaixa);
                        dalMySQL.AdicionaParametros("@NrAdiantamento", daoFinanceiroAReceber.NrAdiantamento);
                        dalMySQL.AdicionaParametros("@FantasiaCliente", daoFinanceiroAReceber.FantasiaCliente);
                        dalMySQL.AdicionaParametros("@TelefoneCliente", daoFinanceiroAReceber.TelefoneCliente);
                        dalMySQL.AdicionaParametros("@EmailCliente", daoFinanceiroAReceber.EmailCliente);
                        dalMySQL.AdicionaParametros("@Radm", daoFinanceiroAReceber.Radm);
                        dalMySQL.AdicionaParametros("@Administrador", daoFinanceiroAReceber.Administrador);
                        dalMySQL.AdicionaParametros("@ComissaoAdministr", daoFinanceiroAReceber.ComissaoAdministr);

                    
                    dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspFinanceiroAReceberInserir");
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: financeiro contas a receber inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir financeiro contas a receber. Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de financeiro contas a receber deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de financeiro contas a receber renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
