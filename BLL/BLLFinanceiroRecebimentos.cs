using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLFinanceiroRecebimentos
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de Financeiro Recebimentos, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de Financeiro Recebimentos não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Financeiro Recebimentos e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear Financeiro Recebimentos Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOFinanceiroRecebimentosList LerCsv(string path)
        {
            DAOFinanceiroRecebimentosList daoFinanceiroRecebimentosList = new DAOFinanceiroRecebimentosList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOFinanceiroRecebimentos daoFinanceiroRecebimentos = new DAOFinanceiroRecebimentos();   
                campo = linha.Split(';');
                index++;

                if (index > 1)
                            {
                    daoRecebimentos.Empresa = campo[0].toString();
                    daoRecebimentos.NumDuplicata = campo[1].toString();
                    daoRecebimentos.Parcela = campo[2].toString();
                    daoRecebimentos.TipoTitulo = campo[3].toString();
                    daoRecebimentos.SituacaoDuplic = campo[4].toString();
                    daoRecebimentos.Emissao = Convert.ToDecimal(campo[5].toString());
                    daoRecebimentos.VencOriginal = Convert.ToDecimal(campo[6].toString());
                    daoRecebimentos.Vencimento = Convert.ToDecimal(campo[7].toString());
                    daoRecebimentos.CodCliente = campo[8].toString();
                    daoRecebimentos.Cliente = campo[9].toString();
                    daoRecebimentos.CodResponsavel = campo[10].toString();
                    daoRecebimentos.Responsavel = campo[11].toString();
                    daoRecebimentos.CodEndosso = campo[12].toString();
                    daoRecebimentos.Endosso = campo[13].toString();
                    daoRecebimentos.PedidoVenda = campo[14].toString();
                    daoRecebimentos.Representante = campo[15].toString();
                    daoRecebimentos.PercComissao = Convert.ToDecimal(campo[16].toString());
                    daoRecebimentos.BaseComissao = Convert.ToDecimal(campo[17].toString());
                    daoRecebimentos.ValorComissao = Convert.ToDecimal(campo[18].toString());
                    daoRecebimentos.Portador = campo[19].toString();
                    daoRecebimentos.NumeroBordero = campo[20].toString();
                    daoRecebimentos.NumeroRemessa = campo[21].toString();
                    daoRecebimentos.NrTituloBanco = campo[22].toString();
                    daoRecebimentos.ContaCorrente = campo[23].toString();
                    daoRecebimentos.CodCarteira = campo[24].toString();
                    daoRecebimentos.Transacao = campo[25].toString();
                    daoRecebimentos.PercDesconto = Convert.ToDecimal(campo[26].toString());
                    daoRecebimentos.NrSolicitacao = campo[27].toString();
                    daoRecebimentos.ValorDuplicata = Convert.ToDecimal(campo[28].toString());
                    daoRecebimentos.SaldoDuplicata = campo[29].toString();
                    daoRecebimentos.Moeda = Convert.ToDecimal(campo[30].toString());
                    daoRecebimentos.Posicao = campo[31].toString();
                    daoRecebimentos.LocalEmpresa = campo[32].toString();
                    daoRecebimentos.SituacaoDuplicata = campo[33].toString();
                    daoRecebimentos.Historico = campo[34].toString();
                    daoRecebimentos.ComplHistorico = campo[35].toString();
                    daoRecebimentos.NumContabil = campo[36].toString();
                    daoRecebimentos.FormaPagto = campo[37].toString();
                    daoRecebimentos.CodBarras = campo[38].toString();
                    daoRecebimentos.LinhaDigitavel = campo[39].toString();
                    daoRecebimentos.DuplicImpressa = campo[40].toString();
                    daoRecebimentos.Previsao = campo[41].toString();
                    daoRecebimentos.NumeroTitulo = campo[42].toString();
                    daoRecebimentos.NotaFiscal = campo[43].toString();
                    daoRecebimentos.Serie = campo[44].toString();
                    daoRecebimentos.CodCancelamento = campo[45].toString();
                    daoRecebimentos.Cancelamento = campo[46].toString();
                    daoRecebimentos.CodigoContabil = campo[47].toString();
                    daoRecebimentos.ValorMoeda = Convert.ToDecimal(campo[48].toString());
                    daoRecebimentos.CodUsuario = campo[49].toString();
                    daoRecebimentos.NumeroCaixa = campo[50].toString();
                    daoRecebimentos.NrAdiantamento = campo[51].toString();
                    daoRecebimentos.FantasiaCliente = campo[52].toString();
                    daoRecebimentos.TelefoneCliente = campo[53].toString();
                    daoRecebimentos.EmailCliente = campo[54].toString();
                    daoRecebimentos.Radm = campo[55].toString();
                    daoRecebimentos.Administrador = campo[56].toString();
                    daoRecebimentos.ComissaoAdministr = Convert.ToDecimal(campo[57].toString());
                    daoRecebimentos.SeqRcbto = campo[58].toString();
                    daoRecebimentos.DataRcnto = campo[59].toString();
                    daoRecebimentos.DataCredito = campo[60].toString();
                    daoRecebimentos.ValorRecebido = Convert.ToDecimal(campo[61].toString());
                    daoRecebimentos.ValorJuros = Convert.ToDecimal(campo[62].toString());
                    daoRecebimentos.ValorDesconto = Convert.ToDecimal(campo[63].toString());
                    daoRecebimentos.HisRcbto = campo[64].toString();
                    daoRecebimentos.HistoricoRcbto = campo[65].toString();
                    daoRecebimentos.NumeroDocumento = campo[66].toString();
                    daoRecebimentos.DoctoRcbto = campo[67].toString();
                    daoRecebimentos.PorRcbto = campo[68].toString();
                    daoRecebimentos.PortadorRcbto = campo[69].toString();
                    daoRecebimentos.ContaCorrenteRcbto = campo[70].toString();
                    daoRecebimentos.NumContabilRcbto = campo[71].toString();
                    daoRecebimentos.Atraso = Convert.ToDecimal(campo[72].toString());


                    daoFinanceiroRecebimentosList.Add(daoFinanceiroRecebimentos);
                    
                }
            }
            return daoFinanceiroRecebimentosList;
        }

        public string InserirDadosBD(DAOFinanceiroRecebimentosList daoFinanceiroRecebimentosList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspFinanceiroRecebimentosDeletar");

            try
            {
                DataTable dataTableFinanceiroRecebimentosList = ConvertToDataTable(daoFinanceiroRecebimentosList);
                foreach (DataRow linha in dataTableFinanceiroRecebimentosList.Rows)
                {
                    DAOFinanceiroRecebimentos daoFinanceiroRecebimentos = new DAOFinanceiroRecebimentos();

                    daoRecebimentos.Empresa = linha["Empresa"].toString();
                    daoRecebimentos.NumDuplicata = linha["NumDuplicata"].toString();
                    daoRecebimentos.Parcela = linha["Parcela"].toString();
                    daoRecebimentos.TipoTitulo = linha["TipoTitulo"].toString();
                    daoRecebimentos.SituacaoDuplic = linha["SituacaoDuplic"].toString();
                    daoRecebimentos.Emissao = Convert.ToDecimal(linha["Emissao"].toString());
                    daoRecebimentos.VencOriginal = Convert.ToDecimal(linha["VencOriginal"].toString());
                    daoRecebimentos.Vencimento = Convert.ToDecimal(linha["Vencimento"].toString());
                    daoRecebimentos.CodCliente = linha["CodCliente"].toString();
                    daoRecebimentos.Cliente = linha["Cliente"].toString();
                    daoRecebimentos.CodResponsavel = linha["CodResponsavel"].toString();
                    daoRecebimentos.Responsavel = linha["Responsavel"].toString();
                    daoRecebimentos.CodEndosso = linha["CodEndosso"].toString();
                    daoRecebimentos.Endosso = linha["Endosso"].toString();
                    daoRecebimentos.PedidoVenda = linha["PedidoVenda"].toString();
                    daoRecebimentos.Representante = linha["Representante"].toString();
                    daoRecebimentos.PercComissao = Convert.ToDecimal(linha["PercComissao"].toString());
                    daoRecebimentos.BaseComissao = Convert.ToDecimal(linha["BaseComissao"].toString());
                    daoRecebimentos.ValorComissao = Convert.ToDecimal(linha["ValorComissao"].toString());
                    daoRecebimentos.Portador = linha["Portador"].toString();
                    daoRecebimentos.NumeroBordero = linha["NumeroBordero"].toString();
                    daoRecebimentos.NumeroRemessa = linha["NumeroRemessa"].toString();
                    daoRecebimentos.NrTituloBanco = linha["NrTituloBanco"].toString();
                    daoRecebimentos.ContaCorrente = linha["ContaCorrente"].toString();
                    daoRecebimentos.CodCarteira = linha["CodCarteira"].toString();
                    daoRecebimentos.Transacao = linha["Transacao"].toString();
                    daoRecebimentos.PercDesconto = Convert.ToDecimal(linha["PercDesconto"].toString());
                    daoRecebimentos.NrSolicitacao = linha["NrSolicitacao"].toString();
                    daoRecebimentos.ValorDuplicata = Convert.ToDecimal(linha["ValorDuplicata"].toString());
                    daoRecebimentos.SaldoDuplicata = linha["SaldoDuplicata"].toString();
                    daoRecebimentos.Moeda = Convert.ToDecimal(linha["Moeda"].toString());
                    daoRecebimentos.Posicao = linha["Posicao"].toString();
                    daoRecebimentos.LocalEmpresa = linha["LocalEmpresa"].toString();
                    daoRecebimentos.SituacaoDuplicata = linha["SituacaoDuplicata"].toString();
                    daoRecebimentos.Historico = linha["Historico"].toString();
                    daoRecebimentos.ComplHistorico = linha["ComplHistorico"].toString();
                    daoRecebimentos.NumContabil = linha["NumContabil"].toString();
                    daoRecebimentos.FormaPagto = linha["FormaPagto"].toString();
                    daoRecebimentos.CodBarras = linha["CodBarras"].toString();
                    daoRecebimentos.LinhaDigitavel = linha["LinhaDigitavel"].toString();
                    daoRecebimentos.DuplicImpressa = linha["DuplicImpressa"].toString();
                    daoRecebimentos.Previsao = linha["Previsao"].toString();
                    daoRecebimentos.NumeroTitulo = linha["NumeroTitulo"].toString();
                    daoRecebimentos.NotaFiscal = linha["NotaFiscal"].toString();
                    daoRecebimentos.Serie = linha["Serie"].toString();
                    daoRecebimentos.CodCancelamento = linha["CodCancelamento"].toString();
                    daoRecebimentos.Cancelamento = linha["Cancelamento"].toString();
                    daoRecebimentos.CodigoContabil = linha["CodigoContabil"].toString();
                    daoRecebimentos.ValorMoeda = Convert.ToDecimal(linha["ValorMoeda"].toString());
                    daoRecebimentos.CodUsuario = linha["CodUsuario"].toString();
                    daoRecebimentos.NumeroCaixa = linha["NumeroCaixa"].toString();
                    daoRecebimentos.NrAdiantamento = linha["NrAdiantamento"].toString();
                    daoRecebimentos.FantasiaCliente = linha["FantasiaCliente"].toString();
                    daoRecebimentos.TelefoneCliente = linha["TelefoneCliente"].toString();
                    daoRecebimentos.EmailCliente = linha["EmailCliente"].toString();
                    daoRecebimentos.Radm = linha["Radm"].toString();
                    daoRecebimentos.Administrador = linha["Administrador"].toString();
                    daoRecebimentos.ComissaoAdministr = Convert.ToDecimal(linha["ComissaoAdministr"].toString());
                    daoRecebimentos.SeqRcbto = linha["SeqRcbto"].toString();
                    daoRecebimentos.DataRcnto = linha["DataRcnto"].toString();
                    daoRecebimentos.DataCredito = linha["DataCredito"].toString();
                    daoRecebimentos.ValorRecebido = Convert.ToDecimal(linha["ValorRecebido"].toString());
                    daoRecebimentos.ValorJuros = Convert.ToDecimal(linha["ValorJuros"].toString());
                    daoRecebimentos.ValorDesconto = Convert.ToDecimal(linha["ValorDesconto"].toString());
                    daoRecebimentos.HisRcbto = linha["HisRcbto"].toString();
                    daoRecebimentos.HistoricoRcbto = linha["HistoricoRcbto"].toString();
                    daoRecebimentos.NumeroDocumento = linha["NumeroDocumento"].toString();
                    daoRecebimentos.DoctoRcbto = linha["DoctoRcbto"].toString();
                    daoRecebimentos.PorRcbto = linha["PorRcbto"].toString();
                    daoRecebimentos.PortadorRcbto = linha["PortadorRcbto"].toString();
                    daoRecebimentos.ContaCorrenteRcbto = linha["ContaCorrenteRcbto"].toString();
                    daoRecebimentos.NumContabilRcbto = linha["NumContabilRcbto"].toString();
                    daoRecebimentos.Atraso = Convert.ToDecimal(linha["Atraso"].toString());

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Empresa", daoRecebimentos.Empresa);
                    dalMySQL.AdicionaParametros("@NumDuplicata", daoRecebimentos.NumDuplicata);
                    dalMySQL.AdicionaParametros("@Parcela", daoRecebimentos.Parcela);
                    dalMySQL.AdicionaParametros("@TipoTitulo", daoRecebimentos.TipoTitulo);
                    dalMySQL.AdicionaParametros("@SituacaoDuplic", daoRecebimentos.SituacaoDuplic);
                    dalMySQL.AdicionaParametros("@Emissao", daoRecebimentos.Emissao);
                    dalMySQL.AdicionaParametros("@VencOriginal", daoRecebimentos.VencOriginal);
                    dalMySQL.AdicionaParametros("@Vencimento", daoRecebimentos.Vencimento);
                    dalMySQL.AdicionaParametros("@CodCliente", daoRecebimentos.CodCliente);
                    dalMySQL.AdicionaParametros("@Cliente", daoRecebimentos.Cliente);
                    dalMySQL.AdicionaParametros("@CodResponsavel", daoRecebimentos.CodResponsavel);
                    dalMySQL.AdicionaParametros("@Responsavel", daoRecebimentos.Responsavel);
                    dalMySQL.AdicionaParametros("@CodEndosso", daoRecebimentos.CodEndosso);
                    dalMySQL.AdicionaParametros("@Endosso", daoRecebimentos.Endosso);
                    dalMySQL.AdicionaParametros("@PedidoVenda", daoRecebimentos.PedidoVenda);
                    dalMySQL.AdicionaParametros("@Representante", daoRecebimentos.Representante);
                    dalMySQL.AdicionaParametros("@PercComissao", daoRecebimentos.PercComissao);
                    dalMySQL.AdicionaParametros("@BaseComissao", daoRecebimentos.BaseComissao);
                    dalMySQL.AdicionaParametros("@ValorComissao", daoRecebimentos.ValorComissao);
                    dalMySQL.AdicionaParametros("@Portador", daoRecebimentos.Portador);
                    dalMySQL.AdicionaParametros("@NumeroBordero", daoRecebimentos.NumeroBordero);
                    dalMySQL.AdicionaParametros("@NumeroRemessa", daoRecebimentos.NumeroRemessa);
                    dalMySQL.AdicionaParametros("@NrTituloBanco", daoRecebimentos.NrTituloBanco);
                    dalMySQL.AdicionaParametros("@ContaCorrente", daoRecebimentos.ContaCorrente);
                    dalMySQL.AdicionaParametros("@CodCarteira", daoRecebimentos.CodCarteira);
                    dalMySQL.AdicionaParametros("@Transacao", daoRecebimentos.Transacao);
                    dalMySQL.AdicionaParametros("@PercDesconto", daoRecebimentos.PercDesconto);
                    dalMySQL.AdicionaParametros("@NrSolicitacao", daoRecebimentos.NrSolicitacao);
                    dalMySQL.AdicionaParametros("@ValorDuplicata", daoRecebimentos.ValorDuplicata);
                    dalMySQL.AdicionaParametros("@SaldoDuplicata", daoRecebimentos.SaldoDuplicata);
                    dalMySQL.AdicionaParametros("@Moeda", daoRecebimentos.Moeda);
                    dalMySQL.AdicionaParametros("@Posicao", daoRecebimentos.Posicao);
                    dalMySQL.AdicionaParametros("@LocalEmpresa", daoRecebimentos.LocalEmpresa);
                    dalMySQL.AdicionaParametros("@SituacaoDuplicata", daoRecebimentos.SituacaoDuplicata);
                    dalMySQL.AdicionaParametros("@Historico", daoRecebimentos.Historico);
                    dalMySQL.AdicionaParametros("@ComplHistorico", daoRecebimentos.ComplHistorico);
                    dalMySQL.AdicionaParametros("@NumContabil", daoRecebimentos.NumContabil);
                    dalMySQL.AdicionaParametros("@FormaPagto", daoRecebimentos.FormaPagto);
                    dalMySQL.AdicionaParametros("@CodBarras", daoRecebimentos.CodBarras);
                    dalMySQL.AdicionaParametros("@LinhaDigitavel", daoRecebimentos.LinhaDigitavel);
                    dalMySQL.AdicionaParametros("@DuplicImpressa", daoRecebimentos.DuplicImpressa);
                    dalMySQL.AdicionaParametros("@Previsao", daoRecebimentos.Previsao);
                    dalMySQL.AdicionaParametros("@NumeroTitulo", daoRecebimentos.NumeroTitulo);
                    dalMySQL.AdicionaParametros("@NotaFiscal", daoRecebimentos.NotaFiscal);
                    dalMySQL.AdicionaParametros("@Serie", daoRecebimentos.Serie);
                    dalMySQL.AdicionaParametros("@CodCancelamento", daoRecebimentos.CodCancelamento);
                    dalMySQL.AdicionaParametros("@Cancelamento", daoRecebimentos.Cancelamento);
                    dalMySQL.AdicionaParametros("@CodigoContabil", daoRecebimentos.CodigoContabil);
                    dalMySQL.AdicionaParametros("@ValorMoeda", daoRecebimentos.ValorMoeda);
                    dalMySQL.AdicionaParametros("@CodUsuario", daoRecebimentos.CodUsuario);
                    dalMySQL.AdicionaParametros("@NumeroCaixa", daoRecebimentos.NumeroCaixa);
                    dalMySQL.AdicionaParametros("@NrAdiantamento", daoRecebimentos.NrAdiantamento);
                    dalMySQL.AdicionaParametros("@FantasiaCliente", daoRecebimentos.FantasiaCliente);
                    dalMySQL.AdicionaParametros("@TelefoneCliente", daoRecebimentos.TelefoneCliente);
                    dalMySQL.AdicionaParametros("@EmailCliente", daoRecebimentos.EmailCliente);
                    dalMySQL.AdicionaParametros("@Radm", daoRecebimentos.Radm);
                    dalMySQL.AdicionaParametros("@Administrador", daoRecebimentos.Administrador);
                    dalMySQL.AdicionaParametros("@ComissaoAdministr", daoRecebimentos.ComissaoAdministr);
                    dalMySQL.AdicionaParametros("@SeqRcbto", daoRecebimentos.SeqRcbto);
                    dalMySQL.AdicionaParametros("@DataRcnto", daoRecebimentos.DataRcnto);
                    dalMySQL.AdicionaParametros("@DataCredito", daoRecebimentos.DataCredito);
                    dalMySQL.AdicionaParametros("@ValorRecebido", daoRecebimentos.ValorRecebido);
                    dalMySQL.AdicionaParametros("@ValorJuros", daoRecebimentos.ValorJuros);
                    dalMySQL.AdicionaParametros("@ValorDesconto", daoRecebimentos.ValorDesconto);
                    dalMySQL.AdicionaParametros("@HisRcbto", daoRecebimentos.HisRcbto);
                    dalMySQL.AdicionaParametros("@HistoricoRcbto", daoRecebimentos.HistoricoRcbto);
                    dalMySQL.AdicionaParametros("@NumeroDocumento", daoRecebimentos.NumeroDocumento);
                    dalMySQL.AdicionaParametros("@DoctoRcbto", daoRecebimentos.DoctoRcbto);
                    dalMySQL.AdicionaParametros("@PorRcbto", daoRecebimentos.PorRcbto);
                    dalMySQL.AdicionaParametros("@PortadorRcbto", daoRecebimentos.PortadorRcbto);
                    dalMySQL.AdicionaParametros("@ContaCorrenteRcbto", daoRecebimentos.ContaCorrenteRcbto);
                    dalMySQL.AdicionaParametros("@NumContabilRcbto", daoRecebimentos.NumContabilRcbto);
                    dalMySQL.AdicionaParametros("@Atraso", daoRecebimentos.Atraso);
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Financeiro Recebimentos inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir Financeiro Recebimentos Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de Financeiro Recebimentos deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de Financeiro Recebimentos renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
