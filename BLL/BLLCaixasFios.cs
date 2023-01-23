using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLCaixasFios
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de caixas fios encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de caixas fios não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Caixas fios e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear Caixas fios. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOCaixasFiosList LerCsv(string path)
        {
            DAOCaixasFiosList daoCaixasFiosList = new DAOCaixasFiosList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOCaixasFios daoCaixasFios = new DAOCaixasFios();
                campo = linha.Split(';');
                index++;

                if (index > 1)
                {
                    daoCaixasFios.Emp = campo[0].ToString();
                    daoCaixasFios.Empresa = campo[1].ToString();
                    daoCaixasFios.TipoTitulo = campo[2].ToString();
                    daoCaixasFios.Portador = campo[3].ToString();
                    daoCaixasFios.Posicao = campo[4].ToString();
                    daoCaixasFios.CentroCusto = campo[5].ToString();
                    daoCaixasFios.DataEmissao = Convert.ToDateTime(campo[6].ToString());
                    daoCaixasFios.DataVencto = Convert.ToDateTime(campo[7].ToString());
                    daoCaixasFios.DataPagto = Convert.ToDateTime(campo[8].ToString());
                    daoCaixasFios.ValorParcela = Convert.ToDecimal(campo[9].ToString());
                    daoCaixasFios.ValorPago = Convert.ToDecimal(campo[10].ToString());
                    daoCaixasFios.ValorJuros = Convert.ToDecimal(campo[11].ToString());
                    daoCaixasFios.ValorDesconto = Convert.ToDecimal(campo[12].ToString());
                    daoCaixasFios.ValorAbatido = Convert.ToDecimal(campo[13].ToString());
                    daoCaixasFios.SaldoParcela = Convert.ToDecimal(campo[14].ToString());

                    daoCaixasFiosList.Add(daoCaixasFios);
                }
            }
            return daoCaixasFiosList;
        }

        public string InserirDadosBD(DAOCaixasFiosList daoCaixasFiosList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspCaixasFiosDeletar");

            try
            {
                DataTable dataTableCaixasFiosList = ConvertToDataTable(daoCaixasFiosList);
                foreach (DataRow linha in dataTableCaixasFiosList.Rows)
                {
                    DAOCaixasFios daoCaixasFios = new DAOCaixasFios();

                    daoCaixasFios.Fornecedor = linha["Fornecedor"].ToString();
                    daoCaixasFios.Duplicata = linha["Duplicata"].ToString();
                    daoCaixasFios.TipoTitulo = linha["TipoTitulo"].ToString();
                    daoCaixasFios.Portador = linha["Portador"].ToString();
                    daoCaixasFios.Posicao = linha["Posicao"].ToString();
                    daoCaixasFios.CentroCusto = linha["CentroCusto"].ToString();
                    daoCaixasFios.DataEmissao = Convert.ToDateTime(linha["DataEmissao"].ToString());
                    daoCaixasFios.DataVencto = Convert.ToDateTime(linha["DataVencto"].ToString());
                    daoCaixasFios.DataPagto = Convert.ToDateTime(linha["DataPagto"].ToString());
                    daoCaixasFios.ValorParcela = Convert.ToDecimal(linha["ValorParcela"].ToString());
                    daoCaixasFios.ValorPago = Convert.ToDecimal(linha["ValorPago"].ToString());
                    daoCaixasFios.ValorJuros = Convert.ToDecimal(linha["ValorJuros"].ToString());
                    daoCaixasFios.ValorDesconto = Convert.ToDecimal(linha["ValorDesconto"].ToString());
                    daoCaixasFios.ValorAbatido = Convert.ToDecimal(linha["ValorAbatido"].ToString());
                    daoCaixasFios.SaldoParcela = Convert.ToDecimal(linha["SaldoParcela"].ToString());

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Fornecedor", daoCaixasFios.Fornecedor);
                    dalMySQL.AdicionaParametros("@Duplicata", daoCaixasFios.Duplicata);
                    dalMySQL.AdicionaParametros("@TipoTitulo", daoCaixasFios.TipoTitulo);
                    dalMySQL.AdicionaParametros("@Portador", daoCaixasFios.Portador);
                    dalMySQL.AdicionaParametros("@Posicao", daoCaixasFios.Posicao);
                    dalMySQL.AdicionaParametros("@CentroCusto", daoCaixasFios.CentroCusto);
                    dalMySQL.AdicionaParametros("@DataEmissao", daoCaixasFios.DataEmissao);
                    dalMySQL.AdicionaParametros("@DataVencto", daoCaixasFios.DataVencto);
                    dalMySQL.AdicionaParametros("@DataPagto", daoCaixasFios.DataPagto);
                    dalMySQL.AdicionaParametros("@ValorParcela", daoCaixasFios.ValorParcela);
                    dalMySQL.AdicionaParametros("@ValorPago", daoCaixasFios.ValorPago);
                    dalMySQL.AdicionaParametros("@ValorJuros", daoCaixasFios.ValorJuros);
                    dalMySQL.AdicionaParametros("@ValorDesconto", daoCaixasFios.ValorDesconto);
                    dalMySQL.AdicionaParametros("@ValorAbatido", daoCaixasFios.ValorAbatido);
                    dalMySQL.AdicionaParametros("@SaldoParcela", daoCaixasFios.SaldoParcela);


                    dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspCaixasFiosInserir");
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Caixas fios inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir Caixas fios. Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de Caixas fios deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de Caixas fios renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
