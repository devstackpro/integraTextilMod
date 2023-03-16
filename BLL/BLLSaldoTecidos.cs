using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLSaldoTecidos
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de saldo de tecidos encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de saldo de tecidos não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: saldo de tecidos e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear saldo de tecidos Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOSaldoTecidosList LerCsv(string path)
        {
            DAOSaldoTecidosList daoSaldoTecidosList = new DAOSaldoTecidosList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOSaldoTecidos daoSaldoTecidos = new DAOSaldoTecidos();   
                campo = linha.Split(';');
                index++;

                if (index > 1)
                            {
                    daoSaldoTecidos.Col = campo[0].ToString();
                    daoSaldoTecidos.Colecao = campo[1].ToString();
                    daoSaldoTecidos.Lin = campo[2].ToString();
                    daoSaldoTecidos.Linha = campo[3].ToString();
                    daoSaldoTecidos.Art = campo[4].ToString();
                    daoSaldoTecidos.Artigo = campo[5].ToString();
                    daoSaldoTecidos.Nivel = campo[6].ToString();
                    daoSaldoTecidos.Grupo = campo[7].ToString();
                    daoSaldoTecidos.Sub = campo[8].ToString();
                    daoSaldoTecidos.Cor = campo[9].ToString();
                    daoSaldoTecidos.Produto = campo[10].ToString();
                    daoSaldoTecidos.NomeGrupo = campo[11].ToString();
                    daoSaldoTecidos.NomeSub = campo[12].ToString();
                    daoSaldoTecidos.NomeCor = campo[13].ToString();
                    daoSaldoTecidos.Narrativa = campo[14].ToString();
                    daoSaldoTecidos.EmEstoque = Convert.ToDecimal(campo[15].ToString());
                    daoSaldoTecidos.EmProducao = Convert.ToDecimal(campo[16].ToString());
                    daoSaldoTecidos.Necessidade = Convert.ToDecimal(campo[17].ToString());
                    daoSaldoTecidos.Saldo = Convert.ToDecimal(campo[18].ToString());
                    daoSaldoTecidos.Sobra = Convert.ToDecimal(campo[19].ToString());
                    daoSaldoTecidos.Falta = Convert.ToDecimal(campo[20].ToString());
                    daoSaldoTecidos.Vendido = Convert.ToDecimal(campo[21].ToString());
                    daoSaldoTecidos.Faturado = Convert.ToDecimal(campo[22].ToString());
                    daoSaldoTecidos.AFaturar = Convert.ToDecimal(campo[23].ToString());
                    daoSaldoTecidos.Cancelado = Convert.ToDecimal(campo[24].ToString());
                    daoSaldoTecidos.Bloqueado = campo[25].ToString();
                    daoSaldoTecidos.Estoque = Convert.ToDecimal(campo[26].ToString());
                    daoSaldoTecidos.ProntaEntrega = Convert.ToDecimal(campo[27].ToString());
                    daoSaldoTecidos.Outros = Convert.ToDecimal(campo[28].ToString());


                    daoSaldoTecidosList.Add(daoSaldoTecidos);
                    
                }
            }
            return daoSaldoTecidosList;
        }

        public string InserirDadosBD(DAOSaldoTecidosList daoSaldoTecidosList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspSaldoTecidosDeletar");

            try
            {
                DataTable dataTableSaldoTecidosList = ConvertToDataTable(daoSaldoTecidosList);
                foreach (DataRow linha in dataTableSaldoTecidosList.Rows)
                {
                    DAOSaldoTecidos daoSaldoTecidos = new DAOSaldoTecidos();
                    daoSaldoTecidos.Col = linha["Col"].ToString();
                    daoSaldoTecidos.Colecao = linha["Colecao"].ToString();
                    daoSaldoTecidos.Lin = linha["Lin"].ToString();
                    daoSaldoTecidos.Linha = linha["Linha"].ToString();
                    daoSaldoTecidos.Art = linha["Art"].ToString();
                    daoSaldoTecidos.Artigo = linha["Artigo"].ToString();
                    daoSaldoTecidos.Nivel = linha["Nivel"].ToString();
                    daoSaldoTecidos.Grupo = linha["Grupo"].ToString();
                    daoSaldoTecidos.Sub = linha["Sub"].ToString();
                    daoSaldoTecidos.Cor = linha["Cor"].ToString();
                    daoSaldoTecidos.Produto = linha["Produto"].ToString();
                    daoSaldoTecidos.NomeGrupo = linha["NomeGrupo"].ToString();
                    daoSaldoTecidos.NomeSub = linha["NomeSub"].ToString();
                    daoSaldoTecidos.NomeCor = linha["NomeCor"].ToString();
                    daoSaldoTecidos.Narrativa = linha["Narrativa"].ToString();
                    daoSaldoTecidos.EmEstoque = Convert.ToDecimal(linha["EmEstoque"].ToString());
                    daoSaldoTecidos.EmProducao = Convert.ToDecimal(linha["EmProducao"].ToString());
                    daoSaldoTecidos.Necessidade = Convert.ToDecimal(linha["Necessidade"].ToString());
                    daoSaldoTecidos.Saldo = Convert.ToDecimal(linha["Saldo"].ToString());
                    daoSaldoTecidos.Sobra = Convert.ToDecimal(linha["Sobra"].ToString());
                    daoSaldoTecidos.Falta = Convert.ToDecimal(linha["Falta"].ToString());
                    daoSaldoTecidos.Vendido = Convert.ToDecimal(linha["Vendido"].ToString());
                    daoSaldoTecidos.Faturado = Convert.ToDecimal(linha["Faturado"].ToString());
                    daoSaldoTecidos.AFaturar = Convert.ToDecimal(linha["AFaturar"].ToString());
                    daoSaldoTecidos.Cancelado = Convert.ToDecimal(linha["Cancelado"].ToString());
                    daoSaldoTecidos.Bloqueado = linha["Bloqueado"].ToString();
                    daoSaldoTecidos.Estoque = Convert.ToDecimal(linha["Estoque"].ToString());
                    daoSaldoTecidos.ProntaEntrega = Convert.ToDecimal(linha["ProntaEntrega"].ToString());
                    daoSaldoTecidos.Outros = Convert.ToDecimal(linha["Outros"].ToString());

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Emp", daoSaldoTecidos.Emp);
                    dalMySQL.AdicionaParametros("@Col", daoSaldoTecidos.Col);
                    dalMySQL.AdicionaParametros("@Colecao", daoSaldoTecidos.Colecao);
                    dalMySQL.AdicionaParametros("@Lin", daoSaldoTecidos.Lin);
                    dalMySQL.AdicionaParametros("@Linha", daoSaldoTecidos.Linha);
                    dalMySQL.AdicionaParametros("@Art", daoSaldoTecidos.Art);
                    dalMySQL.AdicionaParametros("@Artigo", daoSaldoTecidos.Artigo);
                    dalMySQL.AdicionaParametros("@Nivel", daoSaldoTecidos.Nivel);
                    dalMySQL.AdicionaParametros("@Grupo", daoSaldoTecidos.Grupo);
                    dalMySQL.AdicionaParametros("@Sub", daoSaldoTecidos.Sub);
                    dalMySQL.AdicionaParametros("@Cor", daoSaldoTecidos.Cor);
                    dalMySQL.AdicionaParametros("@Produto", daoSaldoTecidos.Produto);
                    dalMySQL.AdicionaParametros("@NomeGrupo", daoSaldoTecidos.NomeGrupo);
                    dalMySQL.AdicionaParametros("@NomeSub", daoSaldoTecidos.NomeSub);
                    dalMySQL.AdicionaParametros("@NomeCor", daoSaldoTecidos.NomeCor);
                    dalMySQL.AdicionaParametros("@Narrativa", daoSaldoTecidos.Narrativa);
                    dalMySQL.AdicionaParametros("@EmEstoque", daoSaldoTecidos.EmEstoque);
                    dalMySQL.AdicionaParametros("@EmProducao", daoSaldoTecidos.EmProducao);
                    dalMySQL.AdicionaParametros("@Necessidade", daoSaldoTecidos.Necessidade);
                    dalMySQL.AdicionaParametros("@Saldo", daoSaldoTecidos.Saldo);
                    dalMySQL.AdicionaParametros("@Sobra", daoSaldoTecidos.Sobra);
                    dalMySQL.AdicionaParametros("@Falta", daoSaldoTecidos.Falta);
                    dalMySQL.AdicionaParametros("@Vendido", daoSaldoTecidos.Vendido);
                    dalMySQL.AdicionaParametros("@Faturado", daoSaldoTecidos.Faturado);
                    dalMySQL.AdicionaParametros("@AFaturar", daoSaldoTecidos.AFaturar);
                    dalMySQL.AdicionaParametros("@Cancelado", daoSaldoTecidos.Cancelado);
                    dalMySQL.AdicionaParametros("@Bloqueado", daoSaldoTecidos.Bloqueado);
                    dalMySQL.AdicionaParametros("@Estoque", daoSaldoTecidos.Estoque);
                    dalMySQL.AdicionaParametros("@ProntaEntrega", daoSaldoTecidos.ProntaEntrega);
                    dalMySQL.AdicionaParametros("@Outros", daoSaldoTecidos.Outros);
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: saldo de tecidos inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir saldo de tecidos Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de saldo de tecidos deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de saldo de tecidos renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
