using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLPosicaoOp
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de posicao das ops, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de posicao das ops não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: posicao das ops e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear posicao das ops Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOPosicaoOpList LerCsv(string path)
        {
            DAOPosicaoOpList daoPosicaoOpList = new DAOPosicaoOpList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOPosicaoOp daoPosicaoOp = new DAOPosicaoOp();   
                campo = linha.Split(';');
                index++;

                if (index > 1)
                            {
                    daoPosicaoOp.Col = campo[0].ToString();
                    daoPosicaoOp.Colecao = campo[1].ToString();
                    daoPosicaoOp.Lin = campo[2].ToString();
                    daoPosicaoOp.Linha = campo[3].ToString();
                    daoPosicaoOp.Art = campo[4].ToString();
                    daoPosicaoOp.Artigo = campo[5].ToString();
                    daoPosicaoOp.Nivel = campo[6].ToString();
                    daoPosicaoOp.Grupo = campo[7].ToString();
                    daoPosicaoOp.Sub = campo[8].ToString();
                    daoPosicaoOp.Cor = campo[9].ToString();
                    daoPosicaoOp.Produto = campo[10].ToString();
                    daoPosicaoOp.NomeGrupo = campo[11].ToString();
                    daoPosicaoOp.NomeSub = campo[12].ToString();
                    daoPosicaoOp.NomeCor = campo[13].ToString();
                    daoPosicaoOp.Narrativa = campo[14].ToString();
                    daoPosicaoOp.EmEstoque = Convert.ToDecimal(campo[15].ToString());
                    daoPosicaoOp.EmProducao = Convert.ToDecimal(campo[16].ToString());
                    daoPosicaoOp.Necessidade = Convert.ToDecimal(campo[17].ToString());
                    daoPosicaoOp.Saldo = Convert.ToDecimal(campo[18].ToString());
                    daoPosicaoOp.Sobra = Convert.ToDecimal(campo[19].ToString());
                    daoPosicaoOp.Falta = Convert.ToDecimal(campo[20].ToString());
                    daoPosicaoOp.Vendido = Convert.ToDecimal(campo[21].ToString());
                    daoPosicaoOp.Faturado = Convert.ToDecimal(campo[22].ToString());
                    daoPosicaoOp.AFaturar = Convert.ToDecimal(campo[23].ToString());
                    daoPosicaoOp.Cancelado = Convert.ToDecimal(campo[24].ToString());
                    daoPosicaoOp.Bloqueado = campo[25].ToString();
                    daoPosicaoOp.Estoque = Convert.ToDecimal(campo[26].ToString());
                    daoPosicaoOp.ProntaEntrega = Convert.ToDecimal(campo[27].ToString());
                    daoPosicaoOp.Outros = Convert.ToDecimal(campo[28].ToString());


                    daoPosicaoOpList.Add(daoPosicaoOp);
                    
                }
            }
            return daoPosicaoOpList;
        }

        public string InserirDadosBD(DAOPosicaoOpList daoPosicaoOpList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspPosicaoOpDeletar");

            try
            {
                DataTable dataTablePosicaoOpList = ConvertToDataTable(daoPosicaoOpList);
                foreach (DataRow linha in dataTablePosicaoOpList.Rows)
                {
                    DAOPosicaoOp daoPosicaoOp = new DAOPosicaoOp();
                    daoPosicaoOp.Col = linha["Col"].ToString();
                    daoPosicaoOp.Colecao = linha["Colecao"].ToString();
                    daoPosicaoOp.Lin = linha["Lin"].ToString();
                    daoPosicaoOp.Linha = linha["Linha"].ToString();
                    daoPosicaoOp.Art = linha["Art"].ToString();
                    daoPosicaoOp.Artigo = linha["Artigo"].ToString();
                    daoPosicaoOp.Nivel = linha["Nivel"].ToString();
                    daoPosicaoOp.Grupo = linha["Grupo"].ToString();
                    daoPosicaoOp.Sub = linha["Sub"].ToString();
                    daoPosicaoOp.Cor = linha["Cor"].ToString();
                    daoPosicaoOp.Produto = linha["Produto"].ToString();
                    daoPosicaoOp.NomeGrupo = linha["NomeGrupo"].ToString();
                    daoPosicaoOp.NomeSub = linha["NomeSub"].ToString();
                    daoPosicaoOp.NomeCor = linha["NomeCor"].ToString();
                    daoPosicaoOp.Narrativa = linha["Narrativa"].ToString();
                    daoPosicaoOp.EmEstoque = Convert.ToDecimal(linha["EmEstoque"].ToString());
                    daoPosicaoOp.EmProducao = Convert.ToDecimal(linha["EmProducao"].ToString());
                    daoPosicaoOp.Necessidade = Convert.ToDecimal(linha["Necessidade"].ToString());
                    daoPosicaoOp.Saldo = Convert.ToDecimal(linha["Saldo"].ToString());
                    daoPosicaoOp.Sobra = Convert.ToDecimal(linha["Sobra"].ToString());
                    daoPosicaoOp.Falta = Convert.ToDecimal(linha["Falta"].ToString());
                    daoPosicaoOp.Vendido = Convert.ToDecimal(linha["Vendido"].ToString());
                    daoPosicaoOp.Faturado = Convert.ToDecimal(linha["Faturado"].ToString());
                    daoPosicaoOp.AFaturar = Convert.ToDecimal(linha["AFaturar"].ToString());
                    daoPosicaoOp.Cancelado = Convert.ToDecimal(linha["Cancelado"].ToString());
                    daoPosicaoOp.Bloqueado = linha["Bloqueado"].ToString();
                    daoPosicaoOp.Estoque = Convert.ToDecimal(linha["Estoque"].ToString());
                    daoPosicaoOp.ProntaEntrega = Convert.ToDecimal(linha["ProntaEntrega"].ToString());
                    daoPosicaoOp.Outros = Convert.ToDecimal(linha["Outros"].ToString());

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Emp", daoPosicaoOp.Emp);
                    dalMySQL.AdicionaParametros("@Col", daoPosicaoOp.Col);
                    dalMySQL.AdicionaParametros("@Colecao", daoPosicaoOp.Colecao);
                    dalMySQL.AdicionaParametros("@Lin", daoPosicaoOp.Lin);
                    dalMySQL.AdicionaParametros("@Linha", daoPosicaoOp.Linha);
                    dalMySQL.AdicionaParametros("@Art", daoPosicaoOp.Art);
                    dalMySQL.AdicionaParametros("@Artigo", daoPosicaoOp.Artigo);
                    dalMySQL.AdicionaParametros("@Nivel", daoPosicaoOp.Nivel);
                    dalMySQL.AdicionaParametros("@Grupo", daoPosicaoOp.Grupo);
                    dalMySQL.AdicionaParametros("@Sub", daoPosicaoOp.Sub);
                    dalMySQL.AdicionaParametros("@Cor", daoPosicaoOp.Cor);
                    dalMySQL.AdicionaParametros("@Produto", daoPosicaoOp.Produto);
                    dalMySQL.AdicionaParametros("@NomeGrupo", daoPosicaoOp.NomeGrupo);
                    dalMySQL.AdicionaParametros("@NomeSub", daoPosicaoOp.NomeSub);
                    dalMySQL.AdicionaParametros("@NomeCor", daoPosicaoOp.NomeCor);
                    dalMySQL.AdicionaParametros("@Narrativa", daoPosicaoOp.Narrativa);
                    dalMySQL.AdicionaParametros("@EmEstoque", daoPosicaoOp.EmEstoque);
                    dalMySQL.AdicionaParametros("@EmProducao", daoPosicaoOp.EmProducao);
                    dalMySQL.AdicionaParametros("@Necessidade", daoPosicaoOp.Necessidade);
                    dalMySQL.AdicionaParametros("@Saldo", daoPosicaoOp.Saldo);
                    dalMySQL.AdicionaParametros("@Sobra", daoPosicaoOp.Sobra);
                    dalMySQL.AdicionaParametros("@Falta", daoPosicaoOp.Falta);
                    dalMySQL.AdicionaParametros("@Vendido", daoPosicaoOp.Vendido);
                    dalMySQL.AdicionaParametros("@Faturado", daoPosicaoOp.Faturado);
                    dalMySQL.AdicionaParametros("@AFaturar", daoPosicaoOp.AFaturar);
                    dalMySQL.AdicionaParametros("@Cancelado", daoPosicaoOp.Cancelado);
                    dalMySQL.AdicionaParametros("@Bloqueado", daoPosicaoOp.Bloqueado);
                    dalMySQL.AdicionaParametros("@Estoque", daoPosicaoOp.Estoque);
                    dalMySQL.AdicionaParametros("@ProntaEntrega", daoPosicaoOp.ProntaEntrega);
                    dalMySQL.AdicionaParametros("@Outros", daoPosicaoOp.Outros);
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: posicao das ops inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir posicao das ops Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de posicao das ops deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de posicao das ops renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
