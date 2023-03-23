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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de posicao das ops encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
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
                    daoPosicaoOp.MaquinaOp = campo[0].ToString();
                    daoPosicaoOp.Op = campo[1].ToString();
                    daoPosicaoOp.Est = campo[2].ToString();
                    daoPosicaoOp.Estagio = campo[3].ToString();
                    daoPosicaoOp.Processo = campo[4].ToString();
                    daoPosicaoOp.Produto = campo[5].ToString();
                    daoPosicaoOp.Tpg = campo[6].ToString();
                    daoPosicaoOp.Um = campo[7].ToString();
                    daoPosicaoOp.Narrativa = campo[8].ToString();
                    daoPosicaoOp.Nivel = campo[9].ToString();
                    daoPosicaoOp.Grupo = campo[10].ToString();
                    daoPosicaoOp.Sub = campo[11].ToString();
                    daoPosicaoOp.Item = campo[12].ToString();
                    daoPosicaoOp.EstagioPosicao = campo[13].ToString();
                    daoPosicaoOp.QtdeProgramado = Convert.ToDecimal(campo[14].ToString());
                    daoPosicaoOp.QtdeCarregado = Convert.ToDecimal(campo[15].ToString());
                    daoPosicaoOp.QtdeProduzida = Convert.ToDecimal(campo[16].ToString());
                    daoPosicaoOp.MaquinaReal = campo[17].ToString();
                    daoPosicaoOp.NomeMaquinaReal = campo[18].ToString();
                    daoPosicaoOp.InicioReal = campo[19].ToString();
                    daoPosicaoOp.TerminoReal = campo[20].ToString();
                    daoPosicaoOp.Obs = campo[21].ToString();


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
                    daoPosicaoOp.MaquinaOp = linha["MaquinaOp"].ToString();
                    daoPosicaoOp.Op = linha["Op"].ToString();
                    daoPosicaoOp.Est = linha["Est"].ToString();
                    daoPosicaoOp.Estagio = linha["Estagio"].ToString();
                    daoPosicaoOp.Processo = linha["Processo"].ToString();
                    daoPosicaoOp.Produto = linha["Produto"].ToString();
                    daoPosicaoOp.Tpg = linha["Tpg"].ToString();
                    daoPosicaoOp.Um = linha["Um"].ToString();
                    daoPosicaoOp.Narrativa = linha["Narrativa"].ToString();
                    daoPosicaoOp.Nivel = linha["Nivel"].ToString();
                    daoPosicaoOp.Grupo = linha["Grupo"].ToString();
                    daoPosicaoOp.Sub = linha["Sub"].ToString();
                    daoPosicaoOp.Item = linha["Item"].ToString();
                    daoPosicaoOp.EstagioPosicao = linha["EstagioPosicao"].ToString();
                    daoPosicaoOp.QtdeProgramado = Convert.ToDecimal(linha["QtdeProgramado"].ToString());
                    daoPosicaoOp.QtdeCarregado = Convert.ToDecimal(linha["QtdeCarregado"].ToString());
                    daoPosicaoOp.QtdeProduzida = Convert.ToDecimal(linha["QtdeProduzida"].ToString());
                    daoPosicaoOp.MaquinaReal = linha["MaquinaReal"].ToString();
                    daoPosicaoOp.NomeMaquinaReal = linha["NomeMaquinaReal"].ToString();
                    daoPosicaoOp.InicioReal = linha["InicioReal"].ToString();
                    daoPosicaoOp.TerminoReal = linha["TerminoReal"].ToString();
                    daoPosicaoOp.Obs= linha["Obs"].ToString();


                    dalMySQL.LimparParametros();

                        dalMySQL.AdicionaParametros("@MaquinaOp", daoPosicaoOp.MaquinaOp);
                        dalMySQL.AdicionaParametros("@Op", daoPosicaoOp.Op);
                        dalMySQL.AdicionaParametros("@Est", daoPosicaoOp.Est);
                        dalMySQL.AdicionaParametros("@Estagio", daoPosicaoOp.Estagio);
                        dalMySQL.AdicionaParametros("@Processo", daoPosicaoOp.Processo);
                        dalMySQL.AdicionaParametros("@Produto", daoPosicaoOp.Produto);
                        dalMySQL.AdicionaParametros("@Tpg", daoPosicaoOp.Tpg);
                        dalMySQL.AdicionaParametros("@Um", daoPosicaoOp.Um);
                        dalMySQL.AdicionaParametros("@Narrativa", daoPosicaoOp.Narrativa);
                        dalMySQL.AdicionaParametros("@Nivel", daoPosicaoOp.Nivel);
                        dalMySQL.AdicionaParametros("@Grupo", daoPosicaoOp.Grupo);
                        dalMySQL.AdicionaParametros("@Sub", daoPosicaoOp.Sub);
                        dalMySQL.AdicionaParametros("@Item", daoPosicaoOp.Item);
                        dalMySQL.AdicionaParametros("@EstagioPosicao", daoPosicaoOp.EstagioPosicao);
                        dalMySQL.AdicionaParametros("@QtdeProgramado", daoPosicaoOp.QtdeProgramado);
                        dalMySQL.AdicionaParametros("@QtdeCarregado", daoPosicaoOp.EQtdeCarregadomp);
                        dalMySQL.AdicionaParametros("@QtdeProduzida", daoPosicaoOp.QtdeProduzida);
                        dalMySQL.AdicionaParametros("@MaquinaReal", daoPosicaoOp.MaquinaReal);
                        dalMySQL.AdicionaParametros("@NomeMaquinaReal", daoPosicaoOp.NomeMaquinaReal);
                        dalMySQL.AdicionaParametros("@InicioReal", daoPosicaoOp.InicioReal);
                        dalMySQL.AdicionaParametros("@TerminoReal", daoPosicaoOp.TerminoReal);
                        dalMySQL.AdicionaParametros("@Obs", daoPosicaoOp.Obs);

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
