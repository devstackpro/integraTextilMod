using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLEstoqueRolos
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de Estoque rolos encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de Estoque rolos não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Estoque rolos e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear Estoque rolos. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOEstoqueRolosList LerCsv(string path)
        {
            DAOEstoqueRolosList daoEstoqueRolosList = new DAOEstoqueRolosList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOEstoqueRolos daoEstoqueRolos = new DAOEstoqueRolos();
                campo = linha.Split(';');
                index++;

                if (index > 1)
                {
                    daoEstoqueRolos.Empresa = campo[0].ToString();
                    daoEstoqueRolos.Dep = campo[1].ToString();
                    daoEstoqueRolos.Deposito = campo[2].ToString();
                    daoEstoqueRolos.DepO = campo[3].ToString();
                    daoEstoqueRolos.DepositoOriginal = campo[4].ToString();
                    daoEstoqueRolos.Tpg = campo[5].ToString();
                    daoEstoqueRolos.TipoProduto = campo[6].ToString();
                    daoEstoqueRolos.Tran = Convert.ToInt32(campo[7].ToString());
                    daoEstoqueRolos.Transacao = campo[8].ToString();
                    daoEstoqueRolos.AtualizaEstoque = campo[9].ToString();
                    daoEstoqueRolos.TipoTransacao = campo[10].ToString();
                    daoEstoqueRolos.CodigoRolo = Convert.ToInt32(campo[11].ToString());
                    daoEstoqueRolos.RoloOrigem = Convert.ToInt32(campo[12].ToString());
                    daoEstoqueRolos.Sit = Convert.ToInt32(campo[13].ToString());
                    daoEstoqueRolos.Situacao = campo[14].ToString();
                    daoEstoqueRolos.Produto = campo[15].ToString();
                    daoEstoqueRolos.Nivel = campo[16].ToString();
                    daoEstoqueRolos.Grupo = campo[17].ToString();
                    daoEstoqueRolos.Sub = campo[18].ToString();
                    daoEstoqueRolos.Item = campo[19].ToString();
                    daoEstoqueRolos.Um = campo[20].ToString();
                    daoEstoqueRolos.Narrativa = campo[21].ToString();
                    daoEstoqueRolos.Lote = campo[22].ToString();
                    daoEstoqueRolos.LoteProduto = campo[23].ToString();
                    daoEstoqueRolos.Quantidade = Convert.ToDecimal(campo[24].ToString());
                    daoEstoqueRolos.PesoBruto = Convert.ToDecimal(campo[25].ToString());
                    daoEstoqueRolos.PesoLiquido = Convert.ToDecimal(campo[26].ToString());
                    daoEstoqueRolos.DataEntrada = campo[27].ToString();
                    daoEstoqueRolos.Op = Convert.ToInt32(campo[28].ToString());
                    daoEstoqueRolos.InicioProdução = campo[29].ToString();
                    daoEstoqueRolos.TerminoProdução = campo[30].ToString();
                    daoEstoqueRolos.GrMaquina = campo[31].ToString();
                    daoEstoqueRolos.SbMaquina = campo[32].ToString();
                    daoEstoqueRolos.NrMaquina = campo[33].ToString();
                    daoEstoqueRolos.Maquina = campo[34].ToString();
                    daoEstoqueRolos.Operador = Convert.ToInt32(campo[35].ToString());
                    daoEstoqueRolos.NomeOperador = campo[36].ToString();
                    daoEstoqueRolos.Nuance = campo[37].ToString();
                    daoEstoqueRolos.Qualidade = Convert.ToInt32(campo[38].ToString());
                    daoEstoqueRolos.Largura = Convert.ToDecimal(campo[39].ToString());
                    daoEstoqueRolos.Gramatura = Convert.ToDecimal(campo[40].ToString());
                    daoEstoqueRolos.Turno = Convert.ToInt32(campo[41].ToString());
                    daoEstoqueRolos.Pontuacao = campo[42].ToString();
                    daoEstoqueRolos.Pedido = Convert.ToInt32(campo[43].ToString());
                    daoEstoqueRolos.SEqPedido = campo[44].ToString();
                    daoEstoqueRolos.Nf = campo[45].ToString();
                    daoEstoqueRolos.Serie = Convert.ToInt32(campo[46].ToString());
                    daoEstoqueRolos.SeqNf = campo[47].ToString();
                    daoEstoqueRolos.UsuarioCardex = campo[48].ToString();
                    daoEstoqueRolos.DataInsercao = campo[49].ToString();
                    daoEstoqueRolos.Tabela = campo[50].ToString();
                    daoEstoqueRolos.Programa = campo[51].ToString();
                    daoEstoqueRolos.EnderecoRolo = Convert.ToInt32(campo[52].ToString());
                    daoEstoqueRolos.CodigoTecelao = campo[53].ToString();
                    daoEstoqueRolos.Romaneio = Convert.ToInt32(campo[54].ToString());
                    daoEstoqueRolos.Batidas = Convert.ToDecimal(campo[55].ToString());
                    daoEstoqueRolos.BatidasMinuto = Convert.ToDecimal(campo[56].ToString());
                    daoEstoqueRolos.BatidasRolo = Convert.ToDecimal(campo[57].ToString());

                    daoEstoqueRolosList.Add(daoEstoqueRolos);
                }
            }
            return daoEstoqueRolosList;
        }

        public string InserirDadosBD(DAOEstoqueRolosList daoEstoqueRolosList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspEstoqueRolosDeletar");

            try
            {
                DataTable dataTableEstoqueRolosList = ConvertToDataTable(daoEstoqueRolosList);
                foreach (DataRow linha in dataTableEstoqueRolosList.Rows)
                {
                    DAOEstoqueRolos daoEstoqueRolos = new DAOEstoqueRolos();

                    daoEstoqueRolos.Dep = linha["Dep"].ToString();
                    daoEstoqueRolos.Deposito = linha["Deposito"].ToString();
                    daoEstoqueRolos.DepO = linha["DepO"].ToString();
                    daoEstoqueRolos.DepositoOriginal = linha["DepositoOriginal"].ToString();
                    daoEstoqueRolos.Tpg = linha["Tpg"].ToString();
                    daoEstoqueRolos.TipoProduto = linha["TipoProduto"].ToString();
                    daoEstoqueRolos.Tran = linha["Tran"].ToString();
                    daoEstoqueRolos.Transacao = linha["Transacao"].ToString();
                    daoEstoqueRolos.AtualizaEstoque = linha["AtualizaEstoque"].ToString();
                    daoEstoqueRolos.TipoTransacao = linha["TipoTransacao"].ToString();
                    daoEstoqueRolos.CodigoRolo = linha["CodigoRolo"].ToString();
                    daoEstoqueRolos.RoloOrigem = linha["RoloOrigem"].ToString();
                    daoEstoqueRolos.Sit = linha["Sit"].ToString();;
                    daoEstoqueRolos.Situacao = linha["Situacao"].ToString();
                    daoEstoqueRolos.Produto = linha["Produto"].ToString();
                    daoEstoqueRolos.Nivel = linha["Nivel"].ToString();
                    daoEstoqueRolos.Grupo = linha["Grupo"].ToString();
                    daoEstoqueRolos.Sub = linha["Sub"].ToString();
                    daoEstoqueRolos.Item = linha["Item"].ToString();
                    daoEstoqueRolos.Um = linha["Um"].ToString();
                    daoEstoqueRolos.Narrativa = linha["Narrativa"].ToString();
                    daoEstoqueRolos.Lote = linha["Lote"].ToString();
                    daoEstoqueRolos.LoteProduto = linha["LoteProduto"].ToString();
                    daoEstoqueRolos.Quantidade = linha["Quantidade"].ToString();
                    daoEstoqueRolos.PesoBruto = linha["PesoBruto"].ToString();
                    daoEstoqueRolos.PesoLiquido = linha["PesoLiquido"].ToString();
                    daoEstoqueRolos.DataEntrada = linha["DataEntrada"].ToString();
                    daoEstoqueRolos.Op = linha["Op"].ToString();
                    daoEstoqueRolos.InicioProdução = linha["InicioProdução"].ToString();
                    daoEstoqueRolos.TerminoProdução = linha["TerminoProdução"].ToString();
                    daoEstoqueRolos.GrMaquina = linha["GrMaquina"].ToString();
                    daoEstoqueRolos.SbMaquina = linha["SbMaquina"].ToString();
                    daoEstoqueRolos.NrMaquina = linha["NrMaquina"].ToString();
                    daoEstoqueRolos.Maquina = linha["Maquina"].ToString();
                    daoEstoqueRolos.Operador = linha["Operador"].ToString();
                    daoEstoqueRolos.NomeOperador = linha["NomeOperador"].ToString();
                    daoEstoqueRolos.Nuance = linha["Nuance"].ToString();
                    daoEstoqueRolos.Qualidade = linha["Qualidade"].ToString();
                    daoEstoqueRolos.Largura = linha["Largura"].ToString();
                    daoEstoqueRolos.Gramatura = linha["Gramatura"].ToString();
                    daoEstoqueRolos.Turno = linha["Turno"].ToString();
                    daoEstoqueRolos.Pontuacao = linha["Pontuacao"].ToString();
                    daoEstoqueRolos.Pedido = linha["Pedido"].ToString();
                    daoEstoqueRolos.SEqPedido = linha["SEqPedido"].ToString();
                    daoEstoqueRolos.Nf = linha["Nf"].ToString();
                    daoEstoqueRolos.Serie = linha["Serie"].ToString();
                    daoEstoqueRolos.SeqNf = linha["SeqNf"].ToString();
                    daoEstoqueRolos.UsuarioCardex = linha["UsuarioCardex"].ToString();
                    daoEstoqueRolos.DataInsercao = linha["DataInsercao"].ToString();
                    daoEstoqueRolos.Tabela = linha["Tabela"].ToString();
                    daoEstoqueRolos.Programa = linha["Programa"].ToString();
                    daoEstoqueRolos.EnderecoRolo = linha["EnderecoRolo"].ToString();
                    daoEstoqueRolos.CodigoTecelao = linha["CodigoTecelao"].ToString();
                    daoEstoqueRolos.Romaneio = linha["Romaneio"].ToString();
                    daoEstoqueRolos.Batidas = linha["Batidas"].ToString();
                    daoEstoqueRolos.BatidasMinuto = linha["BatidasMinuto"].ToString();
                    daoEstoqueRolos.BatidasRolo = linha["BatidasRolo"].ToString();

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Emp", daoEstoqueRolos.Emp);
                    dalMySQL.AdicionaParametros("@Empresa", daoEstoqueRolos.Empresa);
                    dalMySQL.AdicionaParametros("@Dep", daoEstoqueRolos.Dep);
                    dalMySQL.AdicionaParametros("@Deposito", daoEstoqueRolos.Deposito);
                    dalMySQL.AdicionaParametros("@Tran", daoEstoqueRolos.Tran);
                    dalMySQL.AdicionaParametros("@Transacao", daoEstoqueRolos.Transacao);
                    dalMySQL.AdicionaParametros("@Produto", daoEstoqueRolos.Produto);
                    dalMySQL.AdicionaParametros("@Nivel", daoEstoqueRolos.Nivel);
                    dalMySQL.AdicionaParametros("@Grupo", daoEstoqueRolos.Grupo);
                    dalMySQL.AdicionaParametros("@Sub", daoEstoqueRolos.Sub);
                    dalMySQL.AdicionaParametros("@Item", daoEstoqueRolos.Item);
                    dalMySQL.AdicionaParametros("@Um", daoEstoqueRolos.Um);
                    dalMySQL.AdicionaParametros("@Narrativa", daoEstoqueRolos.Narrativa);
                    dalMySQL.AdicionaParametros("@Tpg", daoEstoqueRolos.Tpg);
                    dalMySQL.AdicionaParametros("@TipoGlobal", daoEstoqueRolos.TipoGlobal);
                    dalMySQL.AdicionaParametros("@Lote", daoEstoqueRolos.Lote);
                    dalMySQL.AdicionaParametros("@LoteProduto", daoEstoqueRolos.LoteProduto);
                    dalMySQL.AdicionaParametros("@Quantidade", daoEstoqueRolos.Quantidade);
                    dalMySQL.AdicionaParametros("@PesoBruto", daoEstoqueRolos.PesoBruto);
                    dalMySQL.AdicionaParametros("@PesoLiquido", daoEstoqueRolos.PesoLiquido);
                    dalMySQL.AdicionaParametros("@NumeroVolume", daoEstoqueRolos.NumeroVolume);
                    dalMySQL.AdicionaParametros("@NumeroOrigem", daoEstoqueRolos.NumeroOrigem);
                    dalMySQL.AdicionaParametros("@Situacao", daoEstoqueRolos.Situacao);
                    dalMySQL.AdicionaParametros("@DataEntrada", daoEstoqueRolos.DataEntrada);
                    dalMySQL.AdicionaParametros("@Nf", daoEstoqueRolos.Nf);
                    dalMySQL.AdicionaParametros("@Serie", daoEstoqueRolos.Serie);
                    dalMySQL.AdicionaParametros("@seqNf", daoEstoqueRolos.seqNf);
                    dalMySQL.AdicionaParametros("@Fornecedor", daoEstoqueRolos.Fornecedor);
                    dalMySQL.AdicionaParametros("@NomeFornecedor", daoEstoqueRolos.NomeFornecedor);
                    dalMySQL.AdicionaParametros("@Obs", daoEstoqueRolos.Obs);


                    dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspEstoqueRolosInserir");
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Estoque rolos inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir Estoque rolos. Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de Estoque rolos deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de Estoque rolos renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
