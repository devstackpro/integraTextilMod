using System.Text;
using System.ComponentModel;
using System.Data;
using DAL;
using DAO;

namespace BLL
{
    public class BLLMovimentacaoEstoques
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de movimentação de estoques encontrado, nome: " + arquivoNome + ".  Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + data);
                }

            }
            catch (Exception ex)
            {
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: Relatório de movimentação de estoques não encontrado, nome: nulo. Detalhes: Classe: BLLContasPagas.cs | Metodo: PegarNomeArquivo | " + ex.Message.ToString() + " | " + data);
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
                    bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: movimentação de estoques e renomeado para pasta destino. Detalhes: " + retorno + " | " + data);
                }

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível mover e renomear movimentação de estoques. Detalhes: " + retorno + " | " + data);
            }

            return retorno;


        }

        public DAOMovimentacaoEstoquesList LerCsv(string path)
        {
            DAOMovimentacaoEstoquesList daoMovimentacaoEstoquesList = new DAOMovimentacaoEstoquesList();
            var csv = new StreamReader(File.OpenRead(path));
            string linha;
            string[] campo;
            int index = 0;

            while ((linha = csv.ReadLine()) != null)
            {
                DAOMovimentacaoEstoques daoMovimentacaoEstoques = new DAOMovimentacaoEstoques();   
                campo = linha.Split(';');
                index++;

                if (index > 1)
                {
                    daoMovimentacaoEstoques.Emp = campo[0].ToString();
                    daoMovimentacaoEstoques.Empresa = campo[1].ToString();
                    daoMovimentacaoEstoques.Dep = campo[2].ToString();
                    daoMovimentacaoEstoques.Deposito = campo[3].ToString();
                    daoMovimentacaoEstoques.Nivel = campo[4].ToString();
                    daoMovimentacaoEstoques.Grupo = campo[5].ToString();
                    daoMovimentacaoEstoques.Sub = campo[6].ToString();
                    daoMovimentacaoEstoques.Cor = campo[7].ToString();
                    daoMovimentacaoEstoques.Produto = campo[8].ToString();
                    daoMovimentacaoEstoques.Um = campo[9].ToString();
                    daoMovimentacaoEstoques.CodigoBarras = campo[10].ToString();
                    daoMovimentacaoEstoques.CodigoVelho = campo[11].ToString();
                    daoMovimentacaoEstoques.NomeGrupo = campo[12].ToString();
                    daoMovimentacaoEstoques.NomeSub = campo[13].ToString();
                    daoMovimentacaoEstoques.NomeCor = campo[14].ToString();
                    daoMovimentacaoEstoques.Narrativa = campo[15].ToString();
                    daoMovimentacaoEstoques.Cf = campo[16].ToString();
                    daoMovimentacaoEstoques.Col = campo[17].ToString();
                    daoMovimentacaoEstoques.Lin = campo[18].ToString();
                    daoMovimentacaoEstoques.Linha = Convert.ToDecimal(campo[19].ToString());
                    daoMovimentacaoEstoques.Art = campo[20].ToString();
                    daoMovimentacaoEstoques.Cota = Convert.ToDecimal[21].ToString();
                    daoMovimentacaoEstoques.ArtigoCotas = campo[22].ToString();
                    daoMovimentacaoEstoques.Ces = campo[23].ToString();
                    daoMovimentacaoEstoques.ContaEstoque = campo[24].ToString();
                    daoMovimentacaoEstoques.Tpg = campo[25].ToString();
                    daoMovimentacaoEstoques.TipoProdutoGlobal = campo[26].ToString();
                    daoMovimentacaoEstoques.TprogTpg = campo[27].ToString();
                    daoMovimentacaoEstoques.NivTpg = campo[28].ToString();
                    daoMovimentacaoEstoques.EstTpg = campo[29].ToString();
                    daoMovimentacaoEstoques.Cliente = campo[30].ToString();
                    daoMovimentacaoEstoques.NomeCliente = campo[31].ToString();
                    daoMovimentacaoEstoques.Marca = campo[32].ToString();
                    daoMovimentacaoEstoques.NomeMarca = campo[33].ToString();
                    daoMovimentacaoEstoques.TipoTecido = campo[34].ToString();
                    daoMovimentacaoEstoques.Tpm = campo[35].ToString();
                    daoMovimentacaoEstoques.Ncm = campo[36].ToString();
                    daoMovimentacaoEstoques.Altp = campo[37].ToString();
                    daoMovimentacaoEstoques.Rotp = campo[38].ToString();
                    daoMovimentacaoEstoques.Antc  = campo[39].ToString(); 
                    daoMovimentacaoEstoques.Rotc = campo[40].ToString();
                    daoMovimentacaoEstoques.ValorMedioEstoque = Convert.ToDecimal[41].ToString();
                    daoMovimentacaoEstoques.ValorUltimaCopmpra = Convert.ToDecimal[42].ToString();
                    daoMovimentacaoEstoques.Custo = Convert.ToDecimal[43].ToString();
                    daoMovimentacaoEstoques.CustoInformado = Convert.ToDecimal[44].ToString();
                    daoMovimentacaoEstoques.Lead = campo[45].ToString();
                    daoMovimentacaoEstoques.FamiliaTear = campo[46].ToString();
                    daoMovimentacaoEstoques.LoteTam  = campo[47].ToString();
                    daoMovimentacaoEstoques.PesoLiquido = Convert.ToDecimal[48].ToString();
                    daoMovimentacaoEstoques.PesoRolo = Convert.ToDecimal[49].ToString();
                    daoMovimentacaoEstoques.PesoMinRolo = Convert.ToDecimal[50].ToString();
                    daoMovimentacaoEstoques.DescTamFicha = campo[51].ToString();
                    daoMovimentacaoEstoques.TipoProdQuimico = campo[52].ToString();
                    daoMovimentacaoEstoques.ItemAtivo = campo[53].ToString();
                    daoMovimentacaoEstoques.CodigoContabil = campo[54].ToString();
                    daoMovimentacaoEstoques.CodProcesso = campo[55].ToString();
                    daoMovimentacaoEstoques.Lote = campo[56].ToString();
                    daoMovimentacaoEstoques.LoteProduto = campo[57].ToString();
                    daoMovimentacaoEstoques.SaldoAtual = campo[58].ToString();
                    daoMovimentacaoEstoques.Volumes = Convert.ToInt32(campo[59].ToString());
                    daoMovimentacaoEstoques.QtEstqInicioMes = Convert.ToDecimal[60].ToString();
                    daoMovimentacaoEstoques.QtEstqFinalMes = Convert.ToDecimal[61].ToString();
                    daoMovimentacaoEstoques.UltimaEntrada = Convert.ToDateTime[62].ToString();
                    daoMovimentacaoEstoques.UltimaSaida = Convert.ToDateTime[63].ToString();
                    daoMovimentacaoEstoques.QtSugerida = Convert.ToDecimal[64].ToString();
                    daoMovimentacaoEstoques.QtEmpenhada = Convert.ToDecimal[65].ToString(); 
                    daoMovimentacaoEstoques.CnpjFornecedor = campo[66].ToString();
                    daoMovimentacaoEstoques.NotaFiscal = campo[67].ToString();
                    daoMovimentacaoEstoques.PeriodoEstoque = campo[68].ToString();


                    daoMovimentacaoEstoquesList.Add(daoMovimentacaoEstoques);
                    
                }
            }
            return daoMovimentacaoEstoquesList;
        }

        public string InserirDadosBD(DAOMovimentacaoEstoquesList daoMovimentacaoEstoquesList)
        {
            BLLFerramentas bllFerramentas = new BLLFerramentas();
            string retorno = "";
            dalMySQL.LimparParametros();
            dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspMovimentacaoEstoquesDeletar");

            try
            {
                DataTable dataTableMovimentacaoEstoquesList = ConvertToDataTable(daoMovimentacaoEstoquesList);
                foreach (DataRow linha in dataTableMovimentacaoEstoquesList.Rows)
                {
                    DAOMovimentacaoEstoques daoMovimentacaoEstoques = new DAOMovimentacaoEstoques();

                    daoMovimentacaoEstoques.Emp = linha["Emp"].ToString();
                    daoMovimentacaoEstoques.Empresa = linha["Empresa"].ToString();
                    daoMovimentacaoEstoques.Dep = linha["Dep"].ToString();
                    daoMovimentacaoEstoques.Deposito = linha["Deposito"].ToString();
                    daoMovimentacaoEstoques.Nivel = linha["Nivel"].ToString();
                    daoMovimentacaoEstoques.Grupo = linha["Grupo"].ToString();
                    daoMovimentacaoEstoques.Sub = linha["Sub"].ToString();
                    daoMovimentacaoEstoques.Cor = linha["Cor"].ToString();
                    daoMovimentacaoEstoques.Produto = linha["Produto"].ToString();
                    daoMovimentacaoEstoques.Um = linha["Um"].ToString();
                    daoMovimentacaoEstoques.CodigoBarras = linha["CodigoBarras"].ToString();
                    daoMovimentacaoEstoques.CodigoVelho  = linha["CodigoVelho "].ToString();
                    daoMovimentacaoEstoques.NomeGrupo = linha["NomeGrupo"].ToString();
                    daoMovimentacaoEstoques.NomeSub = linha["NomeSub"].ToString();
                    daoMovimentacaoEstoques.NomeCor = linha["NomeCor"].ToString();
                    daoMovimentacaoEstoques.Narrativa = linha["Narrativa"].ToString();
                    daoMovimentacaoEstoques.Col = linha["Col"].ToString();
                    daoMovimentacaoEstoques.Lin = linha["Lin"].ToString();
                    daoMovimentacaoEstoques.Linha = linha["Linha"].ToString();
                    daoMovimentacaoEstoques.Art = linha["Art"].ToString();
                    daoMovimentacaoEstoques.Cota = linha["Cota"].ToString();
                    daoMovimentacaoEstoques.ArtigoCotas = linha["ArtigoCotas"].ToString();
                    daoMovimentacaoEstoques.Ces = linha["Ces"].ToString();
                    daoMovimentacaoEstoques.ContaEstoque = linha["ContaEstoque"].ToString();
                    daoMovimentacaoEstoques.Tpg = linha["Tpg"].ToString();
                    daoMovimentacaoEstoques.TipoProdutoGlobal = linha["TipoProdutoGlobal"].ToString();
                    daoMovimentacaoEstoques.TprogTpg = linha["TprogTpg"].ToString();
                    daoMovimentacaoEstoques.NivTpg = linha["NivTpg"].ToString();
                    daoMovimentacaoEstoques.EstTpg = linha["EstTpg"].ToString();
                    daoMovimentacaoEstoques.Cliente = linha["Cliente"].ToString();
                    daoMovimentacaoEstoques.NomeCliente = linha["NomeCliente"].ToString();
                    daoMovimentacaoEstoques.Marca = linha["Marca"].ToString();
                    daoMovimentacaoEstoques.NomeMarca  = linha["NomeMarca "].ToString();
                    daoMovimentacaoEstoques.TipoTecido = linha["TipoTecido"].ToString();
                    daoMovimentacaoEstoques.Tpm = linha["Tpm"].ToString();
                    daoMovimentacaoEstoques.Ncm  = linha["Ncm"].ToString();
                    daoMovimentacaoEstoques.Altp  = Convert.ToDecimal(linha["Altp "].ToString());
                    daoMovimentacaoEstoques.Rotp = linha["Rotp"].ToString();
                    daoMovimentacaoEstoques.Antc = linha["Antc"].ToString();
                    daoMovimentacaoEstoques.Rotc = linha["Rotc"].ToString();
                    daoMovimentacaoEstoques.ValorMedioEstoque = Convert.ToDecimal(linha["ValorMedioEstoque"].ToString());
                    daoMovimentacaoEstoques.ValorUltimaCopmpra = Convert.ToDecimal(linha["ValorUltimaCopmpra"].ToString());
                    daoMovimentacaoEstoques.Custo = Convert.ToDecimal(linha["Custo"].ToString());
                    daoMovimentacaoEstoques.CustoInformado = Convert.ToDecimal(linha["CustoInformado"].ToString());
                    daoMovimentacaoEstoques.Lead = linha["Lead"].ToString();
                    daoMovimentacaoEstoques.FamiliaTear = linha["FamiliaTear"].ToString();
                    daoMovimentacaoEstoques.LoteTam = linha["LoteTam"].ToString();
                    daoMovimentacaoEstoques.PesoLiquido = Convert.ToDecimal(linha["PesoLiquido"].ToString());
                    daoMovimentacaoEstoques.PesoRolo = Convert.ToDecimal(linha["PesoRolo"].ToString());
                    daoMovimentacaoEstoques.PesoMinRolo = Convert.ToDecimal(linha["PesoMinRolo"].ToString());
                    daoMovimentacaoEstoques.DescTamFicha = linha["DescTamFicha"].ToString();
                    daoMovimentacaoEstoques.TipoProdQuimico = linha["TipoProdQuimico"].ToString();
                    daoMovimentacaoEstoques.ItemAtivo = linha["ItemAtivo"].ToString();
                    daoMovimentacaoEstoques.CodigoContabil = linha["CodigoContabil"].ToString();
                    daoMovimentacaoEstoques.CodProcesso = linha["CodProcesso"].ToString();
                    daoMovimentacaoEstoques.Lote = linha["Lote"].ToString();
                    daoMovimentacaoEstoques.LoteProduto = linha["LoteProduto"].ToString();
                    daoMovimentacaoEstoques.SaldoAtual = linha["SaldoAtual"].ToString();
                    daoMovimentacaoEstoques.Volumes = Convert.ToInt32(linha["Volumes"].ToString());
                    daoMovimentacaoEstoques.QtEstqInicioMes = Convert.ToDecimal(linha["QtEstqInicioMes"].ToString());
                    daoMovimentacaoEstoques.QtEstqFinalMes = Convert.ToDecimal(linha["QtEstqFinalMes"].ToString());
                    daoMovimentacaoEstoques.UltimaEntrada = Convert.ToDateTime(linha["UltimaEntrada"].ToString());
                    daoMovimentacaoEstoques.UltimaSaida = Convert.ToDateTime(linha["UltimaSaida"].ToString());
                    daoMovimentacaoEstoques.QtSugerida = Convert.ToDecimal(linha["QtSugerida"].ToString());
                    daoMovimentacaoEstoques.QtEmpenhada = Convert.ToDecimal(linha["QtEmpenhada"].ToString());
                    daoMovimentacaoEstoques.CnpjFornecedor = linha["CnpjFornecedor"].ToString();
                    daoMovimentacaoEstoques.NotaFiscal = linha["NotaFiscal"].ToString();
                    daoMovimentacaoEstoques.PeriodoEstoque = linha["PeriodoEstoque"].ToString();

                    dalMySQL.LimparParametros();

                    dalMySQL.AdicionaParametros("@Emp", daoMovimentacaoEstoques.Emp);
                    dalMySQL.AdicionaParametros("@Empresa", daoMovimentacaoEstoques.Empresa);
                    dalMySQL.AdicionaParametros("@Dep", daoMovimentacaoEstoques.Dep );
                    dalMySQL.AdicionaParametros("@Deposito", daoMovimentacaoEstoques.Deposito);
                    dalMySQL.AdicionaParametros("@Nivel", daoMovimentacaoEstoques.Nivel);
                    dalMySQL.AdicionaParametros("@Grupo", daoMovimentacaoEstoques.Grupo);
                    dalMySQL.AdicionaParametros("@Sub", daoMovimentacaoEstoques.Sub);
                    dalMySQL.AdicionaParametros("@Nivel", daoMovimentacaoEstoques.Nivel);
                    dalMySQL.AdicionaParametros("@Cor", daoMovimentacaoEstoques.Cor);
                    dalMySQL.AdicionaParametros("@Produto", daoMovimentacaoEstoques.Produto);
                    dalMySQL.AdicionaParametros("@Um", daoMovimentacaoEstoques.Um);
                    dalMySQL.AdicionaParametros("@CodigoBarras", daoMovimentacaoEstoques.CodigoBarras);
                    dalMySQL.AdicionaParametros("@CodigoVelho", daoMovimentacaoEstoques.CodigoVelho);
                    dalMySQL.AdicionaParametros("@NomeGrupo", daoMovimentacaoEstoques.NomeGrupo);
                    dalMySQL.AdicionaParametros("@NomeSub", daoMovimentacaoEstoques.NomeSub);
                    dalMySQL.AdicionaParametros("@NomeCor", daoMovimentacaoEstoques.NomeCor);
                    dalMySQL.AdicionaParametros("@Narrativa", daoMovimentacaoEstoques.Narrativa);
                    dalMySQL.AdicionaParametros("@Cf", daoMovimentacaoEstoques.Cf);
                    dalMySQL.AdicionaParametros("@Col", daoMovimentacaoEstoques.Col);
                    dalMySQL.AdicionaParametros("@Lin", daoMovimentacaoEstoques.Lin);
                    dalMySQL.AdicionaParametros("@Linha", daoMovimentacaoEstoques.Linha);
                    dalMySQL.AdicionaParametros("@Art", daoMovimentacaoEstoques.Art);
                    dalMySQL.AdicionaParametros("@Cota", daoMovimentacaoEstoques.Cota);
                    dalMySQL.AdicionaParametros("@ArtigoCotas", daoMovimentacaoEstoques.ArtigoCotas);
                    dalMySQL.AdicionaParametros("@Ces", daoMovimentacaoEstoques.Ces);
                    dalMySQL.AdicionaParametros("@ContaEstoque", daoMovimentacaoEstoques.ContaEstoque);
                    dalMySQL.AdicionaParametros("@Tpg", daoMovimentacaoEstoques.Tpg);
                    dalMySQL.AdicionaParametros("@TipoProdutoGlobal", daoMovimentacaoEstoques.TprogTpg);
                    dalMySQL.AdicionaParametros("@TprogTpg", daoMovimentacaoEstoques.TprogTpg);
                    dalMySQL.AdicionaParametros("@NivTpg", daoMovimentacaoEstoques.NivTpg);
                    dalMySQL.AdicionaParametros("@EstTpg", daoMovimentacaoEstoques.EstTpg);
                    dalMySQL.AdicionaParametros("@Cliente", daoMovimentacaoEstoques.Cliente);
                    dalMySQL.AdicionaParametros("@NomeCliente", daoMovimentacaoEstoques.NomeCliente);
                    dalMySQL.AdicionaParametros("@Marca", daoMovimentacaoEstoques.Marca);
                    dalMySQL.AdicionaParametros("@NomeMarca", daoMovimentacaoEstoques.NomeMarca);
                    dalMySQL.AdicionaParametros("@TipoTecido", daoMovimentacaoEstoques.TipoTecido);
                    dalMySQL.AdicionaParametros("@Tpm", daoMovimentacaoEstoques.Tpm);
                    dalMySQL.AdicionaParametros("@Ncm", daoMovimentacaoEstoques.Ncm);
                    dalMySQL.AdicionaParametros("@Altp", daoMovimentacaoEstoques.Altp);
                    dalMySQL.AdicionaParametros("@Rotp", daoMovimentacaoEstoques.Rotp);
                    dalMySQL.AdicionaParametros("@Antc", daoMovimentacaoEstoques.Antc);
                    dalMySQL.AdicionaParametros("@Rotc", daoMovimentacaoEstoques.Rotc);
                    dalMySQL.AdicionaParametros("@ValorMedioEstoque", daoMovimentacaoEstoques.ValorMedioEstoque);
                    dalMySQL.AdicionaParametros("@ValorUltimaCopmpra", daoMovimentacaoEstoques.ValorUltimaCopmpra);
                    dalMySQL.AdicionaParametros("@Custo", daoMovimentacaoEstoques.Custo);
                    dalMySQL.AdicionaParametros("@CustoInformado", daoMovimentacaoEstoques.CustoInformado);
                    dalMySQL.AdicionaParametros("@Lead", daoMovimentacaoEstoques.Lead);
                    dalMySQL.AdicionaParametros("@FamiliaTear", daoMovimentacaoEstoques.FamiliaTear);
                    dalMySQL.AdicionaParametros("@LoteTam", daoMovimentacaoEstoques.LoteTam);
                    dalMySQL.AdicionaParametros("@PesoLiquido", daoMovimentacaoEstoques.PesoLiquido);
                    dalMySQL.AdicionaParametros("@PesoRolo", daoMovimentacaoEstoques.PesoRolo);
                    dalMySQL.AdicionaParametros("@PesoMinRolo", daoMovimentacaoEstoques.PesoMinRolo);
                    dalMySQL.AdicionaParametros("@DescTamFicha", daoMovimentacaoEstoques.DescTamFicha);
                    dalMySQL.AdicionaParametros("@TipoProdQuimico", daoMovimentacaoEstoques.TipoProdQuimico);
                    dalMySQL.AdicionaParametros("@ItemAtivo", daoMovimentacaoEstoques.ItemAtivo);
                    dalMySQL.AdicionaParametros("@CodigoContabil", daoMovimentacaoEstoques.CodigoContabil);
                    dalMySQL.AdicionaParametros("@CodProcesso", daoMovimentacaoEstoques.CodProcesso);
                    dalMySQL.AdicionaParametros("@Lote", daoMovimentacaoEstoques.Lote);
                    dalMySQL.AdicionaParametros("@LoteProduto", daoMovimentacaoEstoques.LoteProduto);
                    dalMySQL.AdicionaParametros("@SaldoAtual", daoMovimentacaoEstoques.SaldoAtual);
                    dalMySQL.AdicionaParametros("@Volumes", daoMovimentacaoEstoques.Volumes);
                    dalMySQL.AdicionaParametros("@QtEstqInicioMes", daoMovimentacaoEstoques.QtEstqInicioMes);
                    dalMySQL.AdicionaParametros("@QtEstqFinalMes", daoMovimentacaoEstoques.QtEstqFinalMes);
                    dalMySQL.AdicionaParametros("@UltimaEntrada", daoMovimentacaoEstoques.UltimaEntrada);
                    dalMySQL.AdicionaParametros("@UltimaSaida", daoMovimentacaoEstoques.UltimaSaida);
                    dalMySQL.AdicionaParametros("@QtSugerida", daoMovimentacaoEstoques.QtSugerida);
                    dalMySQL.AdicionaParametros("@QtEmpenhada", daoMovimentacaoEstoques.QtEmpenhada);
                    dalMySQL.AdicionaParametros("@CnpjFornecedor", daoMovimentacaoEstoques.CnpjFornecedor);
                    dalMySQL.AdicionaParametros("@NotaFiscal", daoMovimentacaoEstoques.NotaFiscal);
                    dalMySQL.AdicionaParametros("@PeriodoEstoque", daoMovimentacaoEstoques.PeriodoEstoque);
                    dalMySQL.ExecutarManipulacao(CommandType.StoredProcedure, "uspMovimentacaoEstoquesInserir");
                }

                retorno = "ok";
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: movimentação de estoques inserida. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível inserir movimentação de estoques. Detalhes: " + retorno + " | " + data);
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
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Sucesso: Relatório de movimentação de estoques deletadas. Detalhes: " + retorno + " | " + data);

            }
            catch (Exception ex)
            {
                retorno = ex.Message;
                bllFerramentas.GravarLog(@"C:\integratextil\logs\logs.txt", "Erro: não foi possível deletar relatório de movimentação de estoques renomeada. Detalhes: " + retorno + " | " + data);
            }

            return retorno;
        }

        #endregion
    }
}
