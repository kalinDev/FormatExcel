using CommandLine;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace FormatExcelNumbers
{
    class Program
    {
        static void Main(string[] args)
        {
                Parser.Default.ParseArguments<FlagOptions>(args)
                .WithParsed(opcoes =>
                {
                    var pastaDocumentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    var caminhoArquivoSaida = Path.Combine(pastaDocumentos, "ExcelFormatado.xlsx");
                    var pacoteDestino = new ExcelPackage();

                    using (var pacoteFonte = new ExcelPackage(new FileInfo(opcoes.FilePath)))
                    {
                        ExcelWorksheet planilhaFonte = pacoteFonte.Workbook.Worksheets[0];
                        ExcelWorksheet planilhaDestino = pacoteDestino.Workbook.Worksheets.Add("DadosFormatados");

                        PreencherCabecalho(planilhaDestino);

                        var linhasDestino = 1;

                        foreach (var colunaNumeros in opcoes.NumbersColumns.Split(',').Select(int.Parse))
                        {
                            for (var linha = opcoes.HasHeader ? 2 : 1; linha <= planilhaFonte.Dimension.Rows; linha++)
                            {
                                var telefone = ObterValorCelula(planilhaFonte, linha, colunaNumeros);
                                var telefoneFormatado = FormatarTelefone(telefone);

                                if (string.IsNullOrEmpty(telefoneFormatado))
                                    continue;

                                linhasDestino++;

                                var cpf = ObterValorCelula(planilhaFonte, linha, opcoes.CpfColumn);
                                var cpfFormatado = FormatarCpf(cpf);

                                PreencherCelula(planilhaDestino, linhasDestino, 1, long.Parse(cpfFormatado));
                                PreencherCelula(planilhaDestino, linhasDestino, 2, long.Parse(telefoneFormatado));
                                PreencherCelula(planilhaDestino, linhasDestino, 3, 0);

                                if (opcoes.EnterpriseColumn != 0)
                                {
                                    var empresa = ObterValorCelula(planilhaFonte, linha, opcoes.EnterpriseColumn);
                                    PreencherCelula(planilhaDestino, linhasDestino, 3, empresa);
                                }
                            }
                        }

                        pacoteDestino.SaveAs(new FileInfo(caminhoArquivoSaida));
                    }

                    Console.WriteLine("Os dados formatados foram salvos no arquivo: " + caminhoArquivoSaida);
                    Console.WriteLine("Pressione qualquer tecla para sair...");
                    Console.ReadKey();
                });
        }

        static string ObterValorCelula(ExcelWorksheet planilha, int linha, int coluna)
        {
            return planilha.Cells[linha, coluna].Value?.ToString() ?? "";
        }

        static void PreencherCelula(ExcelWorksheet planilha, int linha, int coluna, dynamic valor)
        {
            planilha.Cells[linha, coluna].Value = valor;
        }

        static void PreencherCabecalho(ExcelWorksheet planilha)
        {
            PreencherCelula(planilha, 1, 1, "NR_CPF_CNPJ");
            PreencherCelula(planilha, 1, 2, "TEL");
            PreencherCelula(planilha, 1, 3, "TESTE");
            PreencherCelula(planilha, 1, 4, "EMPRESA");
        }

        static string FormatarCpf(string cpfCnpj)
        {
            return string.IsNullOrEmpty(cpfCnpj) ? string.Empty :
                // Remover todos os caracteres não numéricos
                Regex.Replace(cpfCnpj, @"[^\d]", "");
        }

        static string FormatarTelefone(string telefone)
        {
            if (string.IsNullOrEmpty(telefone))
                return string.Empty;

            // Remover todos os caracteres não numéricos
            var somenteNumero = Regex.Replace(telefone, @"[^\d]", "");

            return somenteNumero.Length switch
            {
                10 when somenteNumero[2] >= '6' => somenteNumero.Insert(3, "9"),
                11 when somenteNumero[2] >= '6' => somenteNumero,
                _ => string.Empty
            };
        }
    }
}