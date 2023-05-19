using CommandLine;
using CommandLine.Text;

namespace FormatExcelNumbers;

public class FlagOptions
{
    [Option('p', "path", Required = true, HelpText = "Caminho do arquivo Excel.")]
    public string FilePath { get; set; } = "";

    [Option('h', "header", HelpText = "Indica se o arquivo Excel possui cabeçalho. Por padrão está com valor true")]
    public bool HasHeader { get; set; } = true;

    [Option('e', "Enterprise", HelpText = "Número da coluna da EMPRESA (começando em 1)..")]
    public int EnterpriseColumn { get; set; }
    
    [Option('c', "cpf", Required = true, HelpText = "Número da coluna do CPF (começando em 1).")]
    public int CpfColumn { get; set; }

    [Option('n', "numbers", Required = true,
        HelpText = "Número da(s) coluna(s) de números como flags, separados por vírgula (começando em 1).")]
    public string NumbersColumns { get; set; } = "";

    [Usage(ApplicationAlias = "ExcelTableReader")]
    public static IEnumerable<Example> Examples
    {
        get
        {
            yield return new Example("Ler tabela do Excel", new FlagOptions { FilePath = "caminho/do/arquivo.xlsx", HasHeader = true,  CpfColumn = 1, EnterpriseColumn =2, NumbersColumns = "3,4,5" });
        }
    }
}