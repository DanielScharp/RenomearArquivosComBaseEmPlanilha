using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Console.WriteLine("Aplicação para renomear imagens com base em planilha");


        // Caminho base na área de trabalho
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string appFolder = Path.Combine(desktopPath, "Renomear Arquivos");

        // Caminhos completos dos arquivos
        string planilhaPath = Path.Combine(appFolder, "renomear.xlsx");
        string pastaImagens = Path.Combine(appFolder, "arquivos");

        try
        {
            if(!File.Exists(planilhaPath))
            {
                Console.WriteLine($"Erro: Arquivo 'renomear.xlsx' não encontrado em {appFolder}");
                return;
            }

            if(!Directory.Exists(pastaImagens))
            {
                Console.WriteLine($"Erro: Pasta 'arquivos' não encontrada em {appFolder}");
                return;
            }

            FileInfo fileInfo = new FileInfo(planilhaPath);
            using(ExcelPackage package = new ExcelPackage(fileInfo))
            {
                if(package.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Erro: Nenhuma planilha encontrada no arquivo Excel");
                    return;
                }

                var worksheet = package.Workbook.Worksheets[0];

                if(worksheet.Dimension == null)
                {
                    Console.WriteLine("Erro: Planilha vazia ou sem dados na coluna A");
                    return;
                }
                
                int rowCount = worksheet.Dimension.Rows;
                var placas = new List<string>();

                for(int row = 1; row <= rowCount; row++)
                {
                    var cellValue = worksheet.Cells[row, 1].Value?.ToString();
                    if(!string.IsNullOrEmpty(cellValue))
                    {
                        placas.Add(cellValue.Trim());
                    }
                }

                if(placas.Count == 0)
                {
                    Console.WriteLine("Erro: Nenhuma placa encontrada na coluna A da planilha");
                    return;
                }

                var imagens = Directory.GetFiles(pastaImagens, "image*.*")
                    .OrderBy(f => ExtractNumberFromFileName(f))
                    .ToList();

                if(imagens.Count == 0)
                {
                    Console.WriteLine($"Nenhuma imagem encontrada na pasta {pastaImagens} com o padrão 'image*.*'");
                    return;
                }

                if(placas.Count != imagens.Count)
                {
                    Console.WriteLine($"Aviso: Número de placas ({placas.Count}) diferente do número de imagens ({imagens.Count})");
                    Console.WriteLine("Deseja continuar mesmo assim? (S/N)");
                    var resposta = Console.ReadLine();
                    if(resposta?.ToUpper() != "S")
                    {
                        return;
                    }
                }

                int minCount = Math.Min(placas.Count, imagens.Count);
                for(int i = 0; i < minCount; i++)
                {
                    string placa = placas[i];
                    string caminhoOriginal = imagens[i];
                    string extensao = Path.GetExtension(caminhoOriginal);
                    string nomeArquivo = RemoveInvalidFileNameChars(placa);
                    string novoNome = $"{nomeArquivo}{extensao}";
                    string novoCaminho = Path.Combine(Path.GetDirectoryName(caminhoOriginal), novoNome);

                    if(File.Exists(novoCaminho))
                    {
                        Console.WriteLine($"Pulando: {novoCaminho} já existe");
                        continue;
                    }

                    File.Move(caminhoOriginal, novoCaminho);
                    Console.WriteLine($"Renomeado: {Path.GetFileName(caminhoOriginal)} -> {novoNome}");
                }

                Console.WriteLine($"Processo concluído. {minCount} imagens renomeadas.");
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine($"Erro: {ex.Message}");
            Console.WriteLine($"Detalhes: {ex.StackTrace}");
        }

        Console.WriteLine("Pressione qualquer tecla para sair...");
        Console.ReadKey();
    }

    private static string RemoveInvalidFileNameChars(string input)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        return new string(input.Where(c => !invalidChars.Contains(c)).ToArray());
    }

    private static int ExtractNumberFromFileName(string fileName)
    {
        // Extrai o número do nome do arquivo (ex: "image123" -> 123)
        string numberStr = new string(Path.GetFileNameWithoutExtension(fileName)
            .Where(char.IsDigit)
            .ToArray());

        if(int.TryParse(numberStr, out int number))
        {
            return number;
        }
        return 0;
    }
}