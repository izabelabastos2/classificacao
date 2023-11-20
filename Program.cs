using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        //Para esse código funcionar é necessário executar o seguinte comando no nu-get console: Install-Package EPPlus
        // Configura o contexto de licença do EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        string inputFilePath1 = "D:/repositories/classificacao/entrada_classificacao_GERAL.txt";
        string inputFilePath2 = "D:/repositories/classificacao/entrada_classificacao_PPP.txt";
        string outputFilePath = "D:/repositories/classificacao/saidas/classificacao_final_SERPRO.xlsx";

     
        List<string> classificadosGeral = ReadAndProcessFile(inputFilePath1);
        List<string> classificadosPPP = ReadAndProcessFile(inputFilePath2);


        SaveToExcel(outputFilePath, classificadosGeral, classificadosPPP);

        Console.WriteLine("Resultados salvos com sucesso em: " + outputFilePath);
        Console.ReadLine(); // Aguarda pressionar Enter para fechar a aplicação
    }

    static List<string> ReadAndProcessFile(string filePath)
    {
        List<string> classificados = new List<string>();

        if (File.Exists(filePath))
        {
            string[] lines = File.ReadAllLines(filePath);

            foreach (string line in lines)
            {
                string[] elementos = line.Split('/');
                classificados.AddRange(elementos);
            }
        }
        else
        {
            Console.WriteLine("O arquivo não foi encontrado: " + filePath);
        }

        return classificados;
    }

    static void SaveToExcel(string filePath, List<string> dataGeral, List<string> dataPPP)
    {
        // Cria um novo arquivo Excel
        using (var package = new ExcelPackage())
        {
            // Adiciona abas ao arquivo para cada conjunto de dados
            AddDataToWorksheet(package, "GERAL", dataGeral);
            AddDataToWorksheet(package, "PPP", dataPPP);

            // Salva o arquivo Excel no caminho especificado
            package.SaveAs(new FileInfo(filePath));
        }
    }

    static void AddDataToWorksheet(ExcelPackage package, string sheetName, List<string> data)
    {
        // Adiciona uma aba ao arquivo
        var worksheet = package.Workbook.Worksheets.Add(sheetName);

        // Preenche a planilha com os dados
        for (int i = 0; i < data.Count; i++)
        {
            // Insere cada elemento em uma linha separada
            worksheet.Cells[i + 1, 1].Value = data[i];
        }
    }
}
