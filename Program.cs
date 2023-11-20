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

        // Combina os resultados em uma lista intercalada
        List<string> combinedResults = CombineResults(classificadosGeral, classificadosPPP);

        SaveToExcel(outputFilePath, classificadosGeral, classificadosPPP, combinedResults);

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

    static List<string> CombineResults(List<string> classificadosGeral, List<string> classificadosPPP)
    {
        List<string> combinedResults = new List<string>();

        int i = 0, j = 0;

        // Combina os resultados seguindo a sequência desejada
        while (i < classificadosGeral.Count && j < classificadosPPP.Count)
        {
            // Adiciona 3 elementos do arquivo 1
            for (int count = 0; count < 3 && i < classificadosGeral.Count; count++)
            {
                combinedResults.Add(classificadosGeral[i]);
                i++;
            }

            // Adiciona 1 elemento do arquivo 2
            if (j < classificadosPPP.Count)
            {
                combinedResults.Add(classificadosPPP[j]);
                j++;
            }
        }

        return combinedResults;
    }

    static void SaveToExcel(string filePath, List<string> data1, List<string> data2, List<string> combinedData)
    {

        using (var package = new ExcelPackage())
        {
            AddDataToWorksheet(package, "GERAL", data1);
            AddDataToWorksheet(package, "PPP", data2);
            AddDataToWorksheet(package, "ResultadosCombinados", combinedData);

            package.SaveAs(new FileInfo(filePath));
        }
    }

    static void AddDataToWorksheet(ExcelPackage package, string sheetName, List<string> data)
    {
        var worksheet = package.Workbook.Worksheets.Add(sheetName);

        for (int i = 0; i < data.Count; i++)
        {
            worksheet.Cells[i + 1, 1].Value = data[i];
        }
    }
}
