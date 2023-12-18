using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Data;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.Linq;
using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;
using CsvHelper;
using System.Text.RegularExpressions;
using System.Text;


namespace Console_CSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string filePath = @"C:\DEV\Projetos Cristiano\Projeto-Teste\Console-CSV\CSV\modelolookerReduzidoCSV.csv";
            var reader = ReadCsvFile(filePath);
        }

        public static List<string[]> ReadCsvFile(string filePath)
        {
            List<string[]> records = new List<string[]>();
            string datePattern = @"(""[A-Za-z]{3} \d{1,2}, \d{4}"")"; // Padrão para identificar datas
            int dateIndex = 0; // Indexador para as datas encontradas
            Dictionary<string, string> datePlaceholders = new Dictionary<string, string>();

            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();

                    // Reseta o indexador de datas para cada linha nova
                    dateIndex = 0;
                    datePlaceholders.Clear();

                    // Substitui as datas por placeholders usando Regex.Replace
                    line = Regex.Replace(line, datePattern, match =>
                    {
                        //string placeholder = $"{{DATE{dateIndex}}}";
                        string placeholder = ExtractDateWithRegex(match.Groups[1].Value);//$"{{DATE{dateIndex++}}}";
                        datePlaceholders[placeholder] = placeholder;//match.Groups[1].Value;
                        return placeholder;
                    });

                    // Divide os campos por vírgula
                    string[] fields = line.Split(',');

                    // Restaura as datas nos respectivos lugares
                    for (int i = 0; i < fields.Length; i++)
                    {
                        if (datePlaceholders.TryGetValue(fields[i], out string dateValue))
                        {
                            fields[i] = dateValue;
                        }
                    }

                    // Adiciona o array de campos à lista de registros
                    records.Add(fields);
                }
            }

            var dados = records.Select(array => array.Where(item => !string.IsNullOrEmpty(item)).ToArray()).Where(array => array.Length > 0).ToList();

            foreach (var dado in dados)
            {
                foreach (var t in dado)
                {
                    Console.WriteLine(!string.IsNullOrEmpty(t) ? t : "Nulo");
                }
            }

            return dados;
        }

       public static string ExtractDateWithRegex(string line)
        {
            string pattern = @"(""[A-Za-z]{3} \d{1,2}, \d{4}"")"; // Padrão para encontrar datas no formato "MMM dd, yyyy"
            var match = Regex.Match(line, pattern);
            if (match.Success)
            {
                var date = ConvertDateToPortuguese(match.Groups[1].Value.Trim('"')); // Remove as aspas duplas
                return date; // Remove as aspas duplas
            }

            return null;
        }

        public static string ConvertDateToPortuguese(string dateInEnglish)
        {
            // Parse the English date
            string[] formats = { "MMM dd, yyyy", "MMM d, yyyy" };
            DateTime parsedDate = DateTime.ParseExact(dateInEnglish, formats, CultureInfo.InvariantCulture);

            // Format the date in Portuguese culture ("pt-BR")
            return parsedDate.ToString("dd 'de' MMMM 'de' yyyy", new CultureInfo("pt-BR"));
        }
    }
}
