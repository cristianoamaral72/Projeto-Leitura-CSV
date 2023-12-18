// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");


// Path: Projeto-Teste/Program.cs
 static List<string[]> ReadCsv(string filePath)
{
    List<string[]> valuesList = new List<string[]>();

    using (StreamReader reader = new StreamReader(filePath))
    {
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            string[] values = line.Split(',');
            valuesList.Add(values);
        }
    }

    return valuesList;
}