using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using CsvHelper.Configuration.Attributes;
using System;

namespace Console_CSV
{
    public class CsvHelpers
    {
        public static List<Contrato> ReadContratosFromCsvWithCsvHelper(string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = new List<Contrato>();
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    var record = csv.GetRecord<Contrato>();
                    records.Add(record);
                }
                return records;
            }
        }
    }

    public static class CsvReaderHelper
    {
        public static List<T> ReadCsv<T>(string filePath) where T : class
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",
                TrimOptions = TrimOptions.Trim | TrimOptions.InsideQuotes,
                IgnoreBlankLines = true,
                MissingFieldFound = null, // Ignora campos faltantes
                HeaderValidated = null, // Ignora cabeçalhos inválidos ou não mapeados
                //ShouldSkipRecord = record => record.All(string.IsNullOrEmpty) // Ignora linhas totalmente vazias
            };

            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, config))
            {
                csv.Context.RegisterClassMap<SeuModeloCSVMap>();
                var records = new List<T>();
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    var record = csv.GetRecord<T>();
                    records.Add(record);
                }
                return records;
            }
        }
    }

    public sealed class SeuModeloCSVMap : ClassMap<Contrato>
    {
        public SeuModeloCSVMap()
        {
            // Mapeie as propriedades aqui, se necessário, ou deixe que os atributos [Name] façam isso
            AutoMap(CultureInfo.InvariantCulture);
        }
    }

    public class Contrato
    {

        // Propriedades já mapeadas
        [Name("Num Contrato")]
        public string NumContrato { get; set; }

        [Name("Número Operação")]
        public string NumeroOperacao { get; set; }

        [Name("Num Proposta")]
        public string NumProposta { get; set; }

        [Name("Modalidade")]
        public string Modalidade { get; set; }

        [Name("RG")]
        public string RG { get; set; }

        [Name("Dt Expedicao")]
        public string DtExpedicao { get; set; }

        [Name("Orgao Emissor")]
        public string OrgaoEmissor { get; set; }

        // Continuando com as demais propriedades
        [Name("CPF")]
        public string CPF { get; set; }

        [Name("Nome")]
        public string Nome { get; set; }

        [Name("Matricula")]
        public string Matricula { get; set; }

        [Name("Nascimento")]
        public string Nascimento { get; set; }

        [Name("Sexo")]
        public string Sexo { get; set; }

        [Name("Endereco")]
        public string Endereco { get; set; }

        [Name("NRO")]
        public string NRO { get; set; }

        [Name("Complemento")]
        public string Complemento { get; set; }

        [Name("Bairro")]
        public string Bairro { get; set; }

        [Name("CEP")]
        public string CEP { get; set; }

        [Name("Telefone")]
        public string Telefone { get; set; }

        [Name("Filiacao Mae")]
        public string FiliacaoMae { get; set; }

        [Name("Data Base")]
        public string DataBase { get; set; }

        [Name("Data Reg")]
        public string DataReg { get; set; }

        [Name("Taxa Ope")]
        public string TaxaOpe { get; set; }

        [Name("Qtde Parcelas")]
        public string QtdeParcelas { get; set; }

        [Name("Vlr Parcela")]
        public string VlrParcela { get; set; }

        [Name("Vlr Operacao")]
        public string VlrOperacao { get; set; }

        [Name("IOF")]
        public string IOF { get; set; }

        [Name("Vlr Solicitado")]
        public string VlrSolicitado { get; set; }

        [Name("Vlr Liberado")]
        public string VlrLiberado { get; set; }

        [Name("Pro Cod")]
        public string ProCod { get; set; }

        [Name("Loj Cod")]
        public string LojCod { get; set; }

        [Name("LOJ_NOME")]
        public string LOJ_NOME { get; set; }

        [Name("Loj Raz Soc")]
        public string LojRazSoc { get; set; }

        [Name("Cod Entidade")]
        public string CodEntidade { get; set; }

        [Name("Origem")]
        public string Origem { get; set; }

        [Name("Entidade")]
        public string Entidade { get; set; }

        [Name("Situacao Operacao")]
        public string SituacaoOperacao { get; set; }

        [Name("Cod Operador")]
        public string CodOperador { get; set; }

        [Name("Nome Operador")]
        public string NomeOperador { get; set; }

        [Name("Agencia")]
        public string Agencia { get; set; }

        [Name("Uf Origem")]
        public string UfOrigem { get; set; }

        [Name("Con Dat Pri Vct")]
        public string ConDatPriVct { get; set; }

        [Name("Data Reg2")]
        public string DataReg2 { get; set; }

        [Name("Situacao Função")]
        public string SituacaoFuncao { get; set; }

        [Name("Data Reg Recepçao")]
        public string DataRegRecepcao { get; set; }

        [Name("Situacao Pagamento")]
        public string SituacaoPagamento { get; set; }

        [Name("Valor Liquido")]
        public string ValorLiquido { get; set; }

        [Name("Valor Solicitado Função")]
        public string ValorSolicitadoFuncao { get; set; }

        [Name("Valor Liberado Cliente")]
        public string ValorLiberadoCliente { get; set; }

        [Name("Email")]
        public string Email { get; set; }

        [Name("Banco Origem")]
        public string BancoOrigem { get; set; }

        [Name("Op Original")]
        public string OpOriginal { get; set; }

        [Name("Tipo (Tipo De Contratação Digital/Fisico)")] // Assumindo que "Tipo" é o nome da coluna no CSV
        public string Tipo { get; set; }

        [Name("DsDigitalizacao")]
        public string DsDigitalizacao { get; set; }

        [Name("Formalizacao")]
        public string Formalizacao { get; set; }

        [Name("Data da baixa da CF01")]
        public string DataBaixaCF01 { get; set; }

        [Name("Data Kit Completo")]
        public string DataKitCompleto { get; set; }

        [Name("Nome Documento Kit Completo")]
        public string NomeDocumentoKitCompleto { get; set; }

        [Name("Nome Usuario Importação Kit")]
        public string NomeUsuarioImportacaoKit { get; set; }

        [Name("DtLiquicacao")]
        public string DtLiquicacao { get; set; }

        [Name("MotivoIrregularidade")]
        public string MotivoIrregularidade { get; set; }
    }



    public class Contratov2
    {
        public string NumContrato { get; set; }
        public string NumeroOperacao { get; set; }
        public string NumProposta { get; set; }
        public string Modalidade { get; set; }
        public string RG { get; set; }
        public string DtExpedicao { get; set; }
        public string OrgaoEmissor { get; set; }
        public string CPF { get; set; }
        public string Nome { get; set; }
        public string Matricula { get; set; }
        public string Nascimento { get; set; }
        public string Sexo { get; set; }
        public string Endereco { get; set; }
        public string NRO { get; set; }
        public string Complemento { get; set; }
        public string Bairro { get; set; }
        public string CEP { get; set; }
        public string Telefone { get; set; }
        public string FiliacaoMae { get; set; }
        public string DataBase { get; set; }
        public string DataReg { get; set; }
        public string TaxaOpe { get; set; }
        public string QtdeParcelas { get; set; }
        public string VlrParcela { get; set; }
        public string VlrOperacao { get; set; }
        public string IOF { get; set; }
        public string VlrSolicitado { get; set; }
        public string VlrLiberado { get; set; }
        public string ProCod { get; set; }
        public string LojCod { get; set; }
        public string LOJ_NOME { get; set; }
        public string LojRazSoc { get; set; }
        public string CodEntidade { get; set; }
        public string Origem { get; set; }
        public string Entidade { get; set; }
        public string SituacaoOperacao { get; set; }
        public string CodOperador { get; set; }
        public string NomeOperador { get; set; }
        public string Agencia { get; set; }
        public string UfOrigem { get; set; }
        public string ConDatPriVct { get; set; }    
    }

    public class CsvToExcelConverter
    {
        public static void ConvertCsvToExcel(string csvFilePath, string excelFilePath)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",
                HasHeaderRecord = true, // Ajuste conforme necessário
            };

            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, config))
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // Lendo o cabeçalho
                    csv.Read();
                    csv.ReadHeader();
                    var headerRow = csv.Context.Reader.HeaderRecord;

                    // Adicionando cabeçalho
                    for (int i = 0; i < headerRow.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headerRow[i];
                    }

                    // Adicionando dados
                    int currentRow = 2;
                    while (csv.Read())
                    {
                        for (int i = 0; i < headerRow.Length; i++)
                        {
                            worksheet.Cell(currentRow, i + 1).Value = csv.GetField(i);
                        }
                        currentRow++;
                    }

                    workbook.SaveAs(excelFilePath);
                }
            }
        }

        public static void ReadExcel(string excelFilePath)
        {
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                foreach (var row in range.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        Console.Write(cell.Value.ToString() + " ");
                    }
                    Console.WriteLine();
                }
            }
        }
    }


}