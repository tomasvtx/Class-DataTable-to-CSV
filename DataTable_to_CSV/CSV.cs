using System;
using System.Data;
using System.IO;
using System.Linq;

namespace DataTableToCSV
{
    public static class CSV
    {
        /// <summary>
        /// Exportuje data z DataTable do souboru CSV.
        /// </summary>
        /// <param name="dataTable">DataTable, který obsahuje data k exportu.</param>
        /// <param name="filePath">Cesta k výstupnímu souboru CSV.</param>
        public static void ExportDataTableToCSV(this DataTable dataTable, string filePath)
        {
            try
            {
                // Pokus o vytvoření adresáře, pokud neexistuje
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                using (StreamWriter streamWriter = new StreamWriter(filePath, false))
                {
                    // Export záhlaví
                    ExportHeader(streamWriter, dataTable);

                    // Export dat
                    ExportData(streamWriter, dataTable);
                }
            }
            catch (IOException ex)
            {
                // Vyvoláno, pokud dojde k chybě vstupu nebo výstupu
                Console.WriteLine($"Chyba I/O: {ex.Message}");
            }
            catch (UnauthorizedAccessException ex)
            {
                // Vyvoláno, pokud nejsou povoleny oprávnění pro zápis do souboru
                Console.WriteLine($"Nedostatečná oprávnění: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Ostatní obecné chyby
                Console.WriteLine($"Obecná chyba: {ex.Message}");
                throw new Exception($"Chyba při exportu do CSV: {ex.Message}", ex);
            }
        }

        private static void ExportHeader(StreamWriter streamWriter, DataTable dataTable)
        {
            throw new Exception();
            // Export záhlaví do souboru CSV
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                // Zapíše název sloupce
                streamWriter.Write(dataTable.Columns[i]);
                if (i < dataTable.Columns.Count - 1)
                {
                    // Přidá čárku mezi sloupci (kromě posledního sloupce)
                    streamWriter.Write(",");
                }
            }
            // Přidá nový řádek pro oddělení záhlaví od dat
            streamWriter.WriteLine();
        }

        private static void ExportData(StreamWriter streamWriter, DataTable dataTable)
        {
            // Export dat z DataTable do souboru CSV
            foreach (DataRow dataRow in dataTable.Rows)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    // Získání hodnoty v buňce
                    string value = dataRow[i].ToString();

                    // Pokud hodnota obsahuje čárku, obalíme ji do uvozovek
                    value = value.Contains(',') ? $"\"{value}\"" : value;

                    // Zapíšeme hodnotu do souboru
                    streamWriter.Write(value);

                    if (i < dataTable.Columns.Count - 1)
                    {
                        // Přidáme čárku mezi hodnoty (kromě posledního sloupce)
                        streamWriter.Write(",");
                    }
                }
                // Přidáme nový řádek pro oddělení jednotlivých řádků
                streamWriter.WriteLine();
            }
        }
    }
}
