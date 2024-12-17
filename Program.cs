using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        var url = "https://parigi2024.coni.it/it/italia-team/gli-azzurri.html";
        var web = new HtmlWeb();
        var doc = web.Load(url);

        var athletesNodes = doc.DocumentNode.SelectNodes("//div[@class='atleta']");

        int athletsNumber = 0, counterProbs = 0;

        var excelApp = new Application();
        var workbook = excelApp.Workbooks.Add();
        var worksheet = (Worksheet)workbook.Worksheets[1];

        var headers = new HashSet<string>();
        List<Athlete> athletes = new List<Athlete>();

        if (athletesNodes != null)
        {
            Console.WriteLine($"{DateTime.Now} - Started processing athletes - it may take around 10 minutes");
            foreach (var node in athletesNodes)
            {
                var linkNode = node.SelectSingleNode(".//div[@class='nome']/a");
                var link = linkNode?.GetAttributeValue("href", null);

                if (!string.IsNullOrEmpty(link))
                {
                    var athleteUrl = "https://parigi2024.coni.it" + link;
                    var athleteDoc = web.Load(athleteUrl);

                    var name = linkNode.InnerText.Trim();
                    headers.Add("Nome");
                    var newAthlete = new Athlete();
                    newAthlete.Name = name;

                    var data = new Dictionary<string, string>();

                    var labels = athleteDoc.DocumentNode.SelectNodes("//div[@class='dati']//span[@class='label']");
                    var values = athleteDoc.DocumentNode.SelectNodes("//div[@class='dati']//span[starts-with(@class, 'value')]");
                    var additionalValues = athleteDoc.DocumentNode.SelectNodes("//p[contains(., 'Partecipazioni Olimpiche')]");

                    var disciplineNode = athleteDoc.DocumentNode.SelectSingleNode("//div[@class='disciplina']//a[@class='container-pitto']/span");
                    if (disciplineNode != null)
                    {
                        string discipline = disciplineNode.InnerText.Trim();
                        var label = "Disciplina";
                        var value = discipline;
                        data[label] = value;
                        headers.Add(label);
                    }

                    if (labels != null && values != null)
                    {
                        for (int i = 0; i < labels.Count; i++)
                        {
                            try
                            {
                                var label = labels[i]?.InnerText.Trim().Replace(":", "").Replace("&agrave;", "à");
                                var value = values[i]?.InnerText.Trim();
                                data[label] = value;
                                headers.Add(label);
                            }
                            catch
                            {
                                var label = labels[i]?.InnerText.Trim().Replace(":", "").Replace("&agrave;", "à");
                                if (!string.IsNullOrEmpty(label))
                                {
                                    data[label] = "ERROR 404";
                                    headers.Add(label);
                                }
                                if (!labels[i].InnerText.Contains("Olimpiche"))
                                {
                                    Console.WriteLine($"{DateTime.Now} - Error for athlete {newAthlete.Name} with label {labels[i].InnerText}");
                                    counterProbs++;
                                }
                            }
                        }
                    }

                    if (additionalValues != null)
                    {
                        foreach (var pNode in additionalValues)
                        {
                            var text = pNode.InnerText.Trim();
                            if (text.StartsWith("Partecipazioni Olimpiche:"))
                            {
                                var label = "Partecipazioni Olimpiche";
                                var value = text.Substring(text.IndexOf(':') + 1).Trim();
                                data[label] = value;
                                headers.Add(label);
                            }
                        }
                    }

                    athletsNumber++;
                    data["Nome"] = newAthlete.Name;
                    newAthlete.Data = data;
                    athletes.Add(newAthlete);
                }

                if (athletsNumber % 10 == 0)
                    Console.WriteLine($"{DateTime.Now} - Iterated through {athletsNumber}/{athletesNodes.Count} athletes");
            }

            int col = 1;
            var headersList = headers.ToList();
            var headerToColumn = new Dictionary<string, int>();
            foreach (var header in headersList)
            {
                worksheet.Cells[1, col] = header;
                headerToColumn[header] = col;
                col++;
            }

            int row = 2;

            foreach (var athlete in athletes)
            {
                if(row - 1 % 10 == 0)
                    Console.WriteLine($"{DateTime.Now} - Saved athlete number {row - 1}/{athletes.Count}");

                foreach (var kvp in athlete.Data)
                {
                    if (headerToColumn.ContainsKey(kvp.Key))
                    {
                        worksheet.Cells[row, headerToColumn[kvp.Key]] = kvp.Value;
                    }
                }
                row++;
            }

            Console.WriteLine($"{DateTime.Now} - Total athletes: {athletsNumber}");
            Console.WriteLine($"{DateTime.Now} - Irregularities: {counterProbs}");
        }
        else
        {
            Console.WriteLine($"{DateTime.Now} - No athletes found.");
        }

        var filePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        workbook.SaveAs(filePath);
        workbook.Close();
        excelApp.Quit();

        Console.WriteLine($"{DateTime.Now} - Excel file saved at '{filePath}' successfully.");

        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

        Thread.Sleep(3600000);
    }

    struct Athlete
    {
        public string Name;
        public Dictionary<string, string> Data;
    }
}
