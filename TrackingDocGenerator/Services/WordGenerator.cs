using System;
using System.Collections.Generic;
using TrackingDocGenerator.Models;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace TrackingDocGenerator.Services
{
    public class WordGenerator
    {
        public void GenerateLabel(string outputPath, List<TrackingInfo> items)
        {
            int pageSize = 16;
            int total = items.Count;

            var doc = DocX.Create(outputPath);

            for (int i = 0; i < total; i += pageSize)
            {
                var pageItems = items.GetRange(i, Math.Min(pageSize, total - i));
                var table = doc.AddTable(8, 2);

                int current = 0;

                for (int r = 0; r < 8; r++)
                {
                    for (int c = 0; c < 2; c++)
                    {
                        var cell = table.Rows[r].Cells[c];
                        string text = "";

                        if (current < pageItems.Count)
                        {
                            var item = pageItems[current++];
                            text = $"Sender: {item.Sender}\n" +
                                   $"Receiver: {item.Receiver}\n" +
                                   $"Weight: {item.Weight}\n" +
                                   $"Order ref: {item.TrackingNumber}";
                        }

                        var p = cell.Paragraphs[0];
                        p.Append(text)
                         .Font("Segoe UI")
                         .FontSize(15);
                    }
                }

                table.Alignment = Alignment.center;
                table.AutoFit = AutoFit.Window;

                doc.InsertTable(table);

                if (i + pageSize < total)
                    doc.InsertParagraph().InsertPageBreakAfterSelf();
            }

            doc.Save();
        }
    }
}
