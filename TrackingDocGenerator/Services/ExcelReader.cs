using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TrackingDocGenerator.Models;

namespace TrackingDocGenerator.Services
{
    public class ExcelReader
    {
        public List<TrackingInfo> ReadTrackingData(string filePath)
        {
            var result = new List<TrackingInfo>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

                foreach(var row in rows)
                {
                    var item = new TrackingInfo
                    {
                        TrackingNumber = row.Cell(1).GetString().Trim(), 
                        Weight = row.Cell(2).GetString().Trim(), 
                        Sender = row.Cell(3).GetString().Trim(), 
                        Receiver = row.Cell(4).GetString().Trim()
                    };

                    result.Add(item);
                }
            }


            return result;
        }
    }
}
