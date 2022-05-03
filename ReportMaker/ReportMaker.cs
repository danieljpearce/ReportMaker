using System;
using ClosedXML.Excel;
using System.Reflection;

namespace ReportMaker
{
    public class ReportMaker
    {
        static public void Main(String[] args)
        {
            List<recordDummy> records = new List<recordDummy>();
            for(int i = 0; i < 35; i++)
            {
                records.Add(new recordDummy());
            }
            MakeFile(records, "C:/Users/Daniel/source/repos/ReportMaker/ReportMaker/test.xlsx");
        }
        static public void MakeFile(List<recordDummy> records, String filePath)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Data");
            worksheet.Style.Font.Bold = true;

            var columnTitles = worksheet.Range("A1:Y1");
            columnTitles.Style.Fill.BackgroundColor = XLColor.Gray;
            columnTitles.Style.Font.FontColor = XLColor.White;

            //set columns
            string[] columns = { "Vehicle", "Depot Name", "Start Date", "End Date", "Veh Group Code", "Drive Hours", "Idle Hours",
                "Fuel Used","Miles Per KWh","MPG","Route Count","Data Weeks","Average Monthly Mileage","Total Route Miles",
                "Average Daily Mileage","No Of Days","No Of Work Days","Suitable for Electric","% of Battery Used in a Day",
                "Level","Average EV Cost @ £0.196","Average Diesel Cost @ £1.69","Fuel Cost Savings","CO2 (kg)"};
            for(int column = 0; column < columns.Length; column++)
            {
                worksheet.Cell(1, column+1).Value = columns[column];
            }

            decimal averageEVTotal = 0;
            decimal averageDieselTotal = 0;
            decimal fuelCostSavingsTotal = 0;
            decimal CO2Total = 0;
            int row = 2;
            //each record
            foreach (recordDummy record in records)
            {
                int cell = 1;
                //each property
                foreach (PropertyInfo prop in record.GetType().GetProperties()) { 
                    worksheet.Cell(row, cell).Value = prop.GetValue(record);
                    cell++;                 
                    }

                //color fill the bool
                if (record.SuitableForEV == true)
                {
                    worksheet.Cell(row, 18).Value = "Y";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green;
                }
                else
                {
                    worksheet.Cell(row, 18).Value = "N";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red;
                }

                //percentage bar the battery capacity
                var batPerDay = worksheet.Range(1, 19, records.Count()+1,19);
                batPerDay.AddConditionalFormat().DataBar(XLColor.Green,false).LowestValue().HighestValue();

                //orange the level
                worksheet.Cell(row, 20).Style.Fill.BackgroundColor = XLColor.Orange;
                worksheet.Cell(row, 23).Style.Fill.BackgroundColor = XLColor.Orange;

                //increment totals
                averageEVTotal += record.TotalMileageCostEV;
                averageDieselTotal += record.TotalMileageCostDerv;
                fuelCostSavingsTotal += record.FuelCostSavings;
                CO2Total += record.TotalMileageCO2Kg;

                row++;
            }
            row += 3;
            worksheet.Cell(row, 20).Value = "Total:";
            worksheet.Cell(row, 21).Value = averageEVTotal;
            worksheet.Cell(row, 22).Value = averageDieselTotal;
            worksheet.Cell(row, 23).Value = fuelCostSavingsTotal;
            worksheet.Cell(row, 23).Value = CO2Total;


            worksheet.Cell(1, 20).Style.Fill.BackgroundColor = XLColor.Orange;
            worksheet.Rows().AdjustToContents(1);
            worksheet.Columns().AdjustToContents(1);
            worksheet.Column("C").Width = 20;
            worksheet.Column("D").Width = 20;

            workbook.SaveAs(filePath);
            Environment.Exit(0);
        }
       
    }
}
