using System;
using ClosedXML.Excel;

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
            MakeFile(records,"C:/Users/dan/source/repos/ReportMaker/ReportMaker/test.xlsx");
        }
        static public void MakeFile(List<recordDummy> records,String filePath)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Sample sheet");
            worksheet.Style.Font.Bold = true;

            var columnTitles = worksheet.Range("A1:Y1");
            columnTitles.Style.Fill.BackgroundColor = XLColor.Gray;
            columnTitles.Style.Font.FontColor = XLColor.White;
            worksheet.Cell(1, 1).Value = "Vehicle";//A
            worksheet.Column("A").Width = 12;

            worksheet.Cell(1, 2).Value = "Depot Name";//B
            worksheet.Column("B").Width = 10;

            worksheet.Cell(1, 3).Value = "Start Date";//C
            worksheet.Column("C").Width = 20;

            worksheet.Cell(1, 4).Value = "End Date";//D
            worksheet.Column("D").Width = 20;

            worksheet.Cell(1, 5).Value = "Veh Group Code";//E
            worksheet.Column("E").Width = 20;

            worksheet.Cell(1, 6).Value = "Drive Hours";//F
            worksheet.Column("F").Width = 12;

            worksheet.Cell(1, 7).Value = "Idle Hours";//G
            worksheet.Column("G").Width = 12;

            worksheet.Cell(1, 8).Value = "Fuel Used";//H
            worksheet.Column("H").Width = 10;

            worksheet.Cell(1, 9).Value = "Miles Per KWh";//I
            worksheet.Column("I").Width = 10;

            worksheet.Cell(1, 10).Value = "MPG";//J
            worksheet.Column("J").Width = 10;

            worksheet.Cell(1, 11).Value = "Route Count";//K
            worksheet.Column("K").Width = 10;

            worksheet.Cell(1, 12).Value = "Data Weeks";//L
            worksheet.Column("L").Width = 10;

            worksheet.Cell(1, 13).Value = "Average Monthly Mileage";//M
            worksheet.Column("M").Width = 20;

            worksheet.Cell(1, 14).Value = "Total Route Miles";//N
            worksheet.Column("N").Width = 20;

            worksheet.Cell(1, 15).Value = "Average Daily Mileage";//O
            worksheet.Column("O").Width = 20;

            worksheet.Cell(1, 16).Value = "No Of Days";//P
            worksheet.Column("P").Width = 20;

            worksheet.Cell(1, 17).Value = "No Of Work Days";//Q
            worksheet.Column("Q").Width = 20;

            worksheet.Cell(1, 18).Value = "Suitable for Electric";//R
            worksheet.Column("R").Width = 20;

            worksheet.Cell(1, 19).Value = "% of Battery Used in a Day";//S
            worksheet.Column("S").Width = 25;

            worksheet.Cell(1, 20).Value = "Level";//T
            worksheet.Column("T").Width = 20;

            worksheet.Cell(1, 21).Value = "Average EV Cost @ £0.196";//U
            worksheet.Column("U").Width = 20;

            worksheet.Cell(1, 22).Value = "Average Diesel Cost @ £1.69";//V
            worksheet.Column("V").Width = 20;

            worksheet.Cell(1, 23).Value = "Fuel Cost Savings";//W
            worksheet.Column("W").Width = 20;

            worksheet.Cell(1, 24).Value = "CO2 (kg)";//X
            worksheet.Column("X").Width = 20;

            int row = 2;
            foreach(recordDummy record in records)
            {
                worksheet.Cell(row, 1).Value = record.VehicleRegistration;
                worksheet.Cell(row, 1).Value = record.DepotName;
                worksheet.Cell(row, 3).Value = record.StartDate;
                worksheet.Cell(row, 4).Value = record.EndDate;
                worksheet.Cell(row, 5).Value = record.VehicleGroupCode;
                worksheet.Cell(row, 6).Value = record.DriveHours;
                worksheet.Cell(row, 7).Value = record.IdleHours;
                worksheet.Cell(row, 8).Value = record.FuelUsed;
                worksheet.Cell(row, 9).Value = record.MilesPerKWh;
                worksheet.Cell(row, 10).Value = record.MilesPerGallon;
                worksheet.Cell(row, 11).Value = record.RouteCount;
                worksheet.Cell(row, 12).Value = record.DataCapturedOver;
                worksheet.Cell(row, 13).Value = record.AverageMonthlyMileage;
                worksheet.Cell(row, 14).Value = record.TotalMileage;
                worksheet.Cell(row, 15).Value = record.AverageDailyMileage;
                worksheet.Cell(row, 16).Value = record.DaysCount;
                worksheet.Cell(row, 17).Value = record.WorkingDaysCount;

             
                if(record.SuitableForEV == true)
                {
                    worksheet.Cell(row, 18).Value = "Y";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green;
                }
                else
                {
                    worksheet.Cell(row, 18).Value = "N";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red;
                }

                worksheet.Cell(row, 19).Value = record.DailyBatteryUsagePercentage;//needs extra special formatting
                row++;
            }
            workbook.SaveAs(filePath);  
        }
    }
}
