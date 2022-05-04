using System;
using ClosedXML.Excel;
using System.Reflection;

namespace ReportMaker { 
    public class ReportGenerator
    {
        IXLWorkbook MakeWorkbook(List<Record> records)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Data");
            IXLWorksheet summary = workbook.Worksheets.Add("Summary");

            worksheet.Style.Font.Bold = true;
            worksheet.Style.Font.FontSize = 10;

            var columnTitles = worksheet.Range("A1:Y1");
            columnTitles.Style.Fill.BackgroundColor = XLColor.Gray;
            columnTitles.Style.Font.FontColor = XLColor.White;

            //set columns
            string[] columnNames = { "Vehicle", "Depot Name", "Start Date", "End Date", "Veh Group Code", "Drive Hours", "Idle Hours",
                    "Fuel Used","Miles Per KWh","MPG","Route Count","Data Weeks","Average Monthly Mileage","Total Route Miles",
                    "Average Daily Mileage","No Of Days","No Of Work Days","Suitable for Electric","% of Battery Used in a Day",
                    "Level","Average EV Cost @ £0.196","Average Diesel Cost @ £1.69","Fuel Cost Savings","CO2 (kg)"};
            for (int column = 0; column < columnNames.Length; column++)
            {
                worksheet.Cell(1, column + 1).Value = columnNames[column];
            }

            decimal averageEVTotal = 0;
            decimal averageDieselTotal = 0;
            decimal fuelCostSavingsTotal = 0;
            decimal CO2Total = 0;

            //summary variables 
            int SuitForEV_Y = 0;
            int SuitForEV_N = 0;

            decimal sumOfSavings_Y = 0;
            decimal sumOfSavings_N = 0;

            decimal sumOfCO2_Y = 0;
            decimal sumOfCO2_N = 0;

            int lessThan50 = 0;
            int fiftyOneToSixtyFive = 0;
            int sixtyFiveToEighty = 0;
            int eightyTo100 = 0;
            int oneHundred = 0;


            //each record
            int row = 2;
            int cell = 1;
            foreach (Record record in records)
            {
                cell = 1;
                //each property
                foreach (PropertyInfo prop in record.GetType().GetProperties())
                {
                    worksheet.Cell(row, cell).Value = prop.GetValue(record);
                    cell++;
                }

                //color fill the bool
                if (record.SuitableForEV == true)
                {
                    worksheet.Cell(row, 18).Value = "Y";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green;
                    SuitForEV_Y++;
                    sumOfSavings_Y += record.FuelCostSavings;
                    sumOfCO2_Y += record.TotalMileageCO2Kg;
                }
                else
                {
                    worksheet.Cell(row, 18).Value = "N";
                    worksheet.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red;
                    SuitForEV_N++;
                    sumOfSavings_N += record.FuelCostSavings;
                    sumOfCO2_N += record.TotalMileageCO2Kg;
                }

                //orange the level
                worksheet.Cell(row, 20).Style.Fill.BackgroundColor = XLColor.OrangePeel;
                worksheet.Cell(row, 23).Style.Fill.BackgroundColor = XLColor.OrangePeel;

                //increment level counts for summary
                if (record.BatteryLevelBanding == "<50%")
                {
                    lessThan50++;
                }
                else if (record.BatteryLevelBanding == ">=51% and <65%")
                {
                    fiftyOneToSixtyFive++;
                }
                else if (record.BatteryLevelBanding == ">=65% and <80%")
                {
                    sixtyFiveToEighty++;
                }
                else if (record.BatteryLevelBanding == ">=80% and <100%")
                {
                    eightyTo100++;
                }
                else
                {
                    oneHundred++;
                }
                //increment totals
                averageEVTotal += record.TotalMileageCostEV;
                averageDieselTotal += record.TotalMileageCostDerv;
                fuelCostSavingsTotal += record.FuelCostSavings;
                CO2Total += record.TotalMileageCO2Kg;
                row++;
            }

            //useful ranges
            var suitForElectric = worksheet.Range(1, 18, records.Count + 1, 18);
            var batteryUsageDaily = worksheet.Range(1, 19, records.Count + 1, 19);
            var level = worksheet.Range(1, 20, records.Count + 1, 20);
            var sumOfFuelCostSavings = worksheet.Range(1, 23, records.Count + 1, 23);
            var CO2 = worksheet.Range(1, 24, records.Count + 1, 24);

            //add autofilter
            worksheet.Range(1, 1, row, cell - 1).SetAutoFilter();

            //centre align the suitable for electric
            suitForElectric.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //percentage bar the battery capacity
            batteryUsageDaily.AddConditionalFormat().DataBar(XLColor.Green, XLColor.Green, false).LowestValue().HighestValue();

            //Cells with totals 
            row += 2;
            worksheet.Cell(row, 20).Value = "Total:";
            worksheet.Cell(row, 21).Value = averageEVTotal;
            worksheet.Cell(row, 22).Value = averageDieselTotal;
            worksheet.Cell(row, 23).Value = fuelCostSavingsTotal;
            worksheet.Cell(row, 23).Value = CO2Total;

            //Extra Formatting
            worksheet.Cell(1, 20).Style.Fill.BackgroundColor = XLColor.OrangePeel;
            worksheet.Cell(1, 23).Style.Fill.BackgroundColor = XLColor.OrangePeel;
            worksheet.Rows().AdjustToContents();
            worksheet.Columns().AdjustToContents();
            //stop Start/End date columns from being too small
            worksheet.Column("C").Width = 20;
            worksheet.Column("D").Width = 20;


            //summary page 
            List<IXLCell> formattableCells = new List<IXLCell>();
            //Suitable For EV Summary
            formattableCells.Add(summary.Cell(3, 1));
            formattableCells.Last().Value = "Suitable for EV";

            formattableCells.Add(summary.Cell(3, 2));
            formattableCells.Last().Value = "Count";
            summary.Cell(3, 2).Value = "Count";

            summary.Cell(4, 1).Value = "Y";
            summary.Cell(5, 1).Value = "N";

            formattableCells.Add(summary.Cell(6, 1));
            formattableCells.Last().Value = "Grand Total";

            summary.Cell(4, 2).Value = SuitForEV_Y;
            summary.Cell(5, 2).Value = SuitForEV_N;

            formattableCells.Add(summary.Cell(6, 2));
            formattableCells.Last().Value = SuitForEV_Y + SuitForEV_N;

            //% of Battery Capacity used in Shift Summary
            formattableCells.Add(summary.Cell(3, 4));
            formattableCells.Last().Value = "% Of Battery Capacity Used in Shift";

            formattableCells.Add(summary.Cell(3, 5));
            formattableCells.Last().Value = "Count";

            summary.Cell(4, 4).Value = "<50%";
            summary.Cell(5, 4).Value = ">=51% and <65%";
            summary.Cell(6, 4).Value = ">=65% and <80%";
            summary.Cell(7, 4).Value = ">=80% and <100%";
            summary.Cell(8, 4).Value = ">100%";

            summary.Cell(4, 5).Value = lessThan50;
            summary.Cell(5, 5).Value = fiftyOneToSixtyFive;
            summary.Cell(6, 5).Value = sixtyFiveToEighty;
            summary.Cell(7, 5).Value = eightyTo100;
            summary.Cell(8, 5).Value = oneHundred;

            formattableCells.Add(summary.Cell(9, 4));
            formattableCells.Last().Value = "Grand Total";
            formattableCells.Add(summary.Cell(9, 5));
            formattableCells.Last().Value = lessThan50 + fiftyOneToSixtyFive + sixtyFiveToEighty + eightyTo100 + oneHundred;

            //Suitable for EV - Sum of Fuel Cost Savings
            formattableCells.Add(summary.Cell(3, 7));
            formattableCells.Last().Value = "Suitable For EV";
            formattableCells.Add(summary.Cell(3, 8));
            formattableCells.Last().Value = "Sum of Fuel Cost Savings";

            summary.Cell(4, 7).Value = "Y";
            summary.Cell(5, 7).Value = "N";
            summary.Cell(4, 8).Value = sumOfSavings_Y;
            summary.Cell(5, 8).Value = sumOfSavings_N;

            formattableCells.Add(summary.Cell(6, 7));
            formattableCells.Last().Value = "Grand Total";
            formattableCells.Add(summary.Cell(6, 8));
            formattableCells.Last().Value = sumOfSavings_Y + sumOfSavings_N;

            //Suitable for EV - Sum of CO2
            formattableCells.Add(summary.Cell(9, 7));
            formattableCells.Last().Value = "Suitable For EV";
            formattableCells.Add(summary.Cell(9, 8));
            formattableCells.Last().Value = "Sum of Fuel Cost Savings";

            summary.Cell(10, 7).Value = "Y";
            summary.Cell(11, 7).Value = "N";
            summary.Cell(10, 8).Value = sumOfCO2_Y;
            summary.Cell(11, 8).Value = sumOfCO2_N;

            formattableCells.Add(summary.Cell(12, 7));
            formattableCells.Last().Value = "Grand Total";
            formattableCells.Add(summary.Cell(12, 8));
            formattableCells.Last().Value = sumOfCO2_Y + sumOfCO2_N;


            summary.Rows().AdjustToContents();
            summary.Columns().AdjustToContents();
            foreach (IXLCell formattableCell in formattableCells)
            {
                formattableCell.Style.Fill.BackgroundColor = XLColor.LightSkyBlue;
                formattableCell.Style.Font.Bold = true;
            }
            return workbook;
        }

    }
}