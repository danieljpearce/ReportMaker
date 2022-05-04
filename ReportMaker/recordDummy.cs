using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportMaker
{
    public class RecordDummy 
    {
        public string VehicleRegistration { get; set; } = "VehicleRegistration";
        public string DepotName { get; set; } = "DepotName";
        public DateTime StartDate { get; set; } = DateTime.Now;
        public DateTime EndDate { get; set; } = DateTime.Now;
        public string VehicleGroupCode { get; set; } = "VehicleGroupCode";
        public TimeSpan DriveHours { get; set; } = TimeSpan.FromHours(5);
        public TimeSpan IdleHours { get; set; } = TimeSpan.FromHours(10);
        public decimal FuelUsed { get; set; } = 10;
        public decimal MilesPerKWh { get; set; } = 9;
        public decimal MilesPerGallon { get; set; } = 8;
        public decimal RouteCount { get; set; } = 7;
        public TimeSpan DataCapturedOver { get; set; } = TimeSpan.FromHours(15);
        public decimal AverageMonthlyMileage { get; set; } = 5;
        public decimal TotalMileage { get; set; } = 4;
        public double AverageDailyMileage { get; set; } = Math.Round(new Random().NextDouble()*100);
        public decimal DaysCount { get; set; } = 2;
        public decimal WorkingDaysCount { get; set; } = 1;
        public bool SuitableForEV { get; set; } = false;
        public double DailyBatteryUsagePercentage { get; set; } = Math.Round(new Random().NextDouble() * 100);
        public string BatteryLevelBanding { get; set; } = "BatteryLevelBanding";
        public decimal TotalMileageCostEV { get; set; } = -1;
        public decimal TotalMileageCostDerv { get; set; } = -2;
        public decimal FuelCostSavings { get; set; } = -3;
        public decimal TotalMileageCO2Kg { get; set; } = -4;


    }
}
