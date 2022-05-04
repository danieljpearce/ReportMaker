using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportMaker
{
     public class Record
    {
        public string VehicleRegistration { get; set; }
        public string DepotName { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string VehicleGroupCode { get; set; }
        public TimeSpan DriveHours { get; set; }
        public TimeSpan IdleHours { get; set; }
        public decimal FuelUsed { get; set; }
        public decimal MilesPerKWh { get; set; }
        public decimal MilesPerGallon { get; set; }
        public decimal RouteCount { get; set; }
        public TimeSpan DataCapturedOver { get; set; }
        public decimal AverageMonthlyMileage { get; set; }
        public decimal TotalMileage { get; set; }
        public decimal AverageDailyMileage { get; set; }
        public decimal DaysCount { get; set; }
        public decimal WorkingDaysCount { get; set; }
        public bool SuitableForEV { get; set; }
        public decimal DailyBatteryUsagePercentage { get; set; }
        public string BatteryLevelBanding { get; set; }
        public decimal TotalMileageCostEV { get; set; }
        public decimal TotalMileageCostDerv { get; set; }
        public decimal FuelCostSavings { get; set; }
        public decimal TotalMileageCO2Kg { get; set; }
    }
}
