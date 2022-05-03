using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportMaker
{
    internal class record
    {
        string VehicleRegistration { get; set; }
        string DepotName { get; set; }
        DateTime StartDate { get; set; }
        DateTime EndDate { get; set; }
        string VehicleGroupCode { get; set; }
        TimeSpan DriveHours { get; set; }   
        TimeSpan IdleHours { get; set; }
        decimal FuelUsed { get; set; }
        decimal MilesPerKWh { get; set; }    
        decimal MilesPerGallon { get; set; }
        decimal RouteCount { get; set; }
        TimeSpan DataCapturedOver { get; set; }
        decimal AverageMonthlyMileage { get; set; }
        decimal TotalMileage { get; set; }
        decimal AverageDailyMileage { get; set; }
        decimal DaysCount { get; set; }
        decimal WorkingDaysCount { get; set; }
        bool SuitableForEV { get; set; }
        decimal DailyBatteryUsagePercentage { get; set; }
        string BatteryLevelBanding { get; set; }
        decimal TotalMileageCostEV { get; set; }
        decimal TotalMileageCostDerv { get; set; }
        decimal FuelCostSavings { get; set; }
        decimal TotalMileageCO2Kg { get; set; }
    }
}
