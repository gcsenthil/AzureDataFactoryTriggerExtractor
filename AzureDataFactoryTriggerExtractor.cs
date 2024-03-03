// Add NuGet packages:
// - Azure.ResourceManager.DataFactory
// - Azure.Identity
// - OfficeOpenXml

using Azure.Core;
using Azure.Identity;
using Azure.ResourceManager;
using Azure.ResourceManager.DataFactory;
using Azure.ResourceManager.DataFactory.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        string tenantId = Environment.GetEnvironmentVariable("TENANT_ID");
        string clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
        string clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET");
        string subscriptionId = Environment.GetEnvironmentVariable("SUBSCRIPTION_ID");
        string resourceGroupName = Environment.GetEnvironmentVariable("RESOURCE_GROUP_NAME");
        string factoryName = Environment.GetEnvironmentVariable("FACTORY_NAME");

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret) ||
            string.IsNullOrEmpty(subscriptionId) || string.IsNullOrEmpty(resourceGroupName) || string.IsNullOrEmpty(factoryName))
        {
            Console.WriteLine("Please set all required environment variables.");
            return;
        }

        try
        {
            TokenCredential cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
            ArmClient client = new ArmClient(cred);

            ResourceIdentifier dataFactoryResourceId = DataFactoryResource.CreateResourceIdentifier(subscriptionId, resourceGroupName, factoryName);
            DataFactoryResource dataFactory = client.GetDataFactoryResource(dataFactoryResourceId);

            DataFactoryTriggerCollection collection = dataFactory.GetDataFactoryTriggers();
            List<(DataFactoryScheduleTrigger, ScheduleTriggerRecurrence, string, string, string)> recurrences = new List<(DataFactoryScheduleTrigger, ScheduleTriggerRecurrence, string, string, string)>();

            await foreach (DataFactoryTriggerResource item in collection.GetAllAsync())
            {
                DataFactoryTriggerData resourceData = item.Data;

                if (resourceData.Properties is DataFactoryScheduleTrigger trigger)
                {
                    recurrences.Add((trigger, trigger.Recurrence, resourceData.Name, resourceData.Properties.Description, resourceData.Properties.RuntimeState?.ToString()));
                }
            }

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("ScheduleTriggerRecurrence");
                ExportScheduleTriggerRecurrence(recurrences, worksheet);
                FileInfo excelFile = new FileInfo("Triggers.xlsx");
                excelPackage.SaveAs(excelFile);
            }

            Console.WriteLine("Excel file generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    public static void ExportScheduleTriggerRecurrence(List<(DataFactoryScheduleTrigger, ScheduleTriggerRecurrence, string, string, string)> recurrences, ExcelWorksheet worksheet)
    {
        int row = 2;
        foreach (var recurrence in recurrences)
        {
            worksheet.Cells[row, 1].Value = recurrence.Item3;
            worksheet.Cells[row, 2].Value = recurrence.Item4;
            worksheet.Cells[row, 3].Value = recurrence.Item5;
            worksheet.Cells[row, 4].Value = recurrence.Item2.Frequency.ToString();
            worksheet.Cells[row, 5].Value = recurrence.Item2.Interval;
            worksheet.Cells[row, 6].Value = recurrence.Item2.StartOn?.ToString("yyyy-MM-dd HH:mm:ss");
            worksheet.Cells[row, 7].Value = recurrence.Item2.EndOn?.ToString("yyyy-MM-dd HH:mm:ss");
            worksheet.Cells[row, 8].Value = recurrence.Item2.TimeZone;
            if (recurrence.Item2.Schedule != null)
            {
                worksheet.Cells[row, 9].Value = string.Join(", ", recurrence.Item2.Schedule.Hours);
                worksheet.Cells[row, 10].Value = string.Join(", ", recurrence.Item2.Schedule.Minutes);
                worksheet.Cells[row, 11].Value = string.Join(", ", recurrence.Item2.Schedule.MonthDays);
                worksheet.Cells[row, 12].Value = string.Join(", ", recurrence.Item2.Schedule.WeekDays);
            }

            row++;
        }
    }
}
