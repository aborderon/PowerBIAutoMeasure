using System;
using Microsoft.AnalysisServices.Tabular;
using System.Diagnostics;

namespace PowerBIAutoMeasure
{
    class Program
    {
        /// <summary>
        /// Get PowerBI locat port.
        /// </summary>
        /// <returns>PowerBI local port.</returns>
        static string GetPowerBIPort(){
            var ps1File = @"Get-PowerBI-DiagPort.ps1";
            var startInfo = new ProcessStartInfo()
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy unrestricted -file \"{ps1File}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true
            };
            
            Process process = Process.Start(startInfo);

            process.WaitForExit();
            string output = process.StandardOutput.ReadToEnd();
            if (process.ExitCode != 0) { 
            }

            return(output);
        }

        /// <summary>
        /// Main script.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var tableName = "";

            // Test if input arguments were supplied.
            if (args.Length == 0)
            {
                Console.WriteLine("Please enter a string argument.");
                Console.WriteLine("Usage: <tableName>");
                return;
            } else {
                tableName = args[0];
                Console.WriteLine("Run AutoMeasure for " + tableName + " table.");
            }
 
            Server server = new Server();
            server.Connect(GetPowerBIPort());
 
            Model model = server.Databases[0].Model;
 
            foreach(Table table in model.Tables)
            {
                if(table.Name == tableName){

                    foreach(Column column in table.Columns)
                    {
                        if(
                                (
                                column.DataType == DataType.Int64 ||
                                column.DataType == DataType.Decimal || 
                                column.DataType == DataType.Double 
                                )
                                && column.IsHidden==false)
                        {
                        
                            string measureName = $"{column.Name}_{table.Name}";
                            string expression = $"MAX('{table.Name}'[{column.Name}])";
                            string displayFolder = "Auto Measures";

                            Measure measure = new Measure()
                            {
                                Name = measureName ,
                                Expression = expression,
                                DisplayFolder = displayFolder
                            };
                            
                            measure.Annotations.Add(new Annotation(){Value="This is an Auto Measure"});

                            if(!table.Measures.ContainsName(measureName))
                            {
                                table.Measures.Add(measure);
                            }
                            else
                            {
                                table.Measures[measureName].Expression = expression;
                                table.Measures[measureName].DisplayFolder = displayFolder;
                            }
                        }
                    }

                    Console.WriteLine("Auto Measures generated for " +  table.Name + " table.");
                }
                else {
                    Console.WriteLine("There is no table " +  tableName + " present in the model.");    
                }
            }
            model.SaveChanges();
        }
    }
}