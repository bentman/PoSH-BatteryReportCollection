# SCCM Battery Health Report Collection and Reporting

This repository includes the `Invoke-BatteryReportCollection.ps1` script, which generates battery health reports for Windows computers, stores them in a newly created WMI class, and extends SCCM's hardware inventory to include this new data. Additionally, it provides SQL queries to generate reports in SCCM.

## How to Implement

Follow these steps below to implement this solution:

1. Download the `Invoke-BatteryReportCollection.ps1` script from this repository and customize it according to your SCCM environment's needs.

2. Create a new Package in the SCCM console:
   - Package Name: `Battery Report Collection`
   - Source Folder: Path where `Invoke-BatteryReportCollection.ps1` is stored
   - This package contains source files: Checked
   - Standard Program:
     - Name: `Run Battery Report Collection`
     - Command Line: `powershell.exe -ExecutionPolicy Bypass -File .\Invoke-BatteryReportCollection.ps1`
     - Program can run: `Whether or not a user is logged on`
   - Schedule the program to run weekly, mid-week with a client-staggered schedule.

3. Extend the SCCM Hardware Inventory to include the new WMI class:
   - In the SCCM console, navigate to `Administration > Client Settings > Default Client Settings`.
   - On the `Default Client Settings` properties page, select `Hardware Inventory`.
   - Click `Set Classes`.
   - Click `Add > Connect`.
   - In the `Connect` window, enter a computer name where the `Invoke-BatteryReportCollection.ps1` script has been run and click `Connect`.
   - In the `Select classes from this namespace` list, select the `BatteryReport` class and click `OK`.

4. Deploy the SCCM Package created in step 2 to all laptops or machines with batteries.

5. Use the provided SQL queries to generate SCCM reports as needed:
   - The first query (`batteryReport.sql`) provides a battery report for all computers with batteries.
   - The second query (`aggregateModels.sql`) aggregates computer models together, providing a count of all computer models that have a battery report.

Please ensure to replace the table and field names in the SQL queries to match the ones generated by the `Invoke-BatteryReportCollection.ps1` script in your SCCM database.

## SQL Queries

### Battery Report by Machine

The `batteryReport.sql` script queries the SCCM database for a list of all machines with batteries, along with detailed battery information.

### Aggregate Computer Models

The `aggregateModels.sql` script provides a count of all computer models that have a battery report with aggragate data averaged accross the model sample.

# Function: New-WmiClass

This PowerShell function is used to create a new WMI (Windows Management Instrumentation) class in the namespace provided. 

## Parameters

The function takes two mandatory parameters:

- `namespacePath`: A string indicating the WMI namespace path where the new class will be created.
- `newClassName`: The name of the new class to be created.

## How It Works

The function first checks if the namespace provided exists. If not, it creates a new namespace using the provided name. Then, it checks if a class with the provided name already exists in the namespace. If not, it creates a new WMI class.

The new class has the following properties:

- `ComputerName`: A string property that acts as a key.
- `DesignCapacity`: A UInt32 property.
- `FullChargeCapacity`: A UInt32 property.
- `CycleCount`: A UInt32 property.
- `ActiveRuntime`: A string property.
- `ActiveRuntimeAtDesignCapacity`: A string property.
- `ModernStandby`: A string property.
- `ModernStandbyAtDesignCapacity`: A string property.

All these properties are not nullable, meaning they must always have a value.

## Example Usage

Here's an example of how to use the `New-WmiClass` function:

```powershell
New-WmiClass -namespacePath "root\cimv2\BatteryReport" -newClassName "BatteryReport"
```
In this example, the function will create a new class named "BatteryReport" under the "root\cimv2\BatteryReport" namespace. If the namespace does not exist, it will be created.

## Function: ConvertTo-StandardTimeFormat

The `ConvertTo-StandardTimeFormat` function is an integral part of the script, converting ISO 8601 duration format into a more readable "HH:MM:SS" format.

### Parameters

- `iso8601Duration`: This is the ISO 8601 duration format string that you want to convert into the "HH:MM:SS" format.

### How It Works

The function employs regular expressions to extract hours, minutes, and seconds from the ISO 8601 duration string. Subsequently, it creates a new `TimeSpan` object with these values. Lastly, the function converts this `TimeSpan` into a string with the "HH:MM:SS" format and returns it.

This function primarily aids in the conversion of the Active Runtime and Modern Standby time values obtained from the battery report, making the data more accessible and easier to interpret when stored in the WMI class.

## Contributions

Contributions are welcome. Please open an issue or submit a pull request if you have any suggestions, questions, or would like to contribute to the project.

### GNU General Public License
This script is licensed under the GNU General Public License. You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License or any later version. 

The script is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this script. If not, see <https://www.gnu.org/licenses/>.
