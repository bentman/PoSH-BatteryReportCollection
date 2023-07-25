-- Create function TimeToSec to convert 'HH:MM:SS' string to seconds
CREATE FUNCTION TimeToSec (@Time varchar(12))
RETURNS int
AS 
BEGIN
   DECLARE @H int, @M int, @S int
   SELECT @H = SUBSTRING(@Time, 1, 2),
          @M = SUBSTRING(@Time, 4, 2),
          @S = SUBSTRING(@Time, 7, 2)
   RETURN @H * 3600 + @M * 60 + @S
END

-- Main Query
SELECT 
    SM.Model,
    COUNT(*) AS NumberOfMachines,
    AVG(BR.DesignCapacity) AS AverageDesignCapacity, 
    AVG(BR.FullChargeCapacity) AS AverageFullChargeCapacity, 
    AVG(BR.CycleCount) AS AverageCycleCount,
    AVG(dbo.TimeToSec(BR.ActiveRuntime)) AS AverageActiveRuntime,
    AVG(dbo.TimeToSec(BR.ActiveRuntimeAtDesignCapacity)) AS AverageActiveRuntimeAtDesignCapacity,
    AVG(dbo.TimeToSec(BR.ModernStandby)) AS AverageModernStandby,
    AVG(dbo.TimeToSec(BR.ModernStandbyAtDesignCapacity)) AS AverageModernStandbyAtDesignCapacity
FROM 
    BatteryReport AS BR
INNER JOIN 
    System_Enclosure AS SM ON BR.ComputerName = SM.SystemName
GROUP BY 
    SM.Model
