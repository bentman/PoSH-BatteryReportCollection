SELECT 
    BR.ComputerName, 
    BR.DesignCapacity, 
    BR.FullChargeCapacity, 
    BR.CycleCount, 
    BR.UsageAtDesignCapacity,
    BR.UsageAtFullChargeCapacity,
    BR.StandbyAtDesignCapacity,
    BR.StandbyAtFullChargeCapacity
FROM 
    BatteryReport AS BR
