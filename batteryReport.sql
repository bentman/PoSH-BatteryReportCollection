SELECT 
    BR.ComputerName, 
    BR.DesignCapacity, 
    BR.FullChargeCapacity, 
    BR.CycleCount, 
    BR.ActiveRuntime,
    BR.ActiveRuntimeAtDesignCapacity,
    BR.ModernStandby,
    BR.ModernStandbyAtDesignCapacity
FROM 
    BatteryReport AS BR
