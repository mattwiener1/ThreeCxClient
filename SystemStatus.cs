using System.Collections.Generic;

namespace ThreeCx
{
    public class SystemStatus
    {
        public string? Fqdn { get; set; }
        public string? WebMeetingFqdn { get; set; }
        public string? WebMeetingBestMcu { get; set; }
        public string? Version { get; set; }
        public int RecordingState { get; set; }
        public bool Activated { get; set; }
        public int MaxSimCalls { get; set; }
        public int MaxSimMeetingParticipants { get; set; }
        public long CallHistoryCount { get; set; }
        public int ChatMessagesCount { get; set; }
        public int ExtensionsRegistered { get; set; }
        public bool OwnPush { get; set; }
        public string? Ip { get; set; }
        public bool LocalIpValid { get; set; }
        public string? CurrentLocalIp { get; set; }
        public string? AvailableLocalIps { get; set; }
        public int ExtensionsTotal { get; set; }
        public bool HasUnregisteredSystemExtensions { get; set; }
        public bool HasNotRunningServices { get; set; }
        public int TrunksRegistered { get; set; }
        public int TrunksTotal { get; set; }
        public int CallsActive { get; set; }
        public int BlacklistedIpCount { get; set; }
        public int MemoryUsage { get; set; }
        public int PhysicalMemoryUsage { get; set; }
        public long FreeVirtualMemory { get; set; }
        public long TotalVirtualMemory { get; set; }
        public long FreePhysicalMemory { get; set; }
        public long TotalPhysicalMemory { get; set; }
        public int DiskUsage { get; set; }
        public long FreeDiskSpace { get; set; }
        public long TotalDiskSpace { get; set; }
        public long CpuUsage { get; set; }
        public List<List<object>>? CpuUsageHistory { get; set; }
        public string? MaintenanceExpiresAt { get; set; }
        public bool Support { get; set; }
        public string? ExpirationDate { get; set; }
        public int OutboundRules { get; set; }
        public bool BackupScheduled { get; set; }
        public object? LastBackupDateTime { get; set; }
        public string? ResellerName { get; set; }
        public string? LicenseKey { get; set; }
        public string? ProductCode { get; set; }
        public bool IsSpla { get; set; }
    }
}