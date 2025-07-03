public class InstalledApp
{
    public string Name { get; set; }
    public string Publisher { get; set; }
    public string Version { get; set; }
    public string InstallLocation { get; set; }
    public string DisplayIcon { get; set; }
    public DateTime? RegistryInstallDate { get; set; }
    public DateTime? ResolvedInstallDate { get; set; }
}
