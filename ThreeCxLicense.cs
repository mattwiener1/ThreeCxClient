namespace ThreeCx
{
    public class ThreeCxLicense
    {
        public ThreeCxLicense(string? key, int maxSimCalls)
        {
            Key = key;
            MaxSimCalls = maxSimCalls;
        }

        public string? Key { get; }
        public string? CompanyName { get; set; }
        public string? ContactName { get; set; }
        public string? Email { get; set; }
        public string? AdminEMail { get; set; }
        public string? Telephone { get; set; }
        public string? ResellerName { get; set; }
        public string? ProductCode { get; set; }
        public int MaxSimCalls { get; }
        public bool ProFeatures { get; set; }
        public string? ExpirationDate { get; set; }
    }
}