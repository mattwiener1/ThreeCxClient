// ReSharper disable UnusedAutoPropertyAccessor.Local
namespace ThreeCx
{
    public class SipTrunk
    {
        public SipTrunk(string? externalNumber, string? simCalls, string? type, string? host, string? name, string? id)
        {
            ExternalNumber = externalNumber;
            SimCalls = simCalls;
            Type = type;
            Host = host;
            Name = name;
            Id = id;
        }

        private string? Id { get; set; }

        //public string? Str { get; set; }
        //public string? Number { get; set; }
        private string? Name { get; set; }
        private string? Host { get; set; }
        private string? Type { get; set; }
        private string? SimCalls { get; set; }
        private string? ExternalNumber { get; set; }

        public bool IsRegistered { get; set; }
        
        //public Gateway Gateway { get; set; }
      
        //public class MainDidNumber
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //}

        //public class SimultaneousCalls
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //}
    }
}