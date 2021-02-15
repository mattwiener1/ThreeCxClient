namespace ThreeCx
{
    public class Extension
    {
        public string? Id { get; set; }

        //public bool IsOperator { get; set; }
        public bool IsRegistered { get; set; }


        public string? Number { get; set; }

        public string? FirstName { get; set; }
        public string? LastName { get; set; }

        public string? Email { get; set; }

        public string? Password { get; set; }
        public string? MobileNumber { get; set; }

        public string? OutboundCallerId { get; set; }
        public int Phones { get; set; }

        public string? MacAddress { get; set; }
        // public string? Membership { get; set; }
        // public string? CurrentProfile { get; set; }
        // public int QueueStatus { get; set; }
        //public int Dnd { get; set; }
        //public Warning Warning { get; set; }
    }
}