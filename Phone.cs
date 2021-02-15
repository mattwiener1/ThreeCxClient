#region

using System.Diagnostics.CodeAnalysis;
using MessagePack;

#endregion

namespace ThreeCx
{
    [MessagePackObject]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public class Phone
    {
        private string? _modelShortName;


        [Key(0)]
        public string? WhateverThisThingIs { get; set; }

        [Key(1)]
        public int Id { get; set; }

        [IgnoreMember]
        public string? UserAgent { get; set; }

        [Key(3)]
        public string? LastRegistration { get; set; }

        [IgnoreMember]
        public string? ProvMethod { get; set; }

        [IgnoreMember]

        public int DeviceType { get; set; }

        [Key(6)]
        public string? Model { get; set; }

        [IgnoreMember]
        public string? ModelShortName { get; set; }

        [IgnoreMember]
        public string? ModelDisplayName
        {
            get => _modelShortName ?? Model;
            set => _modelShortName = value;
        }

        [Key(7)]
        public string? Vendor { get; set; }

        [Key(8)]
        public string? FirmwareVersion { get; set; }

        [Key(9)]
        public string? Name { get; set; }

        [Key(10)]
        public string? UserId { get; set; }

        [Key(11)]
        public string? UserPassword { get; set; }

        [Key(12)]
        public string? Pin { get; set; }

        [Key(13)]
        public string? Ip { get; set; }

        [Key(14)]
        public string? InterfaceLink { get; set; }

        [IgnoreMember]

        public int SipPort { get; set; }

        [Key(16)]
        public string? MacAddress { get; set; }

        [IgnoreMember]

        public string? Status { get; set; }

        [Key(18)]
        public string? PhoneWebPassword { get; set; }

        [Key(19)]
        public string? ProvLink { get; set; }

        [Key(20)]
        public bool IsNew { get; set; }

        [Key(21)]
        public bool AssignExtensionEnabled { get; set; }

        [Key(22)]
        public bool CanBeRebooted { get; set; }

        [Key(23)]
        public bool CanBeUpgraded { get; set; }

        [Key(244)]
        public bool CanBeProvisioned { get; set; }

        [Key(25)]
        public bool HasInterface { get; set; }

        [Key(26)]
        public bool IsCustomProvisionTemplate { get; set; }

        [Key(27)]
        public bool UnsupportedFirmware { get; set; }

        [Key(28)]
        public string? HotdeskingExtension { get; set; }

        [IgnoreMember]

        public string? DisplayText { get; set; }

        [Key(30)]
        public string? ExtensionNumber { get; set; }
    }
}