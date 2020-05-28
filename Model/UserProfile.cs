// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

namespace VentiDemoBOT.Model
{
    using System;
    using System.Collections.Generic;
    using System.Security;

    /// <summary>Contains information about a user.</summary>
    public class UserProfile
    {

        public string GivenName { get; set; }

        public string SurName { get; set; }

        public string DisplayName { get; set; }

        public string Id { get; set; }

        public string DomainName { get; set; }

        public string Department { get; set; }

        public string JobTitle { get; set; }

        public string PhoneNumber { get; set; }

        public string Domain { get; set; }

        public string UserPrincipalName { get; set; }

        public string Password { get; set; }

        public bool AccountEnabled { get; set; }

    }
}
