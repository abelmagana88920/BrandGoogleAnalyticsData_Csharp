//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace BrandGoogleAnalyticsData.Models.DB
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblBrandGoogleAanalyticsData
    {
        public int id { get; set; }
        public string brand { get; set; }
        public string group_field { get; set; }
        public Nullable<int> subgroup_id { get; set; }
        public string subgroup_field { get; set; }
        public string name { get; set; }
        public string value { get; set; }
        public Nullable<int> month { get; set; }
        public Nullable<int> year { get; set; }
    }
}