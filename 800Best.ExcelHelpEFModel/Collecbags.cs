//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace _800Best.ExcelHelpEFModel
{
    using System;
    using System.Collections.Generic;
    
    public partial class Collecbags
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Collecbags()
        {
            this.Cost = new HashSet<Cost>();
        }
    
        public long KeyID { get; set; }
        public string ScanSite { get; set; }
        public string ScanType { get; set; }
        public string BagID { get; set; }
        public string ID { get; set; }
        public string ScanPeople { get; set; }
        public System.DateTime ScanTime { get; set; }
        public string RecordTime { get; set; }
        public Nullable<double> Weight { get; set; }
        public string DestinationProvince { get; set; }
        public string DestinationCity { get; set; }
        public string Site { get; set; }
        public string CustomerID { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Cost> Cost { get; set; }
        public virtual Customer Customer { get; set; }
    }
}
