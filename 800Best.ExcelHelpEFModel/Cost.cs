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
    
    public partial class Cost
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Cost()
        {
            this.Collecbags = new HashSet<Collecbags>();
        }
    
        public string CostID { get; set; }
        public string CostType { get; set; }
        public System.DateTime CostTime { get; set; }
        public double CostNum { get; set; }
        public double CostAmount { get; set; }
        public string CostAmountType { get; set; }
        public string Remarks { get; set; }
        public string CustomerID { get; set; }
        public string Tb001CostType { get; set; }
    
        public virtual Customer Customer { get; set; }
        public virtual Tb001 Tb001 { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Collecbags> Collecbags { get; set; }
    }
}
