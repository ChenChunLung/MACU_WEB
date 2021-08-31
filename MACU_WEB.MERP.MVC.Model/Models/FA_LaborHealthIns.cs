using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace MACU_WEB.MERP.MVC.Model.Models
{
    [MetadataType(typeof(FA_LaborHealthInsMetaData))]
    public partial class FA_LaborHealthIns
    {

        public class FA_LaborHealthInsMetaData
        {
            public int Id { get; set; }

            /// <summary>
            /// DepartName 公司
            /// </summary>
            [Display(Name = "部門名稱")]
            [Required]
            public string DepartName { get; set; }
            // [Display(Name = "CATEGORY_NAME", ResourceType = typeof(Resource))]
            //[Required(ErrorMessageResourceName = "ValRequired", ErrorMessageResourceType = typeof(Resource))]

            /// <summary>
            /// PlusInsCompany 公司
            /// </summary>
            [Display(Name = "公司")]
            public string PlusInsCompany { get; set; }

            /// <summary>
            /// CODEING 編碼
            /// </summary>
            //[StringLength(6)]
            [Display(Name = "編碼")]
            public string Coding { get; set; }

            /// <summary>
            /// LABORINS 勞保
            /// </summary>
            [Display(Name = "勞保")]
            public Nullable<int> LaborIns { get; set; }

            /// <summary>
            /// HealthIns 健保
            /// </summary>
            [Display(Name = "健保")]
            public Nullable<int> HealthIns { get; set; }


            public string GroupIns { get; set; }
            public string JobTitle { get; set; }
            public string OnJobDate { get; set; }
            public string ResignDate { get; set; }
            public string Seniority { get; set; }
            public string KeepSecret { get; set; }
            public string Salary { get; set; }
            public string MemberName { get; set; }           
            public string DataYear { get; set; }
            public string DataMonth { get; set; }
        }


    }
}