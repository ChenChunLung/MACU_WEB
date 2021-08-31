using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
//using MultiLangResx.Resources;

namespace MACU_WEB.MERP.MVC.Model.Models
{
    [MetadataType(typeof(FileContentMetaData))]
    public partial class FileContent
    {
        public class FileContentMetaData
        {
            public int Id { get; set; }

            /// <summary>
            /// Name 檔案名稱
            /// </summary>
            [Display(Name = "檔案名稱")]
            [Required]
            public string Name { get; set; }


            public string Url { get; set; }
            public Nullable<int> Size { get; set; }
            public string Type { get; set; }
            public System.DateTime CreateTime { get; set; }
            public string DirType { get; set; }
            public int IsValid { get; set; }
            public string ProgCatg { get; set; }
            public System.DateTime UpdateTime { get; set; }
            public string DataYear { get; set; }
            public string DataMonth { get; set; }

        }


    }
}