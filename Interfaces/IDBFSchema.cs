using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Export
{
    /// <summary>
    /// DBF 資料表欄位規格：一筆資料代表一個欄位
    /// </summary>
    public interface IDBF_Field_Schema
    {
        /// <summary>
        /// 欄位名稱
        /// </summary>
        string Name { get; set; }

        FieldType Type { get; set; }

        int Size { get; set; }

        int Precision { get; set; }

        bool AllowNull { get; set; }        
    }
}
