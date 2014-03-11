using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Data.OleDb;

namespace EMBA.Export
{
    public class DBF_Table
    {
        public List<IDBF_Field_Schema> Fields { get; set; }
        public List<DBF_Row> Rows { get; private set; }
        public string TableName { get; set; }
        
        public DBF_Table()
        {
            this.Fields = new List<IDBF_Field_Schema>();
            this.Rows = new List<DBF_Row>();
            this.TableName = string.Empty;
        }

        /// <summary>
        /// New Row
        /// </summary>
        public DBF_Row NewRow()
        {
            if (this.Fields.Count == 0)
                throw new Exception("請先設定資料規格。");

            DBF_Row row = new DBF_Row(this);

            return row;
        }

        /// <summary>
        /// 清除所有欄位
        /// </summary>
        public void ClearFields()
        {
            this.Fields.Clear();
        }

        /// <summary>
        /// 清除所有欄位的值
        /// </summary>
        public void ClearRows()
        {
            this.Rows.Clear();
        }
    }
}
