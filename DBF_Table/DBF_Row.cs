using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Export
{
    public class DBF_Row
    {
        private Dictionary<string, string> RowData;
        private DBF_Table RowTable;

        public DBF_Row(DBF_Table table)
        {
            RowTable = table;
            RowData = new Dictionary<string, string>();

            foreach (IDBF_Field_Schema field in table.Fields)
            {
                RowData.Add(field.Name, string.Empty);
            }
        }

        public DBF_Table DBF_Table
        {
            get { return this.RowTable; }
        }

        /// <summary>
        /// 取得或設定特定欄位的值
        /// </summary>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        public string this[IDBF_Field_Schema Field]
        {
            get
            {
                if (!RowData.ContainsKey(Field.Name))
                    return null;
                else
                    return RowData[Field.Name];
            }
            set
            {
                if (!RowData.ContainsKey(Field.Name))
                    RowData.Add(Field.Name, value);
                else
                    RowData[Field.Name] = value;
            }
        }

        /// <summary>
        /// 取得或設定特定欄位的值
        /// </summary>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        public string this[string FieldName]
        {
            get
            {
                if (!RowData.ContainsKey(FieldName))
                    return null;
                else
                    return RowData[FieldName];
            }
            set
            {
                if (!RowData.ContainsKey(FieldName))
                    RowData.Add(FieldName, value);
                else
                    RowData[FieldName] = value;
            }
        }
    }
}