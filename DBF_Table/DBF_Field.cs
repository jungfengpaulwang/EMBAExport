using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Export
{
    public class DBF_Field : IDBF_Field_Schema
    {

        #region IDBF_Field_Schema 成員

        public DBF_Field(string Name, FieldType Type, int Size, int Precision, bool AllowNull)
        {
            this.Name = Name;
            this.Type = Type;
            this.Size = Size;
            this.Precision = Precision;
            this.AllowNull = AllowNull;
        }

        public string Name { get; set; }

        public FieldType Type { get; set; }

        public int Size { get; set; }

        public int Precision { get; set; }

        public bool AllowNull { get; set; }

        #endregion
    }
}
