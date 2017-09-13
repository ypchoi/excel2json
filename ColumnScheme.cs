﻿using Excel;

namespace ExcelToJson
{
    class ColumnScheme
    {
        public enum eType
        {
            None,
            Int,
            Float,
            String,
            Bool,
        }

        public eType Type
        {
            get; set;
        }

        public string Name
        {
            get; set;
        }

        public ColumnScheme()
        {
            Type = eType.None;
            Name = "";
        }

        public bool ReadName(IExcelDataReader reader, int i)
        {
            if (reader == null)
                return false;

            if (reader.Read())

            Name = reader.GetString(i);
            return true;
        }

        public bool ReadType(IExcelDataReader reader, int i)
        {
            if (reader == null)
                return false;

            Name = reader.GetString(i);
            return true;
        }
    }

}