using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel;

namespace ExcelToJson
{
    class Sheet
    {
        private IExcelDataReader m_excelReader;
        private List<CellScheme> m_schemes = new List<CellScheme>();
        private List<Row> m_rows = new List<Row>();

        public Sheet()
        {
        }

        public bool Load(string path)
        {
            m_excelReader = null;
            m_schemes.Clear();
            m_rows.Clear();

            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                m_excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                if (m_excelReader == null)
                {
                    Console.WriteLine("Can not make reader");
                    return false;
                }

                if (!ReadFieldNames())
                {
                    return false;
                }

                if (!ReadFieldTypes())
                {
                    return false;
                }

                if (!ReadRows())
                {
                    return false;
                }
            }

            return true;
        }

        bool ReadFieldNames()
        {
            if (!m_excelReader.Read())
            {
                Console.WriteLine("Not exist field name row");
                return false;
            }

            for (int i = 0; i < m_excelReader.FieldCount; ++i)
            {
                CellScheme scheme = new CellScheme();
                scheme.Name = m_excelReader.GetString(i);
                m_schemes.Add(scheme);
            }

            return true;
        }

        bool ReadFieldTypes()
        {
            if (!m_excelReader.Read())
            {
                Console.WriteLine("Not exist field type row");
                return false;
            }

            if (m_excelReader.FieldCount != m_schemes.Count)
            {
                Console.WriteLine("Field name and type count are different");
                return false;
            }

            List<string> invalidTypes = new List<string>();

            for (int i = 0; i < m_schemes.Count; ++i)
            {
                CellScheme scheme = m_schemes[i];
                string text = m_excelReader.GetString(i);
                CellScheme.eType type = CellScheme.eType.None;

                if (Enum.TryParse<CellScheme.eType>(text, true, out type))
                {
                    scheme.Type = type;
                }
                else
                {
                    invalidTypes.Add(text);
                }
            }

            if (invalidTypes.Count != 0)
            {
                string types = string.Join(", ", invalidTypes);
                Console.WriteLine(string.Format("Invalid type {0}\n", types));
                Console.WriteLine("Available types are below");

                foreach (CellScheme.eType t in Enum.GetValues(typeof(CellScheme.eType)))
                {
                    Console.WriteLine("- " + t.ToString());
                }

                return false;
            }

            return true;
        }

        bool ReadRows()
        {
            while (m_excelReader.Read())
            {
                Row row = new Row(m_schemes);
                if (!row.Read(m_excelReader))
                    return false;

                m_rows.Add(row);
            }

            return true;
        }

        public bool Save(string path)
        {
            JsonWriter writer = new JsonWriter("\t");
            return writer.Write(path, m_rows);
        }
    }
}
