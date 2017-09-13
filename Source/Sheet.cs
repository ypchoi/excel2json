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
        private List<ColumnScheme> m_schemes = new List<ColumnScheme>();

        public Sheet()
        {
        }

        public bool Load(string path)
        {
            m_excelReader = null;
            m_schemes.Clear();

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

            m_schemes.AddRange(Enumerable.Repeat(new ColumnScheme(), m_excelReader.FieldCount));

            for (int i = 0; i < m_schemes.Count; ++i)
            {
                ColumnScheme scheme = m_schemes[i];
                scheme.Name = m_excelReader.GetString(i);
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
                ColumnScheme scheme = m_schemes[i];
                string text = m_excelReader.GetString(i);
                ColumnScheme.eType type = ColumnScheme.eType.None;

                if (Enum.TryParse<ColumnScheme.eType>(text, true, out type))
                {
                    scheme.Type = type;
                }
                else
                {
                    invalidTypes.Add(text);
                }
            }

            bool success = (invalidTypes.Count == 0);

            if (!success)
            {
                string types = string.Join(", ", invalidTypes);
                Console.WriteLine(string.Format("Invalid type {0}\n", types));
                Console.WriteLine("Available types are below");

                foreach (ColumnScheme.eType t in Enum.GetValues(typeof(ColumnScheme.eType)))
                {
                    Console.WriteLine("- " + t.ToString());
                }
            }

            return success;
        }

        bool ReadRows()
        {
            while (m_excelReader.Read())
            {
            }

            return true;
        }
    }
}
