using System;
using System.IO;
using System.Collections.Generic;
using Excel;

namespace ExcelToJson
{
    class Document
    {
        List<Sheet> m_sheets = new List<Sheet>();

        public Document()
        {
        }

        public bool Load(string path)
        {
            m_sheets.Clear();

            try
            {
                using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    if (reader == null)
                    {
                        Console.WriteLine("Can not make reader");
                        return false;
                    }

                    do
                    {
                        Sheet sheet = new Sheet(reader.Name);
                        if (!sheet.Load(reader))
                            continue;

                        m_sheets.Add(sheet);
                    }
                    while (reader.NextResult());

                    return true;
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Can not read file : {0}", path);
                return false;
            }
        }

        public bool Save(string pathIn)
        {
            if (m_sheets.Count == 1)
            {
                Sheet sheet = m_sheets[0];
                return m_sheets[0].Save(pathIn);
            }
            else
            {
                foreach (Sheet sheet in m_sheets)
                {
                    string postfix = sheet.Name.Replace(' ', '_');

                    int index = pathIn.LastIndexOf('.');
                    if (index < 0)
                        index = pathIn.Length - 1;

                    string path = pathIn.Insert(index, string.Format("_{0}", postfix));

                    if (!sheet.Save(path))
                        return false;
                }

                return true;
            }
        }
    }
}
