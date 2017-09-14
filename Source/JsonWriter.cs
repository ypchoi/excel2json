using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToJson
{
    class JsonWriter
    {
        string m_rowPrefix;

        public JsonWriter(string rowPrefix)
        {
            m_rowPrefix = rowPrefix;
        }

        public bool Write(string path, List<Row> rows)
        {
            Console.WriteLine("Write {0} rows", rows.Count);

            try
            {
                using (StreamWriter file = new StreamWriter(path))
                {
                    file.WriteLine("[");

                    for (int i = 0; i < rows.Count; ++i)
                    {
                        file.Write(m_rowPrefix);
                        file.Write(rows[i].ToString());

                        string postfix = (i < rows.Count - 1) ? "," : "";
                        file.WriteLine(postfix);
                    }

                    file.WriteLine("]");

                    Console.WriteLine("Writing finished");
                    return true;
                }
            }
            catch (DirectoryNotFoundException /*e*/)
            {
                Console.WriteLine("Can not open file : {0}", path);
                return false;
            }
        }
    }
}
