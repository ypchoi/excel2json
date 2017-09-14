using System;
using System.Collections.Generic;
using Excel;

namespace ExcelToJson
{
    class Row
    {
        List<Cell> m_cells = null;

        public Row(List<CellScheme> schemes)
        {
            m_cells = new List<Cell>();

            foreach (CellScheme scheme in schemes)
            {
                m_cells.Add(Cell.Create(scheme));
            }
        }

        public bool Read(IExcelDataReader reader)
        {
            if (reader == null)
                return false;

            for (int i = 0; i < m_cells.Count; ++i)
            {
                string text = reader.GetString(i);
                if (text == null)
                    continue;

                if (!m_cells[i].Parse(text))
                {
                    Console.WriteLine("Parse failed Line {0}, Column {1} : {2}", reader.Depth, i + 1, text);
                    return false;
                }
            }

            return true;
        }

        public Cell Get(int column)
        {
            if (0 <= column && column < m_cells.Count)
                return m_cells[column];

            return null;
        }

        public override string ToString()
        {
            List<string> cells = new List<string>();
            foreach (Cell cell in m_cells)
                cells.Add(cell.ToString());

            string total = string.Join(", ", cells);
            return string.Format("{{{0}}}", total);
        }
    }
}
