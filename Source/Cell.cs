namespace ExcelToJson
{
    abstract class Cell
    {
        protected CellScheme m_scheme;

        public Cell(CellScheme scheme)
        {
            m_scheme = scheme;
        }

        public abstract bool Parse(string text);

        public static Cell Create(CellScheme scheme)
        {
            switch (scheme.Type)
            {
                case CellScheme.eType.Int:
                    return new CellInt(scheme);
                case CellScheme.eType.Float:
                    return new CellFloat(scheme);
                case CellScheme.eType.Bool:
                    return new CellBool(scheme);
                case CellScheme.eType.String:
                    return new CellString(scheme);
                default:
                    return null;
            }
        }
    }

    class CellInt : Cell
    {
        int m_value = 0;

        public CellInt(CellScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            return int.TryParse(text, out m_value);
        }

        public override string ToString()
        {
            return string.Format("\"{0}\":{1}", m_scheme.Name, m_value.ToString());
        }
    }

    class CellFloat : Cell
    {
        float m_value;

        public CellFloat(CellScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            return float.TryParse(text, out m_value);
        }

        public override string ToString()
        {
            return string.Format("\"{0}\":{1}", m_scheme.Name, m_value.ToString());
        }
    }

    class CellBool : Cell
    {
        bool m_value;

        public CellBool(CellScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            if (bool.TryParse(text, out m_value))
                return true;

            int value = 0;
            if (int.TryParse(text, out value))
            {
                m_value = (value != 0);
                return true;
            }

            return false;
        }

        public override string ToString()
        {
            return string.Format("\"{0}\":{1}", m_scheme.Name, m_value ? 1 : 0);
        }
    }

    class CellString : Cell
    {
        string m_value;

        public CellString(CellScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            m_value = text;
            return true;
        }

        public override string ToString()
        {
            return string.Format("\"{0}\":\"{1}\"", m_scheme.Name, m_value.ToString());
        }
    }
}
