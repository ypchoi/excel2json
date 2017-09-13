namespace ExcelToJson
{
    abstract class Column
    {
        protected FieldScheme m_scheme;

        public Column(FieldScheme scheme)
        {
            m_scheme = scheme;
        }

        public abstract bool Parse(string text);

        static Column Create(FieldScheme.eType type, FieldScheme scheme)
        {
            switch (type)
            {
                case FieldScheme.eType.Int:
                    return new ColumnInt(scheme);
                case FieldScheme.eType.Float:
                    return new ColumnFloat(scheme);
                case FieldScheme.eType.Bool:
                    return new ColumnBool(scheme);
                case FieldScheme.eType.String:
                    return new ColumnString(scheme);
                default:
                    return null;
            }
        }
    }

    class ColumnInt : Column
    {
        int m_value = 0;

        public ColumnInt(FieldScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            return int.TryParse(text, out m_value);
        }

        public override string ToString()
        {
            return string.Format("\"{0}\": {1}", m_scheme.Name, m_value.ToString());
        }
    }

    class ColumnFloat : Column
    {
        float m_value;

        public ColumnFloat(FieldScheme scheme)
            : base(scheme)
        {
        }

        public override bool Parse(string text)
        {
            return float.TryParse(text, out m_value);
        }

        public override string ToString()
        {
            return string.Format("\"{0}\": {1}", m_scheme.Name, m_value.ToString());
        }
    }

    class ColumnBool : Column
    {
        bool m_value;

        public ColumnBool(FieldScheme scheme)
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
            return string.Format("\"{0}\": {1}", m_scheme.Name, m_value.ToString());
        }
    }

    class ColumnString : Column
    {
        string m_value;

        public ColumnString(FieldScheme scheme)
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
            return string.Format("\"{0}\": \"{1}\"", m_scheme.Name, m_value.ToString());
        }
    }
}
