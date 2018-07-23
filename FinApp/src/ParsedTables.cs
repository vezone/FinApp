namespace FinApp.src
{
    class ParsedTables
    {
        public int csdIndex { get; set; }
        //public double[] CacheSpendingDate { get; set; }
        public string Name { get; set; }

        public int Count => m_Content.Count;

        private System.Collections.Generic.List<Table> m_Content;
        public  System.Collections.Generic.List<Table> Content
        {
            get
            {
                return m_Content;
            }
            set
            {
                Content = value;
            }
        }
        public  Table this[int id]
        {
            get
            {
                return Content[id];
            }
            set
            {
                Content[id] = value;
            }
        }

        public ParsedTables(string date)
        {
            Name = date;
            m_Content = new System.Collections.Generic.List<Table>(1);
        }
        public ParsedTables(Table[] tables)
        {
            foreach (Table table in tables)
            {
                m_Content.Add(table);
            }
        }

        public void Add(Table table)
        {
            m_Content.Add(table);
        }
        public void Remove(Table table)
        {
            m_Content.Remove(table);
        }

        public double GetCacheSpending()
        {
            double result = 0.0;
            for (int j = 0; j < Content.Count; j++)
            {
                result += Content[j].GetSummary();
            }

            return result;
        }
        public string GetStringCacheSpending()
        {
            string str = "";
            for (int j = 0; j < Content.Count; j++)
            {
                str += Content[j].SecondName + ": ";
                str += Content[j].GetSummary() + "\n";
            }
            return str;
        }

        public string LongestName()
        { 
            string longestName = "";
            for (int i = 0; i < m_Content.Count; i++)
            {
                if (m_Content[i].Name.Length > longestName.Length)
                {
                    longestName = m_Content[i].Name;
                }
            }
            return longestName;
        }

        public override string ToString()
        {
            string toString = "";

            foreach (Table product in m_Content)
            {
                toString += product.ToString() + " ";
            }
            return m_Content.ToString();
        }

    }
}
