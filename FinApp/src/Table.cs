namespace FinApp.src
{
    class Table
    {
        public string Name       { get; set; }
        public string FirstName  => Name.Split('_')[0]; 
        public string SecondName => Name.Split('_')[1];
        public string LowName    => $"{FirstName.Split('.')[1]}.{FirstName.Split('.')[2]}";
        
        private System.Collections.Generic.List<Product> m_Content;
        public  System.Collections.Generic.List<Product> Content
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

        public Product this[int id]
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

        public Table()
        {
            m_Content = new System.Collections.Generic.List<Product>(0);
        }
        public Table(Product[] products)
        {
            foreach (Product product in products)
            {
                m_Content.Add(product);
            }
        }

        public void Add(Product product)
        {
            m_Content.Add(product);
        }
        public void Remove(Product product)
        {
            m_Content.Remove(product);
        }
        
        public override string ToString()
        {
            string toString = "";

            foreach (Product product in m_Content)
            {
                toString += product.ToString() + "\n";
            }
            return toString;
        }

        public double GetSummary()
        {
            double summary = 0.0;
            foreach (Product product in Content)
            {
                summary += System.Double.Parse(product.Cost);
            }
            return summary;
        }
        public double GetNumber()
        {
            double summary = 0.0;
            foreach (Product product in Content)
            {
                summary += System.Double.Parse(product.Number);
            }
            return summary;
        }
    }

    class ParsedTables
    {
        public int csdIndex { get; set; }
        public double[] CacheSpendingDate { get; set; }
        public string Date { get; set; }

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
            Date = date;
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
        
        public override string ToString()
        {
            string toString = "";

            foreach (Table product in m_Content)
            {
                toString += product.ToString() + " ";
            }
            return m_Content.ToString();
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

            //str += Date + "\n";
            for (int j = 0; j < Content.Count; j++)
            {
                str += Content[j].SecondName + ": ";
                str += Content[j].GetSummary() + "\n";
            }

            return str;
        }

    }
}
