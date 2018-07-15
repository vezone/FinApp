namespace FinApp.src
{
    class Table
    {
        public string Name       { get; set; }
        public string FirstName  => Name.Split('_')[0]; 
        public string SecondName => Name.Split('_')[1];
        public string LowName    => $"{FirstName.Split('.')[1]}.{FirstName.Split('.')[2]}";

        public int Count => m_Content.Count;

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

        public int MostPopularProduct
        {
            get
            {
                int number = 0;
                for (int i = 0; i < m_Content.Count; i++)
                {
                    if (System.Int32.Parse(m_Content[i].Number) > number)
                    {
                        number = System.Int32.Parse(m_Content[i].Number);
                    }
                }
                return number;
            }
        }
        public double MostExpensiveProduct
        {
            get
            {
                double cost = 0.0;
                for (int i = 0; i < m_Content.Count; i++)
                {
                    if (System.Double.Parse(m_Content[i].Cost) > cost)
                    {
                        cost = System.Int32.Parse(m_Content[i].Cost);
                    }
                }
                return cost;
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
        public Table(Table value)
        {
            Name = value.Name;
            m_Content = value.Content;
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
}
