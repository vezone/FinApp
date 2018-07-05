using System;

namespace FinApp.src
{
    class Product
    {

        public string Name { get; set; }
        public string Number { get; set; }
        public string Price { get; set; }
        public string Cost { get; set; }

        private Product() { }

        public Product(string name, string number, string price)
        {
            Name = name;
            Number = number;
            Price = price;
            Cost = (Int32.Parse(number) * Int32.Parse(price)).ToString();
        }

        public Product(string name, string number, string price, string cost)
        {
            Name = name;
            Number = number;
            Price = price;

            string realCost
                = (Double.Parse(Number) * Double.Parse(Price))
                .ToString();

            if (cost == realCost)
            {
                Cost = cost;
            }
            else
            {
                Cost = realCost;
            }
        }

        public override string ToString()
        {
            return Name + Number + Price + Cost;
        }

        public bool IsCostCorrect()
        {
            string realCost
                = (Double.Parse(Number) * Double.Parse(Price))
                .ToString();

            if (Cost == realCost) { return true; }
            else { return false; }
        }
    }
}
