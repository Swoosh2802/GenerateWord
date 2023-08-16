using MacValvesWordGenerate.Model.Interfaces;

namespace MacValvesWordGenerate.Model
{
    public class People : IPeople
    {
        public string _name;
        public string _function;
        public string _customer;
        public People(string name, string function, string customer)
        {
            _name = name;
            _function = function;
            _customer = customer;
        }
        public string Name
        {
            get => _name; set => _name = value;
        }
        public string Function
        {
            get => _function;
            set => _function = value;
        }
        public string Customer
        {
            get => _customer;
            set => _customer = value;
        }

    }
}
