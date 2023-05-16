using MacValvesWordGenerate.Model.Interfaces;

namespace MacValvesWordGenerate.Model
{
    public class Report : IReport
    { 
        public string _name;
        public Report(string name)
        {
            _name = name;
        }
        public string Name
        {
            get => _name;
        }
    }
}
