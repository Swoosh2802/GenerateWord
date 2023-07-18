using Bindings.Commands;
using MacValvesWordGenerate.Model;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows.Documents;
using System.Windows.Input;

namespace MacValvesWordGenerate.ViewModels
{
    public class WordViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private string customerInput;
        private string templatePath;
        private string cityInput;
        private string applicationInput;
        private string fileNeeded;
        private string distributorName;
        private string distributorSurname;
        private string distributorFunction;
        private DateTime dateInput;
        private ObservableCollection<People> peopleCollection;
        public ObservableCollection<People> PeopleCollection
        {
            get
            {
                return peopleCollection;
            }
            set
            {
                peopleCollection = value;
                NotifyPropertyChanged("PeopleCollection");
            }
        }
        public ICommand AddPeopleCommand { get; private set; }
        public string CityInput
        {
            get { return cityInput; }
            set { cityInput = value; }
        }
        public string ApplicationInput
        {
            get { return applicationInput; }
            set { applicationInput = value; }
        }
        public string TemplatePath
        {
            get { return templatePath; }
            set
            {
                templatePath = "Template: " + value; NotifyPropertyChanged(nameof(TemplatePath));
            }
        }
        public string CustomerInput
        {
            get { return customerInput; }
            set
            {
                customerInput = value;
            }
        }
        public DateTime DateInput
        {
            get => dateInput;
            set { dateInput = value; }
        }
        public string FileNeeded
        {
            get { return fileNeeded; }
            set
            {
                fileNeeded = value;
            }
        }
        public string DistributorName
        {
            get => distributorName;
            set { distributorName = value; }
        }
        public string DistributorSurname
        {
            get => distributorSurname;
            set { distributorSurname = value; }
        }
        public string DistributorFunction
        {
            get => distributorFunction;
            set { distributorFunction = value; }
        }
        public ICommand PressGenerateButton { get; }
        public ICommand PressChooseTemplateButton { get; }
        public WordViewModel()
        {
            TemplatePath = GetPathOfLastTemplateUsed();
            PressGenerateButton = ParameterlessRelayCommand.From(GenerateButton);
            PressChooseTemplateButton = ParameterlessRelayCommand.From(ChooseTemplateButton);
            PeopleCollection = new ObservableCollection<People>();
            AddPeopleCommand = ParameterizedRelayCommand<People>.From(AddPeople);
            DateInput = DateTime.Now;
        }

        private void AddPeople(object parameter)
        {
            NotifyPropertyChanged("PeopleCollection");
            PeopleCollection.Add(new People("", "", "", ""));
        }


        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ChooseTemplateButton()
        {
            FileResourceChooser frc = new FileResourceChooser();
            TemplatePath = frc.ResourceIdentifier;
            string path = JsonConvert.SerializeObject(TemplatePath);
            using var tw = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\wordDocGenerate.json", false);
            tw.WriteLine(path);
            tw.Close();
        }

        private string GetPathOfLastTemplateUsed()
        {
            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "wordDocGenerate.json");
            string toReturnPath;
            try
            {
                using (StreamReader r = new StreamReader(path))
                {
                    string json = r.ReadToEnd();
                    toReturnPath = JsonConvert.DeserializeObject<string>(json);
                    toReturnPath = toReturnPath.Substring(toReturnPath.IndexOf(" ") + 1);
                    return toReturnPath;
                }
            }
            catch (FileNotFoundException)
            {
                return "";
            }
        }


        private string GeneratePeopleText()
        {
            string toReturn = "";
            foreach (People poeple in PeopleCollection)
            {
                toReturn += "- " + poeple.Name + " " + poeple.Surname + " (" + poeple.Function + " - " + poeple.Customer + ")" + "\r";
            }
            return toReturn;
        }

        private void GenerateButton()
        {
            var test = TemplatePath[(TemplatePath.IndexOf(":") + 2)..];
            if (test.Equals(""))
            {
                FileNeeded = "Veuillez sélectionner un fichier de template";
                NotifyPropertyChanged("FileNeeded");
            }
            else
            {
                try
                {
                    var app = new Application();
                    Application wordApp = new Application { Visible = true };
                    Document aDoc = wordApp.Documents.Open(test, ReadOnly: false, Visible: true);
                    aDoc.Activate();
                    Microsoft.Office.Interop.Word.Range range = aDoc.Content;
                    GenerateWord(range);

                    foreach (Microsoft.Office.Interop.Word.Section section in aDoc.Sections)
                    {
                        Microsoft.Office.Interop.Word.Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        GenerateWord(headerRange);
                        Microsoft.Office.Interop.Word.Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        GenerateWord(footerRange);
                    }

                    FileNeeded = "";
                    NotifyPropertyChanged("FileNeeded");
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    FileNeeded = "Le fichier de template choisi est incorrect";
                    NotifyPropertyChanged("FileNeeded");
                }
            }
        }

        private void GenerateWord(Microsoft.Office.Interop.Word.Range range)
        {
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: "{{CUSTOMER}}", ReplaceWith: customerInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{CITY}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{APPLICATION}}", ReplaceWith: applicationInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{DistributorName}}", ReplaceWith: distributorName, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{DistributorSurname}}", ReplaceWith: distributorSurname, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{DistributorFunction}}", ReplaceWith: distributorFunction, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{PARTICIPANTS}}", ReplaceWith: GeneratePeopleText(), Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{DATE}}", ReplaceWith: DateInput.ToString("dd/MM/yyyy"), Replace: WdReplace.wdReplaceAll);
        }
    }
}
