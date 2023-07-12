using Bindings.Commands;
using MacValvesWordGenerate.Model;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Input;

namespace MacValvesWordGenerate.ViewModels
{
    public class WordViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private String nameInput;
        private String templatePath;
        private String cityInput;
        private String applicationInput;
        private ObservableCollection<People> someCollection;
        public ObservableCollection<People> SomeCollection
        {
            get
            {
                return someCollection;
            }
            set
            {
                someCollection = value;
                NotifyPropertyChanged("SomeCollection");
            }
        }
        public ICommand TestCommand { get; private set; }


        public String CityInput
        {
            get { return cityInput; }
            set { cityInput = value; }
        }

        public String ApplicationInput
        {
            get { return applicationInput; }
            set { applicationInput = value; }
        }

        public String TemplatePath
        {
            get { return templatePath; }
            set
            {
                templatePath = "Template: "+ value; NotifyPropertyChanged(nameof(TemplatePath));
            }
        }
        public String NameInput
        {
            get { return nameInput; }
            set
            {
                nameInput = value;
            }
        }
        public ICommand PressGenerateButton { get; }
        public ICommand PressChooseTemplateButton { get; }
        public WordViewModel()
        {
            TemplatePath = "";
            PressGenerateButton = ParameterlessRelayCommand.From(GenerateButton);
            PressChooseTemplateButton = ParameterlessRelayCommand.From(ChooseTemplateButton);
            SomeCollection = new ObservableCollection<People>();
            TestCommand = ParameterizedRelayCommand<People>.From(CommandMethod);
        }

        private void CommandMethod(object parameter)
        {
            NotifyPropertyChanged("SomeCollection");
            SomeCollection.Add(new People("","","",""));
        }


        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ChooseTemplateButton()
        {
            FileResourceChooser frc = new FileResourceChooser();
            TemplatePath = frc.ResourceIdentifier;
        }


        private string generatePeopleText()
        {
            string toReturn = "";
            foreach (People poeple in someCollection){
                toReturn += "- " + poeple.Name + " " + poeple.Surname + " (" + poeple.Function + " - " + poeple.Customer + ")"+ Environment.NewLine;
            }

            return toReturn;
        }

        private void GenerateButton()
        {
            var app = new Application();
            Application wordApp = new Application { Visible = true };
            var test = TemplatePath[(TemplatePath.IndexOf(":") + 2)..];
            Document aDoc = wordApp.Documents.Open(test, ReadOnly: false, Visible: true);
            aDoc.Activate();
            Microsoft.Office.Interop.Word.Range range = aDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{CITY}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{APPLICATION}}", ReplaceWith: applicationInput, Replace: WdReplace.wdReplaceAll);

            range.Find.Execute(FindText: "{{PARTICIPANTS}}", ReplaceWith: generatePeopleText(), Replace: WdReplace.wdReplaceAll);

            foreach (Section section in aDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
                headerRange.Find.Execute(FindText: "{{CITY}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
                headerRange.Find.Execute(FindText: "{{APPLICATION}}", ReplaceWith: applicationInput, Replace: WdReplace.wdReplaceAll);
            }

            foreach (Section wordSection in aDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
                footerRange.Find.Execute(FindText: "{{CITY}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
                footerRange.Find.Execute(FindText: "{{APPLICATION}}", ReplaceWith: applicationInput, Replace: WdReplace.wdReplaceAll);
            }
        }

    }
}
