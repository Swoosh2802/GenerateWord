using Bindings.Commands;
using Microsoft.Office.Interop.Word;
using System;
using System.ComponentModel;
using System.Windows.Input;

namespace MacValvesWordGenerate.ViewModels
{
    public class WordViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public string PeopleName { get; set; }

        private String nameInput;
        public String NameInput
        {
            get { return nameInput; }
            set
            {
                nameInput = value;
            }
        }
        public ICommand PressGenerateButton { get; }
        public WordViewModel()
        {
            PeopleName = "test";
            PressGenerateButton = ParameterlessRelayCommand.From(GenerateButton);
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void GenerateButton()
        {
            var app = new Application();
            Application wordApp = new Application { Visible = true };
            Document aDoc = wordApp.Documents.Open("C:\\template.docx", ReadOnly: false, Visible: true);
            aDoc.Activate();
            WordManager.FindAndReplace(wordApp, "{{NAME}}", nameInput);
        }
    }
}
