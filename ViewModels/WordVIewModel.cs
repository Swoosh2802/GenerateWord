using Bindings.Commands;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace MacValvesWordGenerate.ViewModels
{
    public class WordViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public string PeopleName { get; set; }
        public ICommand PressNameButton { get; }
        public WordViewModel()
        {
            PeopleName = "test";
            PressNameButton = ParameterlessRelayCommand.From(NameButton);
        }

        //public string ParagraphNumber => "Paragraphe #" + (_reader.CurrentParagraphIndex + 1);

        //public string ParagraphContent => _reader.CurrentText;

        //public ObservableCollection<ActionButton> Actions => _actions;

        //public ObservableCollection<ProgressListItem> Progress => _observableProgress;


        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        
        public void Refresh()
        {
            //NotifyPropertyChanged(nameof(ParagraphContent));
            //NotifyPropertyChanged(nameof(ParagraphNumber));
        }


        private void NameButton()
        {
            Console.WriteLine("appui");
            PeopleName = "poueeet";
            NotifyPropertyChanged(nameof(PeopleName));
        }
    }

}
