using System;

namespace MacValvesWordGenerate.ViewModels
{
    public class ViewModelLocator
    {
        public ViewModelLocator()
        {
         ViewModel = new WordViewModel();
        }

        public WordViewModel ViewModel { get; }
    }
}
