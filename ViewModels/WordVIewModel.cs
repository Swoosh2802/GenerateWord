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

        private String nameInput;
        private String templatePath;
        private String cityInput;


        public String CityInput
        {
            get { return cityInput; }
            set { cityInput = value; }
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

        private void GenerateButton()
        {
            var app = new Application();
            Application wordApp = new Application { Visible = true };
            var test = TemplatePath[(TemplatePath.IndexOf(":") + 2)..];
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(test, ReadOnly: false, Visible: true);
            aDoc.Activate();
            //WordManager.FindAndReplace(wordApp, "{{NAME}}", nameInput);

            Microsoft.Office.Interop.Word.Range range = aDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
            range.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);

            
            foreach (Section section in aDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
                headerRange.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);

                foreach (Shape shape in headerRange.ShapeRange)
                {
                    shape.TextFrame.TextRange.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
                }
            }

            foreach (Section wordSection in aDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Find.Execute(FindText: "{{NAME}}", ReplaceWith: nameInput, Replace: WdReplace.wdReplaceAll);
                footerRange.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);


                foreach (Shape shape in footerRange.ShapeRange)
                {
                    shape.TextFrame.TextRange.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput, Replace: WdReplace.wdReplaceAll);
                }
            }

            /*

            foreach (Microsoft.Office.Interop.Word.Shape shape in aDoc.Shapes)
            {
                shape.TextFrame.TextRange.Find.Execute(FindText: "{{City}}", ReplaceWith: cityInput);
                /*
                rngTextBox.Find.Execute(FindText: "{{City}}",
                                    true, true,
                                    false, false,
                                    false, true,
                                    true, false, cityInput,
                                    false, false,
                                    false, false,
                                    true);
            }
            

            foreach (Microsoft.Office.Interop.Word.Shape shape in aDoc.Shapes)
            {
                if (shape.AlternativeText.Contains("{{City}}"))
                {
                    shape.TextFrame.TextRange.Text = cityInput;
                }
            }*/
        }
    }
}
