using System;
using System.Collections.Generic;
using System.Text;
using files.Interfaces;
using Microsoft.Win32;

namespace MacValvesWordGenerate
{
    public class FileResourceChooser : IChooseResource
    {
        public string ResourceIdentifier
        {
            get
            {
                OpenFileDialog dlg = new OpenFileDialog();
                string filePath = string.Empty;
                if (dlg.ShowDialog() == true)
                {
                    filePath = dlg.FileName;
                }
                return filePath;
            }
        }
    }

}
