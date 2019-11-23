using ExcelParser.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace ExcelParser.MVVM
{
    class DefaultDialogService : IDialogService
    {
        public string FilePath { get; set; }

        public bool OpenFileDialog()
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() != true) return false;
            FilePath = openFileDialog.FileName;
            return true;
        }

        public bool SaveFileDialog()
        {
            var saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() != true) return false;
            FilePath = saveFileDialog.FileName;
            return true;
        }

        public void ShowMessage(string message) => MessageBox.Show(message);
    }
}