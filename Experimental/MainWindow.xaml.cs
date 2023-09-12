using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Experimental
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord, 
                ref matchWildCards, ref matchSoundLike, 
                ref nmatchAllForms, ref forward, 
                ref wrap, ref format, ref replaceWithText, 
                ref replace, ref matchKashida, 
                ref matchDiactitics, ref matchAlefHamza,    
                ref matchControl);
        }

        private void CreateWordDocument(object fileName, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if(File.Exists((string)fileName))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //Find and replace
                this.FindAndReplace(wordApp, "<name>", "Hovo");
                this.FindAndReplace(wordApp, "<title>", "Greetings");
            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            // Save as
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File created!");
        }
        private void CreateWord_Click(object sender, RoutedEventArgs e)
        {
            CreateWordDocument(@"C:\Users\geghov\Desktop\temp.docx", @"C:\Users\geghov\Desktop\output.docx");
        }
    }
}
