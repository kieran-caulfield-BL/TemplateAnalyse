using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;

namespace TemplateAnalyse
{
        public static class Globals
    {
        public static DirectoryInfo directoryInfo { get; set; }

        static Globals()
        {

        }
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            string tempPath = "";

            if (browser.ShowDialog() == DialogResult.OK)
            {
                tempPath = browser.SelectedPath; // prints path

                Globals.directoryInfo = new DirectoryInfo(tempPath);

                if (Globals.directoryInfo.Exists)
                {
                    try
                    {
                        treeView1.Nodes.Add(LoadDirectory(Globals.directoryInfo));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }


        private TreeNode LoadDirectory(DirectoryInfo di)
        {
            if (!di.Exists)
                return null;

            TreeNode output = new TreeNode(di.Name, 0, 0);

            foreach (var subDir in di.GetDirectories())
            {
                try
                {
                    output.Nodes.Add(LoadDirectory(subDir));
                }
                catch (IOException ex)
                {
                    //handle error
                }
                catch { }
            }

            foreach (var file in di.GetFiles())
            {
                if (file.Exists)
                {
                    output.Nodes.Add(file.Name);
                }
            }

            return output;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string SelectedDoc = e.Node.Text;
            int SelectedDocIndex = e.Node.Index;

            string document = Path.Combine(Globals.directoryInfo.FullName, SelectedDoc);

            SearchAndHighlight(document);

        }

        public static void SearchAndHighlight(string document)
        {
            /*using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexIfCondition = new Regex("[&If");

                int found = regexIfCondition.Matches(docText).Count;

                MessageBox.Show(Path.GetFileName(document) + " - " + found.ToString());

                //using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                //{
                //    sw.Write(docText);
                //}
            }*/

            Regex regexIfCondition = new Regex("&If");
            
            Microsoft.Office.Interop.Word.Application Word97 = new Microsoft.Office.Interop.Word.Application();
            Word97.WordBasic.DisableAutoMacros();

            Document doc = Word97.Documents.Open(document);

            //Get all words
            string allWords = doc.Content.Text;

            int found = regexIfCondition.Matches(allWords).Count;

            doc.Close();
            Word97.Quit();

            MessageBox.Show(Path.GetFileName(document) + " - " + found.ToString());
        }
    }
}
