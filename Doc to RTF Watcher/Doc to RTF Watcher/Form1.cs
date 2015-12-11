using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO; //For the directory
using Word = Microsoft.Office.Interop.Word; //Right-click solution, Add>Reference>Microsoft word 14.0 Object Library
using System.Timers;    //For the timer

namespace Doc_to_RTF_Watcher
{
    public partial class WatcherForm : Form
    {
        //Watcher variables
        Boolean watching = false;
        Boolean processing = false;
        string[] oldFiles;
        string[] newFiles;
        string[] emptyFiles;
        System.Timers.Timer watchTimer = new System.Timers.Timer(5000);


        public WatcherForm()
        {
            InitializeComponent();
            watchTimer.Enabled = false;
            watchTimer.Elapsed += new ElapsedEventHandler(watchTimer_Elapsed);
        }

        void watchTimer_Elapsed(object sender, ElapsedEventArgs e)
        {   //This gets called once every five seconds while the timer is running
            //Creates a list of all the .doc and .docx files in the folder. If they're different, then we have new files to process
            newFiles = Directory.GetFiles(dirBox.Text, "*.doc*");
            if (!newFiles.SequenceEqual(oldFiles))
            {
                MessageBox.Show("Change detected!");
                oldFiles = newFiles;
            }
        }

        void startWatcher()
        {   //starts the timer
            //assumes that the input folder should be empty, meaning that we'll handle blank files well
            oldFiles = Directory.GetFiles(dirBox.Text, "*.doc*");
            watchTimer.Enabled = true;
            watchTimer.Stop();
            watchTimer.Start();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {   //Allow the user to GUI their way to the folder if they don't have a keyboard
            DialogResult result = folderBrowser.ShowDialog();
            if (result==DialogResult.OK)
            {
                dirBox.Text = folderBrowser.SelectedPath;
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            startWatcher();
        }

        private void runConversion()
        {   //Converts all the .doc and .docx files in the folder into RTFs

            //Create an output folder if it doesn't exist
            if (!Directory.Exists(dirBox.Text + @"\Output"))
            {
                Directory.CreateDirectory(dirBox.Text + @"\Output");
            }

            //First get all the files in the folder that we need to convert
            string[] fileNames = Directory.GetFiles(dirBox.Text, "*.doc*");

            //Now create word and get transforming
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document wordDoc;
            object unknown = Type.Missing;

            //Start Word and open the document
            foreach (string fileName in fileNames)
            {   //Open the file
                Console.WriteLine(fileName);
                object newName = dirBox.Text + @"\Output\" + Path.GetFileNameWithoutExtension(fileName) + ".rtf";
                object format = Word.WdSaveFormat.wdFormatRTF;  //Have to set this to an instance of an object; we can't just use the reference in SaveAs2.
                try
                {   //This returns an error even though the file is fine
                    wordDoc = wordApp.Documents.Open(fileName);
                    wordDoc.SaveAs2(newName, format);
                    wordDoc.Close();
                }
                catch (Exception) { }
            }

            //Close Word and release the objects
            wordApp.Quit();
            releaseObject(wordApp);
            wordDoc = null; //We have to do this because of the try loop above
            releaseObject(wordDoc);

            //Finally, delete the files we've converted
            foreach (string fileName in fileNames)
            {
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception) { }
            }
        }

        //Release interop objects
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception x)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();   //Call in the binman
            }
        }
    }
}
