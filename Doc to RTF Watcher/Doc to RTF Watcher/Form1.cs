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
        int status = 0; //0 = off, 1 = running, 2 = checking, 3 = processing
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
            //Creates a list of all the .doc and .docx files in the folder. If they're different, then we have new files to process;-
            switch (status)
            {
                case 0:
                    //If we're not running don't do anything
                    //This shouldn't ever get called but it might hang over a bit
                    break;
                case 1: //If we're watching
                    newFiles = Directory.GetFiles(dirBox.Text, "*.doc*");
                    if (!newFiles.SequenceEqual(oldFiles))
                    {   //So we have new files to process.
                        //Before we do the processing, we want to check that all the files can be read
                        status = 2; //Update status flag to show that we're checking the files
                                    //This will cause us to wait for a second or so, which is fine
                        oldFiles = newFiles;
                    }
                    break;
                case 2: //Checking
                    //Check all files in the folder, not just the ones that triggered change
                    string[] checkFiles = Directory.GetFiles(dirBox.Text, "*.doc.*");
                    Boolean readyToGo = true;
                    foreach (string fileName in checkFiles)
                    {
                        FileInfo fi = new FileInfo(fileName);
                        if (isFileLocked(fi))
                        {   //If a single file is locked then there's no point going any further; we stop and wait for the next tick
                            readyToGo = false;
                            break;  //No point checking them all
                        }
                    }
                    if (readyToGo)
                    {   //If all the files were unlocked, then move on to processing
                        status = 3;
                        runConversion();
                    }
                    break;
                case 3: //Processing
                        //Don't check for new files as that's done by the startWatcher method
                    break;
            }
        }

        protected virtual bool isFileLocked(FileInfo file)
        {   //Returns whether a file is locked
            //YES YES I KNOW THIS IS VULNERABLE TO RACE CONDITIONS
            //but it really doesn't matter because it's only ever going to be used on a closed, private network by *editors*
            //A more valid criticism would be that this is akin to waking someone up to ask them whether they were asleep
            //But it does make for a much simpler flow above, so I think it's worth it.
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {   //The file is unavailable for some reason
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
            return false;
        }

        void startWatcher()
        {   //starts the timer
            status = 1; //We're watching now
            //assumes that the input folder should be empty, meaning that we'll handle blank files well
            oldFiles = Directory.GetFiles(dirBox.Text, "therecannotpossiblybeafilethatmatchesthispatternfgdjskalghfdjklsghfdjkghdfkglhfdsjkghfdjksghfdjkslghfdjksghfdjksglhfdjkghfjdlsghirikgfdio");
            //Ok, quick explanation of that line - there's no good way of feeding an "empty" array into an existing array - you have to reinitalise and that causes its own problems
            //So we'll give it a list of (hopefully) non-existent files.
            watchTimer.Enabled = true;
            watchTimer.Stop();
            watchTimer.Start();
        }

        void stopWatcher()
        {   //Stops the timer and unlocks the UI
            status = 0;
            watchTimer.Enabled = false;
            watchTimer.Stop();
            startButton.Text = "Start";
            dirBox.Enabled = true;
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
            switch (status)
            {
                case 0: //Inactive
                    startWatcher(); //Start the watcher
                    startButton.Text = "Stop";
                    dirBox.Enabled = false;
                    break;
                case 1: //Watching but not processing or checking
                    //We can just stop the watcher straight away
                    stopWatcher();
                    break;
                case 2: //If it's processing or checking then tell it to shutdown when it's done
                case 3:
                    MessageBox.Show("Unable to stop the watcher because a conversion is currently processing. Please wait a few seconds and try again.");
                    break;
            }
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

            //Delete the files we've converted
            foreach (string fileName in fileNames)
            {
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception) { }
            }

            //Finally, restart the watcher
            startWatcher();
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
