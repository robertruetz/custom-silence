using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using SoundForge;

/*
 * TODO: Need to handle input errors.
 * 
 * Implement status strip for user alerts
 * 
 * Implement clear datagridview after complete
 *      -reset form.
 * 
 * Refactor output folder creation instead of selection
 *      -means refactoring sortSilence method.
 */


public class Form1 : Form
{
    IScriptableApp formApp;
    private Label dropLabel;

    string chosenFolder;
    private DataGridView dataGridView1;
    private DataGridViewTextBoxColumn startChar;
    private DataGridViewTextBoxColumn front;
    private DataGridViewTextBoxColumn back;
    private Button runButton;
    List<string> audioPathList = new List<string>();
    
    public void addSfApp(IScriptableApp appIn)
    {
        // add sound Forge app to the form
        this.formApp = appIn;
    }

    //This method has to be public for SF to see the controls.
    public void InitializeComponent()
    {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dropLabel = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.startChar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.front = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.back = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.runButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dropLabel
            // 
            this.dropLabel.AutoSize = true;
            this.dropLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dropLabel.Location = new System.Drawing.Point(11, 23);
            this.dropLabel.Name = "dropLabel";
            this.dropLabel.Size = new System.Drawing.Size(315, 20);
            this.dropLabel.TabIndex = 8;
            this.dropLabel.Text = "Drop a folder of audio files onto this window";
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.startChar,
            this.front,
            this.back});
            this.dataGridView1.Enabled = false;
            this.dataGridView1.Location = new System.Drawing.Point(56, 71);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 20;
            this.dataGridView1.Size = new System.Drawing.Size(225, 190);
            this.dataGridView1.TabIndex = 9;
            // 
            // startChar
            // 
            this.startChar.HeaderText = "Starting Char";
            this.startChar.Name = "startChar";
            this.startChar.Width = 50;
            // 
            // front
            // 
            this.front.HeaderText = "Leading Silence (ms)";
            this.front.Name = "front";
            this.front.Width = 75;
            // 
            // back
            // 
            this.back.HeaderText = "Trailing Silence (ms)";
            this.back.Name = "back";
            this.back.Width = 75;
            // 
            // runButton
            // 
            this.runButton.Location = new System.Drawing.Point(129, 276);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(75, 23);
            this.runButton.TabIndex = 10;
            this.runButton.Text = "Run";
            this.runButton.UseVisualStyleBackColor = true;
            this.runButton.Click += new System.EventHandler(this.runButton_Click);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.ClientSize = new System.Drawing.Size(336, 311);
            this.Controls.Add(this.runButton);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dropLabel);
            this.Name = "Form1";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    void Form1_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
        {
            e.Effect = DragDropEffects.All;
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    void Form1_DragDrop(object sender, DragEventArgs e)
    {
        this.chosenFolder = String.Empty;

        foreach (string s in (string[])e.Data.GetData(DataFormats.FileDrop, false))
        {
            if (Directory.Exists(s))
            {
                string[] fileList = Directory.GetFiles(s);

                foreach (string str in fileList)
                {
                    formApp.OutputText(str);
                    if (str.ToLower().EndsWith("wav") || str.ToLower().EndsWith("mp3"))
                    {
                        this.audioPathList.Add(str);
                    }
                }
            }
            else
            {
                return;
            }
        }

        if (this.audioPathList.Count > 0)
        {
            this.chosenFolder = Directory.GetParent(this.audioPathList[0]).ToString();
            this.dropLabel.Text = (this.chosenFolder + " Loaded");
            this.dataGridView1.Enabled = true;
        }
        else
        {
            MessageBox.Show("Folder does not contain WAV or MP3 files. Try again.");
        }
    }

    private void runButton_Click(object sender, EventArgs e)
    {
        Dictionary<string, double[]> ruleDict = new Dictionary<string, double[]>();

        // Get output directory
        string outPutFolder = SfHelpers.ChooseDirectory("Where do you want to save the files after processing?", this.chosenFolder);

        //pull our silence rules out of the datagrid and add them to our dictionary
        if (dataGridView1.Rows.Count > 1)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                List<string> gridContents = new List<string>();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        formApp.OutputText(cell.Value.ToString());
                        gridContents.Add(cell.Value.ToString().Trim());
                    }
                }
                
                // Problem is here. Nothing is getting added to the Dictionary.
                
                if (gridContents.Count > 0)
                {
                    string[] sil = { gridContents[1], gridContents[2] };
                    double[] silences = parseAndPrepareSilences(sil);
                    ruleDict.Add(gridContents[0], silences);
                    
                }
            }

            formApp.OutputText("Dictionary Loaded");
            
            sortSilence(ruleDict, audioPathList, outPutFolder);

            dropLabel.Text = "Drop a folder of audio files onto this window.";
            dataGridView1.Rows.Clear();
        }
    }

    private double[] parseAndPrepareSilences(string[] strings)
    {
        double[] output = new double[strings.Length];
        for (int count = 0; count < strings.Length; count++)
        {
            formApp.OutputText(strings[count]);
            double goodNum = 0;
            if (Double.TryParse(strings[count], out goodNum))
            {
                // divide by 1000 to convert to milliseconds
                formApp.OutputText("Parse Successful");
                output[count] = goodNum / 1000;
            }
        }
        return output;
    }

    private double[] addTwoDoubleArrays(double[] recArray, double[] giveArray)
    {
        recArray[0] += giveArray[0];
        recArray[1] += giveArray[1];

        formApp.OutputText(string.Format("Silences are {0} and {1}.", recArray[0], recArray[1]));

        return recArray;
    }

    private void sortSilence(Dictionary<string, double[]> rules, List<string> audioPaths, string outFolder)
    {
        foreach (string file in audioPaths)
        {
            bool willBeProcessed = false;
            double[] silences = { 0, 0 };
            // handle universal rules first
            if (rules.ContainsKey("*"))
            {
                // add the universal rule amounts to our silence array amounts
                silences = addTwoDoubleArrays(silences, rules["*"]);
                willBeProcessed = true;
            }

            //check the first char of the fileName and apply any rules
            string fName = Path.GetFileName(file);
            string startChar = fName[0].ToString();

            if (rules.ContainsKey(startChar))
            {
                formApp.OutputText("Match Found");
                silences = addTwoDoubleArrays(silences, rules[startChar]);
                willBeProcessed = true;
            }

            if (willBeProcessed)
            {
                addSilence(file, silences, this.formApp, outFolder);
            }
            else
            {
                // Any files not processed need to be copied into the output folder.
                string outPath = Path.Combine(outFolder, fName);
                File.Copy(file, outPath);
            }
        }
    }

    //Might want to consider deleting this method. It is not currently used
    private string createOutputDirectory(string folderName)
    {
        // currently not used. Look for deletion in refactoring.
        formApp.OutputText(folderName);
        string outputFolder = Path.Combine(folderName, "_SilenceAdded");
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        return outputFolder;

    }

    private List<double> millisecondsToDecimal(List<double> inputs)
    {
        List<double> outputs = new List<double>();
        foreach (double d in inputs)
        {
            // if d is 0, we don't want to divide it into milliseconds
            if (d <= 0)
            {
                outputs.Add(0);
            }
            else
            {
                outputs.Add(d / 1000);
            }
        }

        return outputs;

    }

    private void addSilence(string audioFileName, double[] silences, IScriptableApp appl, string outPutDir)
    {
        formApp.OutputText(string.Format("Adding {0}ms to front and {1} to the back of {2}", silences[0], silences[1], Path.GetFileName(audioFileName)));
        
        ISfFileHost oldFile = appl.OpenFile(audioFileName, true, true);
        ISfFileHost newFile = appl.NewFile(oldFile.DataFormat, true);

        newFile.OverwriteAudio(0, 0, oldFile, new SfAudioSelection(oldFile));

        long leadSilence = newFile.SecondsToPosition(silences[0]);
        long trailSilence = newFile.SecondsToPosition(silences[1]);

        if (leadSilence > 0)
            newFile.InsertSilence(0, leadSilence);
        if (trailSilence > 0)
            newFile.InsertSilence(newFile.Length, trailSilence);
            
        //next line could fail if user does not have permission to create directory. Implement try/catch
        string outName = Path.Combine(outPutDir, Path.GetFileName(audioFileName));
        newFile.SaveAs(outName, oldFile.SaveFormat.Guid, "Default Template", RenderOptions.WaitForDoneOrCancel);

        oldFile.Close(CloseOptions.DiscardChanges);
        newFile.Close(CloseOptions.SaveChanges);
    }
}

public class EntryPoint
{
    public void Begin(IScriptableApp app)
    {
        Form1 theForm = new Form1();
        theForm.InitializeComponent();
        theForm.addSfApp(app);
        theForm.ShowDialog();
        
        //Perform script operations here.
    }


    public void FromSoundForge(IScriptableApp app)
    {
        ForgeApp = app; //execution begins here
        app.SetStatusText(String.Format("Script '{0}' is running.", Script.Name));
        Begin(app);
        app.SetStatusText(String.Format("Script '{0}' is done.", Script.Name));
    }
    public static IScriptableApp ForgeApp = null;
    public static void DPF(string sz) { ForgeApp.OutputText(sz); }
    public static void DPF(string fmt, params object[] args) { ForgeApp.OutputText(String.Format(fmt, args)); }
} //EntryPoint



