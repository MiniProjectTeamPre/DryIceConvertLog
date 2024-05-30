using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DryIceConvertLog {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private ConvertDryice convert = new ConvertDryice();
        private PathDryice pathDryice = new PathDryice();
        private VarInput varInput = new VarInput();


        private void button1_Click(object sender, EventArgs e) {
            bool flagRead = varInput.ListExcelFilesInInputFolder(pathDryice.PathInput);

            if (flagRead)
            {
                string inputFilePath = Path.Combine(pathDryice.PathInput, varInput.Names[0]);

                string outputFilePath = Path.Combine(pathDryice.PathOutput, varInput.Names[0]);
                outputFilePath = Path.ChangeExtension(outputFilePath, ".csv");

                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                }

                bool flagProcess = convert.Process(varInput.Datas, outputFilePath);

                if (flagProcess)
                {
                    if (File.Exists(inputFilePath))
                    {
                        File.Delete(inputFilePath);
                    }
                    else
                    {
                        MessageBox.Show("Input file does not exist.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    MessageBox.Show("Process completed successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Error occurred during the process.", "Process Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Error occurred while reading the input files.", "Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e) {


        }
    }
}
