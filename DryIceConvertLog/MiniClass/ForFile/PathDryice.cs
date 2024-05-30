using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public class PathDryice {
    public PathDryice() {
        CheckAndCreateFolder(PathInput);
        CheckAndCreateFolder(PathOutput);
    }

    public string PathInput = @"D:\DryiceConvertLog\Input";
    public string PathOutput = @"D:\DryiceConvertLog\Output";

    private void CheckAndCreateFolder(string folderPath) {
        if (!Directory.Exists(folderPath))
        {
            try
            {
                Directory.CreateDirectory(folderPath);
            } catch (Exception ex)
            {
                MessageBox.Show($"Error creating folder: {folderPath}{Environment.NewLine}{ex.Message}");
            }
        }
    }
}
