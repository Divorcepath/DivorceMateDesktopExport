# DivorceMateDesktopExport

This repository provides a tool to export data from a DivorceMate database into CSV files using a VBS script.

## Prerequisites

Before running the export script, ensure that you have a backup of your DivorceMate database, which is typically saved in a folder such as `C://DIVORCEmate One/DMone.mdb`.

## How to Use

1. **Create a Backup**  
   Ensure you have a backup copy of your DivorceMate database. This is typically stored in a location such as:  
   `C://DIVORCEmate One/DMone.mdb`

2. **Run the VBS Script**  
   Open a command prompt and run the VBS script with the following command:

   ```bash
   cscript //nologo divorcemate-export.vbs "path/to/DIVORCEmate One/DMone.mdb" "output_directory"
   ```

   Replace "path/to/DIVORCEmate One/DMone.mdb" with the actual path to the backup copy of your DivorceMate database.

   Replace "output_directory" with the path where you would like the CSV files to be saved.

3. **Export Results**
   After running the script, CSV files will be created in the specified output directory, along with an export_log.txt file that contains details about the export process.

4. **Email the Export**
   Compress the CSV files and the export_log.txt file into a ZIP archive.
   Send the ZIP file via email to:
   help@divorcepath.com

**Notes**

Ensure the DivorceMate database is properly backed up before running the script.

The script will generate separate CSV files for the exported data, which can be easily reviewed or processed as needed.

**License**

This script is provided "as-is," without any representations or warranties of any kind, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or non-infringement. In no event shall the authors or contributors be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the use of this script.

We are not the licensor or trademark holder of DivorceMate, nor do we claim any affiliation with DivorceMate. DivorceMate is a trademark of its respective owners, and all rights associated with DivorceMate are owned by them.

By using this script, you agree to these terms.
