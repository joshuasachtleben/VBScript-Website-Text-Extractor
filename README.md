# VBScript-Website-Text-Extractor
This project contains a VBScript file that navigates to a URL for the Federal Reserve Bank Services website that hosts an ePayments Routing Directory.  The script clicks an agreement acceptance button programmatically and downloads the text file hosted within the redirected page.

To execute this VBScript file, you will need to run the command "wscript text.vbs" from a command line in the same director that contains the script file.

The output of this script will be saved to the C drive unless the outFile variable is modified to a different path.  To write to the C drive on Windows, you may need to run your Command Prompt as an Administrator.  Just right-click on the Command Prompt icon and choose "Run as Administrator".
