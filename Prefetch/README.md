# PECmd_looper

Python 3 Script used for testing behavior in Microsoft Windows Prefetch files.

It utilizes the program PECmd for parsing Prefetch files.

## Prerequisites

- Python 3
- psutil Python module
- script parameters setup

## Script options (MENU)

- 0 - Copy Prefetch files from the default location
- 1 - Only analyze
- 2 - Only extract from json
- 3 - Analyze and extract
- 4 - View JSON files
- 5 - Compare JSON files (References)
- 8 - SpeedTest (0->1->2->delete)
- 9 - SpeedTest x 10 loops (starts and exits Chrome)
- delete - Delete all files in the folder
- X - Exit

Note: Option 9 can crush Google Chrome.

## Paramteters to set

- folder - Where files will be copied and where PECmd will look for Prefetch files
- source - Location of Prefetch files (only option 0)
- pf_copy_prefix - Prefetch file prefix to find e.g. name of the program "CHROME.EXE"
- program_path - path to program executable (only option 9)
- program_name - Name of program to start/close from task manager (only option 9)

# PECmd
Please note that the script uses the tool PECmd created by Eric Zimmerman (saericzimmerman@gmail.com) 

Source: https://github.com/EricZimmerman/PECmd
