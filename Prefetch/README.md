# PECmd_looper

Python 3 Script used for testing behavior in Microsoft Windows Prefetch files.

It utilizes the program PECmd for parsing Prefetch files.

## Prerequisites

- Python 3
- psutil Python module
- script parameters setup

## Paramteters to set

- folder - Where files will be copied and where PECmd will look for Prefetch files
- source - Location of Prefetch files (only option 0)
- pf_copy_prefix - Prefetch file prefix to find e.g. name of the program "CHROME.EXE"
- program_path - path to program executable (only option 9)
- program_name - Name of program to start/close from task manager (only option 9)

# PECmd
Please note that the script uses the tool PECmd created by Eric Zimmerman (saericzimmerman@gmail.com) 

Source: https://github.com/EricZimmerman/PECmd
