# PowerShell_fmwrk

The PowerShell Framework is intended to be used by a network
administrator to perform a number of tasks.

The focus of this framework is on simplicity and modularity

## Setup
The Setup folder contains a batch file for setting up the local environment.
This batch file calls PowerShell to download and install the appropriate
Windows 10 cmdlets.

## Usage
This framework is designed to be run out of a nested folder structure hierarchy.
You can re-organize your file structure as you see fit.

Each PowerShell script can be executed if the script begins with:
1 ExecuteOut      # Run the selected script immediately
2 ExecuteOutOpen  # Opens the script in ISE and executes
3 anything else   # Opens the script in ISE

Each PowerShell script should begin with the appropriate execution action followed
by a brief description of what the script does.

## Organization
The framework is designed to be easily re-organized as needed by the administrator.
The current organizational structure may be re-arranged if other configurations
make more sense. Each folder should be topic oriented, each script should be verb
oriented.

### Local
The Local folder contains scripts intended to be run from the framework dirrectly
on a user's PC. These scripts should work when copied to a flash drive.

### Remote
These scripts are intended to be run from the administrator's machine, and allow
the administrator to perform various tasks.
These scripts are grouped by:
1 Active Directory
2 Office 365
3 Direct
