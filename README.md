# NIST IFC File Analyzer

A pre-built Windows version of NIST IFC File Analyzer (IFA) is available [here](https://www.nist.gov/services-resources/software/ifc-file-analyzer).  

These are the instructions for building the NIST IFC File Analyzer from the source code.  IFA generates a spreadsheet from an [IFC](https://technical.buildingsmart.org/) file.

## Prerequisites

The IFC File Analyzer can only be built and run on Windows computers.  [Microsoft Excel](https://products.office.com/excel) is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.  

**You must install and run the NIST version of the IFC File Analyzer before running your own version.**

- Go to the [IFC File Analyzer](https://www.nist.gov/services-resources/software/ifc-file-analyzer) and click on the Attachment zip file to download the software
- Extract IFC-File-Analyzer.exe from the zip file and run it.  This will install the IFCsvr toolkit that is used to read IFC files.

Download the IFA files from GitHub to a directory on your computer.

- The name of the directory is not important
- The IFC File Analyzer is written in [Tcl](https://www.tcl.tk/)
- Some of the Tcl code is based on [CAWT](http://www.cawt.tcl3d.org/)

freeWrap wraps the IFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>.  More recent versions of freeWrap will **not** work with the IFA.
- Extract freewrap.exe and put it in the same directory as the IFA files that were downloaded from GitHub.

Install the ActiveTcl **8.5 32-bit** version of Tcl.

- Download the ActiveTcl installer from <https://www.activestate.com/products/activetcl/downloads/>.  You will have to create an ActiveState account.
- The Windows installer file name is: ActiveTcl-8.5.18.0.nnnnnn-win32-ix86-threaded.exe
- IFA can be built only with ActiveTcl 8.5.18 (32-bit).  ActiveTcl 8.6.n and 64-bit versions are not supported.
- Run the installer and use the default installation folders

Several Tcl packages from ActiveTcl also need to be installed.  Open a command prompt window, change to C:\\Tcl\\bin, or wherever Tcl was installed, and enter the following three commands:

```
teacup install tcom
teacup install twapi
teacup install Iwidgets
```

## Build the IFC File Analyzer

First, edit the source code file ifa.tcl and uncomment the lines at the top of the file that start with 'lappend auto_path C:/Tcl/lib/teapot/package/...'  Change 'C:/Tcl' if Tcl is installed in a different directory.

Then, open a command prompt window and change to the directory with the IFA Tcl files and freewrap.  To create the executable ifa.exe, enter the command:

```
freewrap -f ifa-files.txt
```

**Optionally, build the IFC File Analyzer command-line version**

- Download freewrapTCLSH.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>
- Extract freewrapTCLSH.exe to the directory with the IFA Tcl files
- Edit ifa-files.txt and change the first line 'ifa.tcl' to 'ifa_cl.tcl'
- Edit ifa_cl.tcl similar to ifa.tcl above
- To create ifa_cl.exe, enter the command: freewrapTCLSH -f ifa-files.txt

## Differences from the NIST-built version of IFC File Analyzer

Some features are not available in the user-built version including tooltips and unzipping compressed IFC files.  Some of the features are restored if the NIST-built version is run first.

## Contact

[Robert Lipman](https://www.nist.gov/people/robert-r-lipman), <robert.lipman@nist.gov>

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/public_affairs/disclaimer.cfm)
