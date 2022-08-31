# NIST IFC File Analyzer

A pre-built Windows version of NIST IFC File Analyzer (IFA) is available [here](https://www.nist.gov/services-resources/software/ifc-file-analyzer).  

These are the instructions for building the NIST IFC File Analyzer from the source code.  IFA generates a spreadsheet from an [IFC](https://technical.buildingsmart.org/) file.

## Prerequisites

The IFC File Analyzer can only be built and run on Windows computers.  This is due to a dependence on the IFCsvr toolkit that is used to read and parse IFC files.  That toolkit only runs on Windows.

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.  

**You must install and run the NIST version of the IFC File Analyzer before running your own version.**

- Go to the [IFC File Analyzer](https://www.nist.gov/services-resources/software/ifc-file-analyzer) to download the software
- Extract IFC-File-Analyzer.exe from the zip file and run it.  This will install the IFCsvr toolkit that is used to read IFC files.  The toolkit only runs on Windows.

Download the IFA files from GitHub to a directory on your computer.

- The name of the directory is not important
- The IFC File Analyzer is written in [Tcl](https://www.tcl.tk/)
- Some of the Tcl code is based on [CAWT](http://www.cawt.tcl3d.org/)

freeWrap wraps the IFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>.  More recent versions of freeWrap will **not** work with the IFA.
- Extract freewrap.exe and put it in the same directory as the IFA files that were downloaded from GitHub.

Several Tcl packages not included in freeWrap also need to be installed.

- teapot.zip contains the additional Tcl packages
- Create a directory C:/Tcl/lib
- Unzip teapot.zip to the 'lib' directory to create C:/Tcl/lib/teapot

## Build the IFC File Analyzer

Edit the source code file ifa.tcl and uncomment the lines at the top of the file that start with 'lappend auto_path C:/Tcl/lib/teapot/package/...'

Open a command prompt window and change to the directory with the IFA Tcl files and freewrap.  To create the executable ifa.exe, enter the command:

```
freewrap -f ifa-files.txt
```

To build the command-line version:

- Download freewrapTCLSH.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>
- Extract freewrapTCLSH.exe to the directory with the IFA Tcl files
- Edit ifa-files.txt and change the first line 'ifa.tcl' to 'ifa_cl.tcl'
- Edit ifa_cl.tcl similar to ifa.tcl above
- To create ifa_cl.exe, enter the command: freewrapTCLSH -f ifa-files.txt

## Other

- Some features are not available in the user-built version including tooltips and unzipping compressed IFC files.
- Internally at NIST, IFA is built with [ActiveTcl 8.5.18 32-bit](https://www.activestate.com/products/tcl/) and the [Tcl Dev Kit](https://www.activestate.com/blog/tcl-dev-kit-now-open-source/) which is now an open source project.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
