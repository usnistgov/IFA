# NIST IFC File Analyzer

The [NIST IFC File Analyzer](https://www.nist.gov/services-resources/software/ifc-file-analyzer) (IFA) generates a spreadsheet from an [IFC](https://technical.buildingsmart.org/) file.
Download a pre-built Windows version of IFA with the Release link (zip file) to the right. 
Follow the instructions below to build your own version of IFA from the source code.  

## Prerequisites

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.  

Download the IFA files from GitHub to a directory on your computer.

- The name of the directory is not important
- The IFC File Analyzer is written in [Tcl](https://www.tcl.tk/)
- Some of the Tcl code is based on [CAWT](http://www.cawt.tcl3d.org/)

freeWrap wraps the IFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>.  More recent versions of freeWrap will **not** work with wrapping IFA.
- Extract freewrap.exe and put it in the same directory as the IFA files that were downloaded from GitHub.

Several Tcl packages not included in freeWrap also need to be installed.

- teapot.zip contains the additional Tcl packages
- Create a directory C:/Tcl/lib
- Unzip teapot.zip to the 'lib' directory to create C:/Tcl/lib/teapot

## Build the IFC File Analyzer

- Edit the source code file ifa.tcl and uncomment the lines at the top of the file that start with 'lappend auto_path C:/Tcl/lib/teapot/package/...'
- Open a command prompt window and change to the directory with the IFA Tcl files and freewrap.
- To generate the executable **ifa.exe**, enter the command: freewrap -f ifa-files.txt

Optionally build the command-line version:

- Download freewrapTCLSH.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>
- Extract freewrapTCLSH.exe to the directory with the IFA Tcl files
- Edit ifa-files.txt and change the first line 'ifa.tcl' to 'ifa_cl.tcl'
- Edit ifa_cl.tcl similar to ifa.tcl above
- To generate **ifa_cl.exe**, enter the command: freewrapTCLSH -f ifa-files.txt

## Running the Software

**You must install and run the NIST version of the IFC File Analyzer before running your own version.**
- Click on Release to the right and download the zip file.
- Extract IFC-File-Analyzer.exe from the zip file, run it and process an IFC file to install other software.
- Some features are not available in the user-built version including tooltips and unzipping compressed IFC files.
- Internally at NIST, IFA is built with [ActiveTcl 8.5.18 32-bit](https://www.activestate.com/products/tcl/) and the [Tcl Dev Kit](https://www.activestate.com/blog/tcl-dev-kit-now-open-source/) which is now an open source project.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
