# This is the main routine for the IFC File Analyzer command-line version

# Website - https://www.nist.gov/services-resources/software/ifc-file-analyzer
# NIST Disclaimer - https://www.nist.gov/disclaimer
# Source code - https://github.com/usnistgov/IFA

global env

set scriptName [info script]
set wdir [file dirname [info script]]
set auto_path [linsert $auto_path 0 $wdir]

puts "\n--------------------------------------------------------------------------------"
puts "NIST IFC File Analyzer [getVersion]"

if {[catch {
  package require Tclx
  package require tcom
  package require twapi
} emsg]} {
  set dir $wdir
  set c1 [string first [file tail [info nameofexecutable]] $dir]
  if {$c1 != -1} {set dir [string range $dir 0 $c1-1]}
  if {[string first "couldn't load library" $emsg] != -1} {
    append emsg "\n\nAlthough the message above indicates that a library is missing, that is NOT the cause of the problem.  The problem is sometimes related to the directory where you are running the software.\n\n   [file nativename $dir]"
    append emsg "\n\n1 - The directory has accented, non-English, or symbol characters"
    append emsg "\n2 - The directory is on a different computer"
    append emsg "\n3 - No permissions to run the software in the directory"
    append emsg "\n4 - Other computer configuration problems"
    append emsg "\n\nTry these workarounds to run the software:"
    append emsg "\n1 - From a directory without any special characters in the pathname, or from your home directory, or desktop"
    append emsg "\n2 - Installed on your local computer"
    append emsg "\n3 - As Administrator"
    append emsg "\n4 - On a different computer"
  }
  puts "\nError: $emsg"
  exit
}

foreach id {XL_OPEN INVERSE SORT PR_BEAM PR_PROF PR_PROP PR_HVAC PR_UNIT PR_COMM PR_RELA \
            PR_ELEC PR_QUAN PR_REPR PR_SRVC PR_ANAL PR_PRES PR_MTRL PR_INFR PR_GEOM EX_PROP} {set opt($id) 1}

foreach id {COUNT HIDELINKS PR_USER XL_FPREC XL_KEEPOPEN EX_LP EX_A2P3D EX_ANAL} {set opt($id) 0}

set opt(DEBUGINV) 0
set opt(XLSCSV) Excel

# -----------------------------------------------------------------------------------------------------
# IFC pecific data
setData_IFC

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

set userWriteDir $mydocs
set writeDir ""
set writeDirType 0
set row_limit 503

set openFileList {}
set fileDir  $mydocs
set fileDir1 $mydocs

set filemenuinc 4
set lenlist 25

set writeDir $userWriteDir

set dispCmd ""
set dispCmds {}

# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}
set ifcsvrdir [file join $pf32 IFCsvrR300 dll]

set lastXLS  ""
set lastXLS1 ""
set ifaVersion 0

# read (source) options file
set optionsFile [file nativename [file join $fileDir IFC-File-Analyzer-options.dat]]
if {[file exists $optionsFile]} {
  catch {source $optionsFile}
}
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}

#-------------------------------------------------------------------------------
# check for IFCsvr
if {![file exists [file join $pf32 IFCsvrR300 dll IFCsvrR300.dll]]} {
  outputMsg " "
  errorMsg "IFCsvr needs to be installed for the IFC File Analyzer to read IFC files."
  outputMsg "\nInstall IFCsvr -------------------------------------------------------------" blue
  outputMsg " 1 - Run the GUI version of the IFC File Analyzer (IFC-File-Analyzer.exe)"
  outputMsg " 2 - Follow the instructions to install IFCsvr"
  outputMsg " 3 - Rerun this software"
  exit
}

# no arguments, no file, print help, and exit

if {$argc == 1} {set arg [string tolower [lindex $argv 0]]}
if {$argc == 0 || ($argc == 1 && ($arg == "help" || $arg == "-help" || $arg == "-h" || $arg == "-v"))} {
  puts "\nUsage: IFC-File-Analyzer-CL.exe myfile.ifc \[csv\] \[noopen\]

Optional command line settings:
  csv     Generate CSV files
  noopen  Do not open the Spreadhseet after it has been generated

 Most options last used in the GUI version are used in this program.  If 'myfile.ifc'
 has spaces, put quotes around the file name \"C:/my dir/my file.ifc\"

 When the IFC file is opened, sytnax errors and warnings might appear at the beginning
 of the output.  Existing Spreadsheets are always overwritten.

Disclaimers
 NIST Disclaimer: https://www.nist.gov/disclaimer

 This software uses IFCsvr and Microsoft Excel that are covered by their own Software
 License Agreements.  If you are using this software in your own application, please
 explicitly acknowledge NIST as the source of the software.

Credits
- Reading and parsing IFC files:
   IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
   IFCsvr has been modified by NIST to include IFC4x3
   The license agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc"

  exit
}

# get arguments and initialize variables
for {set i 1} {$i <= 100} {incr i} {
  set arg [string tolower [lindex $argv $i]]
  if {$arg != ""} {
    lappend larg $arg
    if {[string first "noopen" $arg] == 0} {set opt(XL_OPEN) 0}
    if {[string first "csv"    $arg] == 0} {set opt(XLSCSV) "CSV"}
  }
}

# IFC file name
set localName [lindex $argv 0]
if {[string first ":" $localName] == -1} {set localName [file join [pwd] $localName]}
set localName [file nativename $localName]
set remoteName $localName
set fext [string tolower [file extension $localName]]

if {[file exists $localName]} {
  genExcel
} else {
  outputMsg "File does not exist: [truncFileName $localName]"
}
