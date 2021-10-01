# This software was developed at the National Institute of Standards and Technology by employees of
# the Federal Government in the course of their official duties.  Pursuant to Title 17 Section 105 of
# the United States Code this software is not subject to copyright protection and is in the public
# domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for
# its use by other parties, and makes no guarantees, expressed or implied, about its quality,
# reliability, or any other characteristic.

# This software is provided by NIST as a public service.  You may use, copy and distribute copies of
# the software in any medium, provided that you keep intact this entire notice.  You may improve,
# modify and create derivative works of the software or any portion of the software, and you may copy
# and distribute such modifications or works.  Modified works should carry a notice stating that you
# changed the software and should note the date and nature of any such change.  Please explicitly
# acknowledge NIST as the source of the software.

# See the NIST Disclaimer at https://www.nist.gov/disclaimer
# The latest version of the source code is available at: https://github.com/usnistgov/IFA

# This is the main routine for the IFC File Analyzer GUI version

global env tcl_platform

set scriptName [info script]
set wdir [file dirname [info script]]
set auto_path [linsert $auto_path 0 $wdir]

# for building your own version with freewrap, uncomment and modify C:/Tcl/lib/teapot directory if necessary
# the lappend commands add package locations to auto_path, must be before package commands below
# see 20 lines below for two more lappend commands
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itk3.4
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itcl3.4
#lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/Iwidgets4.0.2

# Tcl packages, check if they will load
if {[catch {
  package require Iwidgets 4.0
  package require Tclx
  package require tcom
  package require twapi
} emsg]} {
  set dir $wdir
  set c1 [string first [file tail [info nameofexecutable]] $dir]
  if {$c1 != -1} {set dir [string range $dir 0 $c1-1]}
  if {[string first "couldn't load library" $emsg] != -1} {
    append emsg "\n\nAlthough the message above indicates that a library is missing, that is NOT the root cause of the problem.  The problem is usually related to:"
    append emsg "\n\n1 - the directory you are running the software from has accented, non-English, or symbol characters in the pathname\n    [file nativename $dir]"
    append emsg "\n2 - permissions to run the software in the directory"
    append emsg "\n3 - other computer configuration problems"
    append emsg "\n\nTry the following workarounds to run the software:"
    append emsg "\n\n1 - from a directory without any special characters in the pathname, or from your home directory, or desktop"
    append emsg "\n2 - as Administrator"
    append emsg "\n3 - on a different computer"
  }
  append emsg "\n\nPlease send a screenshot of this dialog to Robert Lipman (robert.lipman@nist.gov) if you cannot run the IFC File Analyzer."
  set choice [tk_messageBox -type ok -icon error -title "ERROR running the IFC File Analyzer" -message $emsg]
  exit
}

# for building your own version with freewrap, also uncomment and modify the lappend commands
catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/tooltip1.4.4
  package require tooltip
}

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# detect if NIST version
set nistVersion 0
foreach item $auto_path {if {[string first "IFC-File-Analyzer" $item] != -1} {set nistVersion 1}}

# -----------------------------------------------------------------------------------------------------
# initialize variables
foreach id {XL_OPEN INVERSE SORT PR_BEAM PR_PROF PR_PROP PR_HVAC PR_UNIT PR_COMM PR_RELA \
            PR_ELEC PR_QUAN PR_REPR PR_SRVC PR_ANAL PR_PRES PR_MTRL PR_INFR PR_GEOM EX_PROP} {set opt($id) 1}

foreach id {COUNT HIDELINKS PR_USER XL_FPREC XL_KEEPOPEN EX_LP EX_A2P3D EX_ANAL} {set opt($id) 0}

set opt(DEBUGINV) 0
set opt(XLSCSV) "Excel"

set edmWriteToFile 0
set edmWhereRules 0
set eeWriteToFile  0

set userWriteDir $mydocs
set writeDir ""
set writeDirType 0
set row_limit 503

set openFileList {}
set fileDir  $mydocs
set fileDir1 $mydocs
set optionsFile [file nativename [file join $fileDir IFC-File-Analyzer-options.dat]]

set filemenuinc 4
set lenlist 25
set upgrade 0

set writeDir $userWriteDir

set dispCmd ""
set dispCmds {}

set lastXLS  ""
set lastXLS1 ""
set ifaVersion 0

# initialize data
setData_IFC

# -----------------------------------------------------------------------------------------------------
# check for options file and source
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
  } emsg]} {
    set endMsg "Error reading Options file [truncFileName $optionsFile]: $emsg"
  }
}

if {[info exists verite]}       {set ifaVersion $verite}
if {[info exists writeDirType]} {if {$writeDirType == 1} {set writeDirType 0}}
if {[info exists userWriteDir]} {if {![file exists $userWriteDir]} {set userWriteDir $mydocs}}
if {[info exists fileDir]}      {if {![file exists $fileDir]}      {set fileDir      $mydocs}}
if {[info exists fileDir1]}     {if {![file exists $fileDir1]}     {set fileDir1     $mydocs}}

# fix row limit
if {$row_limit < 103 || ([string range $row_limit end-1 end] != "03" && \
   [string range $row_limit end-1 end] != "76" && [string range $row_limit end-1 end] != "36")} {set row_limit 103}

foreach item {EX_ARBP FN_APPEND PR_TYPE XL_XLSX XL_LINK1 XL_LINK2 XL_LINK3 XL_ORIENT XL_SCROLL XL_KEEPOPEN writeDirType} {catch {unset opt($item)}}
foreach item {verite firsttime flag(FIRSTTIME)} {catch {unset $item}}

# -------------------------------------------------------------------------------
# get programs that can open IFC files
getDisplayPrograms

#-------------------------------------------------------------------------------
# user interface
guiStartWindow

# top menu
set Menu [menu .menubar]
. config -men .menubar
foreach m {File Websites Help} {
  set $m [menu .menubar.m$m -tearoff 1]
  .menubar add cascade -label $m -menu .menubar.m$m
}

# check if menu font is Segoe UI
catch {
  set ff [join [$File cget -font]]
  if {[string first "Segoe" $ff] == -1} {
    $File     configure -font [list {Segoe UI}]
    $Websites configure -font [list {Segoe UI}]
    $Help     configure -font [list {Segoe UI}]
  }
}

# file menu
guiFileMenu

set progtime 0
foreach item {ifa ifa_gen ifa_proc ifa_ent ifa_data ifa_indent ifa_gui ifa_multi ifa_attr ifa_inv ifa_ifc} {
  set fname [file join $wdir $item.tcl]
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

#-------------------------------------------------------------------------------
# Help and Websites menu
guiHelpMenu
guiWebsitesMenu

# tabs
set nb [ttk::notebook .tnb]
pack $nb -fill both -expand true

# status tab
guiStatusTab

# options tab
guiProcess

# inverse relationships and expand placement
guiInverseExpand

# display option
guiDisplayResult
pack $fopt -side top -fill both -expand true -anchor nw

# spreadsheet tab
guiSpreadsheet

# generate logo, progress bars
guiButtons

# switch to options tab (any text output will switch back to the status tab)
.tnb select .tnb.opt

# error messages from before GUI was available
if {[info exists endMsg]} {
  outputMsg " "
  errorMsg $endMsg
  .tnb select .tnb.status
}

#-------------------------------------------------------------------------------
# first time user
if {$ifaVersion == 0} {
  helpOverview
  set ifaVersion [getVersion]
  setShortcuts
  saveState

# what's new message
} elseif {$ifaVersion < [getVersion]} {
  set newstr {}
  set newifc {}
  if {$ifaVersion < 3.0} {lappend newstr "- Many minor improvements"}
  if {$ifaVersion < 3.01} {
    lappend newstr "- Renamed spreadsheets from 'myfile_ifc.xlsx' to 'myfile-ifa.xlsx'"
    lappend newstr "- Help > Function Keys"
    lappend newifc IFC4X2
    lappend newifc IFC4X3_RC4
  }
  if {$ifaVersion < 3.03} {lappend newifc IFC4X3_DEV}
  if {[llength $newifc] > 0} {lappend newstr "- Support for [join $newifc], see Help > IFC Support"}
  if {[llength $newstr] > 0} {
    outputMsg "\nWhat's New (Version: [getVersion]  Updated: [string trim [clock format $progtime -format "%e %b %Y"]])" blue
    foreach str $newstr {outputMsg $str}
    .tnb select .tnb.status
    update idletasks
  }
  set ifaVersion [getVersion]
  setShortcuts
  saveState
}
update idletasks

#-------------------------------------------------------------------------------
# install IFCsvr
installIFCsvr
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]

focus .

# check command line arguments or drag-and-drop
if {$argv != ""} {
  set localName [lindex $argv 0]
  if {[file dirname $localName] == "."} {
    set localName [file join [pwd] $localName]
  }
  if {$localName != ""} {
    set localNameList [list $localName]
    outputMsg "Ready to process: [file tail $localName] ([fileSize $localName])" blue
    $buttons(genExcel) configure -state normal
    $buttons(appDisplay) configure -state normal
    focus $buttons(genExcel)
    set fext [string tolower [file extension $localName]]
  }
}

set writeDir $userWriteDir
checkValues

set pid2 [twapi::get_process_ids -name "IFC-File-Analyzer.exe"]
set anapid $pid2
global anapid

if {[llength $pid2] > 1} {
  set msg "There are at least ([expr {[llength $pid2]-1}]) other instances of the IFC File Analyzer already running.\n\nDo you want to close them?"
  set choice [tk_messageBox -type yesno -default yes -message $msg -icon question -title "Close?"]
  if {$choice == "yes"} {
    foreach pid $pid2 {
      if {$pid != [pid]} {catch {twapi::end_process $pid -force}}
    }
    outputMsg "Other IFC File Analyzers closed" red
  }
}

if {$writeDirType == 2} {
  outputMsg " "
  errorMsg "Output files will be written to a User-Defined directory (Spreadsheet tab)"
  .tnb select .tnb.status
}

# warn about output type
if {$opt(XLSCSV) == "CSV"} {
  outputMsg " "
  errorMsg "CSV files will be generated (Options tab)"
  .tnb select .tnb.status
}

# set window minimum size
update idletasks
wm minsize . [winfo reqwidth .] [expr {int([winfo reqheight .]*1.05)}]

#debugData
