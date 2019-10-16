# This software was developed at the National Institute of Standards and Technology by employees of 
# the Federal Government in the course of their official duties.  Pursuant to Title 17 Section 105 
# of the United States Code this software is not subject to copyright protection and is in the 
# public domain. This software is an experimental system.  NIST assumes no responsibility whatsoever 
# for its use by other parties, and makes no guarantees, expressed or implied, about its quality, 
# reliability, or any other characteristic.  We would appreciate acknowledgement if the software is 
# used.
# 
# This software can be redistributed and/or modified freely provided that any derivative works bear 
# some notice that they are derived from it, and any modified versions bear some notice that they 
# have been modified. 

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
  set choice [tk_messageBox -type ok -icon error -title "ERROR" -message "ERROR: $emsg\n\nThere might be a problem running this program from a directory with accented, non-English, or symbol characters in the pathname.\n\n[file nativename $dir]\n\nRun the software from a directory without any of the special characters in the pathname.\n\nPlease contact Robert Lipman (robert.lipman@nist.gov) for other problems."]
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
# initialize variables, set opt to 1
foreach id {XL_OPEN XL_LINK1 XL_FPREC EX_A2P3D EX_LP EX_ANAL COUNT INVERSE SORT PR_USER \
            PR_BEAM PR_PROF PR_PROP PR_GUID PR_HVAC PR_UNIT PR_COMM PR_RELA \
            PR_ELEC PR_QUAN PR_REPR PR_SRVC PR_ANAL PR_PRES PR_MTRL PR_GEOM} {set opt($id) 1}

set opt(COUNT) 0
set opt(SORT) 1

set opt(PR_GUID) 0
set opt(PR_GEOM) 0
set opt(PR_USER) 0

set opt(XL_FPREC) 0
set opt(XL_KEEPOPEN) 0
set opt(FN_APPEND) 0

set opt(EX_LP)    0
set opt(EX_A2P3D) 0
set opt(EX_ANAL)  0

set opt(DEBUGINV) 0

set opt(XLSCSV) "Excel"

set edmWriteToFile 0
set edmWhereRules 0
set eeWriteToFile  0

set userWriteDir $mydocs
set writeDir ""
set writeDirType 0
set maxfiles 1000
set row_limit 1003

set openFileList {}
set fileDir  $mydocs
set fileDir1 $mydocs
set optionsFile [file nativename [file join $fileDir IFC-File-Analyzer-options.dat]]

set filemenuinc 4
set lenlist 25
set upgrade 0
set upgradeIFCsvr 0
set yrexcel ""

set writeDir $userWriteDir

set userXLSFile ""

set dispCmd ""
set dispCmds {}

set flag(FIRSTTIME) 1
set lastXLS  ""
set lastXLS1 ""
set verite 0

# initialize data
setData_IFC

# -----------------------------------------------------------------------------------------------------
# check for options file and source
set optionserr ""
if {[file exists $optionsFile]} {
  catch {source $optionsFile} optionserr
  if {[string first "+" $optionserr] == 0} {set optionserr ""}
  foreach item {PR_TYPE XL_XLSX XL_LINK2 XL_LINK3 XL_ORIENT XL_SCROLL XL_KEEPOPEN} {
    catch {unset opt($item)}
  }
}
if {[info exists userWriteDir]} {if {![file exists $userWriteDir]} {set userWriteDir $mydocs}}
if {[info exists fileDir]}      {if {![file exists $fileDir]}      {set fileDir      $mydocs}}
if {[info exists fileDir1]}     {if {![file exists $fileDir1]}     {set fileDir1     $mydocs}}
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}
if {[info exists firsttime]} {set flag(FIRSTTIME) $firsttime}
if {$row_limit < 103 || ([string range $row_limit end-1 end] != "03" && \
   [string range $row_limit end-1 end] != "76" && [string range $row_limit end-1 end] != "36")} {set row_limit 103}

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

# check if menu font is Segoe UI on Windows 7 or greater
catch {
  if {$tcl_platform(osVersion) >= 6.0} {
    set ff [join [$File cget -font]]
    if {[string first "Segoe" $ff] == -1} {
      $File     configure -font [list {Segoe UI}]
      $Websites configure -font [list {Segoe UI}]
      $Help     configure -font [list {Segoe UI}]
    }
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

# inverse relationships
guiInverse

# count duplicates
guiDuplicates

# expland placement
guiExpandPlacement

# display option
guiDisplayResult
pack $fopt -side top -fill both -expand true -anchor nw

# spreadsheet tab
guiSpreadsheet

# generate logo, progress bars
guiButtons

# switch to options tab (any text output will switch back to the status tab)
.tnb select .tnb.opt

#-------------------------------------------------------------------------------
# first time user
if {$flag(FIRSTTIME)} {
  helpOverview
  displayDisclaimer
  set verite [getVersion]
  set flag(FIRSTTIME) 0
  saveState
  setShortcuts
  outputMsg " "
  saveState

# what's new message
} elseif {$verite < [getVersion]} {
  set verite [getVersion]
  saveState
  setShortcuts
}
update idletasks

#-------------------------------------------------------------------------------
# check for update every year
if {$upgrade > 0} {
  set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
  if {$lastupgrade > 365} {
    set choice [tk_messageBox -type yesno -default yes -title "Check for Update" \
      -message "Do you want to check for a newer version of the IFC File Analyzer?\n \nThe last check for an update was $lastupgrade days ago." -icon question]
    if {$choice == "yes"} {
      set os "$tcl_platform(os) $tcl_platform(osVersion)"
      regsub -all " " $os "" os
      regsub "WindowsNT" $os "" os
      if {$pf64 != ""} {append os ".64"}
      set url "https://concrete.nist.gov/cgi-bin/ctv/ifa_upgrade.cgi?version=[getVersion]&auto=$lastupgrade&os=$os"
      if {[info exists yrexcel]} {if {$yrexcel != ""} {append url "&yr=[expr {$yrexcel-2000}]"}}
      displayURL $url
    }
    set upgrade [clock seconds]
    saveState
  }
} else {
  set upgrade [clock seconds]
  saveState
}

#-------------------------------------------------------------------------------
# install IFCsvr
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]
if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {
  installIFCsvr
} else {
  set ifcsvrTime [file mtime [file join $wdir exe ifcsvrr300_setup_1008_en-update.msi]]
  if {$ifcsvrTime > $upgradeIFCsvr} {installIFCsvr 1}
}

focus .

# check command line arguments or drag-and-drop
if {$argv != ""} {
  set localName [lindex $argv 0]
  if {[file dirname $localName] == "."} {
    set localName [file join [pwd] $localName]
  }
  if {$localName != ""} {
    set localNameList [list $localName]
    outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
    $buttons(genExcel) configure -state normal
    $buttons(appDisplay) configure -state normal
    focus $buttons(genExcel)
    set fext [string tolower [file extension $localName]]
  }
}

set writeDir $userWriteDir
checkValues

if {[string length $optionserr] > 5} {
  errorMsg "ERROR reading options file: $optionsFile\n $optionserr"
  errorMsg "Some previously saved options might be lost."
  .tnb select .tnb.status
}

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

if {$writeDirType == 1} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined file name (Spreadsheet tab)"
} elseif {$writeDirType == 2} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined directory (Spreadsheet tab)"
}

# set window minimum size
update idletasks
wm minsize . [winfo reqwidth .] [expr {int([winfo reqheight .]*1.05)}]
