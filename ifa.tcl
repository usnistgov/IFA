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

global env tcl_platform

set scriptName [info script]
set wdir [file dirname [info script]]
set auto_path [linsert $auto_path 0 $wdir]

# for building your own version with freewrap, uncomment and modify C:/Tcl/lib/teapot directory if necessary
# the lappend commands add package locations to auto_path, must be before package commands below
# see 20 lines below for two more lappend commands
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itk3.4
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itcl3.4
lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/Iwidgets4.0.2

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
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/tooltip1.4.4
  package require tooltip
}

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

# -----------------------------------------------------------------------------------------------------
# IFC specific data
setData_IFC

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

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

# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# detect if NIST version
set nistVersion 0
foreach item $auto_path {if {[string first "IFC-File-Analyzer" $item] != -1} {set nistVersion 1}}

set flag(FIRSTTIME) 1
set lastXLS  ""
set lastXLS1 ""
set verite 0

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

# check if menu font is Segoe UI on Windows 7
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

#-------------------------------------------------------------------------------
# file menu

guiFileMenu

#-------------------------------------------------------------------------------
# Help menu
 
set progtime 0
foreach item {ifa ifa_gen ifa_proc ifa_ent ifa_data ifa_indent ifa_gui ifa_multi ifa_attr ifa_inv ifa_ifc} {
  set fname [file join $wdir $item.tcl]
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

#proc whatsNew {} {
#  global progtime verite mydocs 
#  
#  if {$verite > 0 && $verite < [getVersion]} {outputMsg "\nThe previous version of the IFC File Analyzer was: $verite" red}
#
#outputMsg "\nWhat's New (v[getVersion])" blue
#outputMsg "- Support for IFC4 although new entities in addendums are not supported
#- Support for CSV file output (Options tab)
#- Open file in Default IFC Viewer (Options tab)"
#
#  .tnb select .tnb.status
#  update idletasks
#}

guiHelpMenu

#-------------------------------------------------------------------------------
# Websites menu

$Websites add command -label "IFC File Analyzer"                          -command {displayURL https://www.nist.gov/services-resources/software/ifc-file-analyzer}                                                               
$Websites add command -label "Journal of NIST Research (citation)"        -command {displayURL https://dx.doi.org/10.6028/jres.122.015}
$Websites add command -label "Developing Coverage Analysis for IFC Files" -command {displayURL https://www.nist.gov/publications/developing-coverage-analysis-ifc-files}                                              
$Websites add command -label "Assessment of Conformance and Interoperability Testing Methods" -command {displayURL https://www.nist.gov/publications/assessment-conformance-and-interoperability-testing-methods-used-construction-industry}                                              
$Websites add separator
$Websites add command -label "buildingSMART"           -command {displayURL https://www.buildingsmart.org/}                                              
$Websites add command -label "IFC Technical Resources" -command {displayURL https://technical.buildingsmart.org/}                                                 
$Websites add command -label "IFC Documentation"       -command {displayURL https://technical.buildingsmart.org/standards/ifc/ifc-schema-specifications/}                     
$Websites add command -label "IFC Implementations"     -command {displayURL https://technical.buildingsmart.org/community/software-implementations/}                             
$Websites add command -label "Free IFC Software"       -command {displayURL http://www.ifcwiki.org/index.php?title=Freeware}                                                   
$Websites add command -label "Common BIM Files"        -command {displayURL https://www.nibs.org/page/bsa_commonbimfiles}                                          
#$Websites add command -label "IFCsvr toolkit"           -command {displayURL https://groups.yahoo.com/neo/groups/ifcsvr-users/info}

#-------------------------------------------------------------------------------
# tabs
set nb [ttk::notebook .tnb]
pack $nb -fill both -expand true

#-------------------------------------------------------------------------------
# status tab

guiStatusTab

#-------------------------------------------------------------------------------
# options tab

set cb 0
set wopt [ttk::panedwindow $nb.opt -orient horizontal]
$nb add $wopt -text " Options " -padding 2
set fopt [frame $wopt.fopt -bd 2 -relief sunken]

set fopta [ttk::labelframe $fopt.a -text " Process "]

# option to process user-defined entities
guiUserDefinedEntities

set fopta1 [frame $fopta.1 -bd 0]
foreach item {{" Building Elements" opt(PR_BEAM)} \
              {" HVAC"              opt(PR_HVAC)} \
              {" Electrical"        opt(PR_ELEC)} \
              {" Building Services" opt(PR_SRVC)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] \
    -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  set tt [string range $idx 3 end]
  if {[info exists type($tt)]} {
    set ttmsg "There are [llength $type($tt)] [string trim [lindex $item 0]] entities.  These entities are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttmsg [processToolTip $ttmsg $tt]
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  }
}
pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y

set fopta2 [frame $fopta.2 -bd 0]
foreach item {{" Structural Analysis" opt(PR_ANAL)} \
              {" Profile"             opt(PR_PROF)} \
              {" Material"            opt(PR_MTRL)} \
              {" Property"            opt(PR_PROP)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] \
    -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  set tt [string range $idx 3 end]
  if {[info exists type($tt)]} {
    set ttmsg "There are [llength $type($tt)] [string trim [lindex $item 0]] entities.  These entities are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttmsg [processToolTip $ttmsg $tt]
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  } elseif {[lindex $item 0] == " Material"} {
    set ttmsg "These are [string trim [lindex $item 0]] entities.  They are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttlen 0
    foreach item [lsort $ifcall] {
      if {[string first "Materia" $item] != -1 && \
          [string first "Propert" $item] == -1 && \
          [string first "IfcRel" $item] == -1 && [string first "Relationship" $item] == -1} {
        append ttmsg "$item   "
        incr ttlen [string length $item]
        if {$ttlen > 80} {
          append ttmsg "\n"
          set ttlen 0
        }
        lappend ifcProcess $item
      }
    }
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  } elseif {[lindex $item 0] == " Property"} {
    set ttmsg "These are [string trim [lindex $item 0]] entities.  They are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttlen 0
    foreach item [lsort $ifcall] {
      if {([string first "Propert" $item] != -1 || \
           [string first "IfcDoorStyle" $item] == 0 || \
           [string first "IfcWindowStyle" $item] == 0) && \
           [string first "IfcRel" $item] == -1 && [string first "Relationship" $item] == -1} {
        append ttmsg "$item   "
        incr ttlen [string length $item]
        if {$ttlen > 80} {
          append ttmsg "\n"
          set ttlen 0
        }
        lappend ifcProcess $item
      }
    }
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  }
}
pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y

set fopta3 [frame $fopta.3 -bd 0]
foreach item {{" Representation"  opt(PR_REPR)} \
              {" Relationship"    opt(PR_RELA)} \
              {" Presentation" opt(PR_PRES)} \
              {" Other"        opt(PR_COMM)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] \
    -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  set tt [string range $idx 3 end]
  if {[info exists type($tt)]} {
    set ttmsg "There are [llength $type($tt)] [string trim [lindex $item 0]] entities.  These entities are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttmsg [processToolTip $ttmsg $tt]
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  } elseif {[lindex $item 0] == " Relationship"} {
    set ttmsg "These are [string trim [lindex $item 0]] entities.  They are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttlen 0
    foreach item [lsort $ifcall] {
      if {[string first "Relationship" $item] != -1 || \
          [string first "IfcRel" $item] == 0} {
        append ttmsg "$item   "
        incr ttlen [string length $item]
        if {$ttlen > 80} {
          append ttmsg "\n"
          set ttlen 0
        }
        lappend ifcProcess $item
      }
    }
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  }
}
pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y

set fopta4 [frame $fopta.4 -bd 0]
foreach item {{" Geometry"     opt(PR_GEOM)} \
              {" Quantity"     opt(PR_QUAN)} \
              {" Unit"         opt(PR_UNIT)} \
              {" Include GUID" opt(PR_GUID)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $fopta4.$cb -text [lindex $item 0] \
    -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  set tt [string range $idx 3 end]
  if {[info exists type($tt)]} {
    set ttmsg "There are [llength $type($tt)] [string trim [lindex $item 0]] entities.  These entities are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    if {$tt == "PR_GEOM"} {append ttmsg "For large IFC files, this option can slow down the processing of the file and increase the size of the spreadsheet.\nUse the Count Duplicates and/or Maximum Rows options to speed up the processing Geometry entities.\n\n"}
    set ttmsg [processToolTip $ttmsg $tt]
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  } elseif {[lindex $item 0] == " Quantity"} {
    set ttmsg "These are [string trim [lindex $item 0]] entities.  They are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttlen 0
    foreach item [lsort $ifcall] {
      if {[string first "Quantit" $item] != -1} {
        append ttmsg "$item   "
        incr ttlen [string length $item]
        if {$ttlen > 80} {
          append ttmsg "\n"
          set ttlen 0
        }
        lappend ifcProcess $item
      }
    }
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  } elseif {[lindex $item 0] == " Unit"} {
    set ttmsg "These are [string trim [lindex $item 0]] entities.  They are found in IFC2x3 and/or IFC4.  IFC4.0.n addendums and IFC4.n versions are not supported.\nSee Websites > IFC Documentation\n\n"
    set ttlen 0
    foreach item [lsort $ifcall] {
      if {([string first "Unit" $item] != -1 && \
           [string first "Protective" $item] == -1 && \
           [string first "Unitary" $item] == -1) || [string first "DimensionalExponents" $item] != -1} {
        append ttmsg "$item   "
        incr ttlen [string length $item]
        if {$ttlen > 80} {
          append ttmsg "\n"
          set ttlen 0
        }
        lappend ifcProcess $item
      }
    }
    catch {tooltip::tooltip $buttons($idx) $ttmsg}
  }
}
catch {tooltip::tooltip $buttons(optPR_GUID) "Include the Globally Unique Identifier (GUID) and\nIfcOwnerHistory for each entity in a worksheet.\n\nThe GUID is checked for uniqueness."}
pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y

pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both

#-------------------------------------------------------------------------------
# inverse relationships

guiInverse

#-------------------------------------------------------------------------------
# count duplicates, sort entities

set foptbf  [frame $fopt.bf -bd 0]
guiDuplicates
#guiSort
pack $foptbf -side top -anchor w -pady 0 -fill x

#-------------------------------------------------------------------------------
# expand

set foptd [ttk::labelframe $fopt.1 -text " Expand "]
set foptd1 [frame $foptd.1 -bd 0]
foreach item {{" IfcLocalPlacement" opt(EX_LP)} \
              {" IfcAxis2Placement" opt(EX_A2P3D)} \
              {" Include Structural Analysis entities" opt(EX_ANAL)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] \
    -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
}
pack $foptd1 -side left -anchor w -pady 0 -padx 0 -fill y
pack $foptd -side top -anchor w -pady {5 2} -padx 10 -fill both
catch {tooltip::tooltip $foptd "These options will expand the selected entity attributes that are referred to on an entity being processed.\n\nFor example, selecting IfcLocalPlacement will show the attribute values of PlacementRelTo and RelativePlacement for\nIfcLocalPlacement for every building element.\nExpanding IfcAxis2Placement will show the corresponding attribute values for Location, Axis, and RefDirection.\n\nThis option does not work well where building elements of the same type have different levels of coordinate system nesting.\n\nExpanding Structural Analysis entities also applies to loads, reactions, and displacements.\n\nThe columns used for the expanded entities are grouped together and displayed with different colors.\nUse the \"-\" symbols above the columns or the \"1\" at the top left of the spreadsheet to collapse the columns."}

#-------------------------------------------------------------------------------
# max rows
# guiMaxRows

#-------------------------------------------------------------------------------
# display option

guiDisplayResult

pack $fopt -side top -fill both -expand true -anchor nw

#-------------------------------------------------------------------------------
# spreadsheet tab

guiSpreadsheet
pack $fxls -side top -fill both -expand true -anchor nw

#-------------------------------------------------------------------------------
# generate button and images, can't put the actual button in the proc, causes error generating file summary spreadsheet

if {$tcl_platform(osVersion) < 6.0} {
  set ftrans [frame .ftrans1 -bd 2 -background "#E0DFE3"]
} else {
  set ftrans [frame .ftrans1 -bd 2 -background "#F0F0F0"]
}
set buttons(genExcel) [ttk::button $ftrans.generate1 -text "Generate Spreadsheet" -padding 4 \
  -state disabled -command {
    saveState
    if {![info exists localNameList]} {
      set localName [getFirstFile]
      if {$localName != ""} {
        set localNameList [list $localName]
        genExcel
      }
    } elseif {[llength $localNameList] == 1} {
      genExcel
    } else {
      openMultiFile 2
    }
  }]
pack $ftrans.generate1 -side left -padx 10

guiButtons

# switch to options tab (any text output will switch back to the status tab)
.tnb select .tnb.opt

#-------------------------------------------------------------------------------
# first time user
if {$flag(FIRSTTIME)} {
  helpOverview
  #whatsNew
  displayDisclaimer

  set verite [getVersion]
  set flag(FIRSTTIME) 0
  saveState
  setShortcuts
  
  outputMsg " "
  #errorMsg "Use F6 and F5 to change the font size.  Right-click to save the text."
  saveState

# what's new message
} elseif {$verite < [getVersion]} {
  #whatsNew

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
    } else {
      #tk_messageBox -type ok -default ok -title "Check for Update" \
      #  -message "You can always check for an update by going to\nHelp > Check for Update" -icon info
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

#-------------------------------------------------------------------------------
#proc compareLists {str l1 l2} {
#  set l3 [intersect3 $l1 $l2]
#  outputMsg "\n$str" red
#  outputMsg "Unique to L1   ([llength [lindex $l3 0]])\n  [lindex $l3 0]"
#  outputMsg "Common to both ([llength [lindex $l3 1]])\n  [lindex $l3 1]"
#  outputMsg "Unique to L2   ([llength [lindex $l3 2]])\n  [lindex $l3 2]"
#}
#
#compareLists "all to process" $ifcall $ifcProcess
