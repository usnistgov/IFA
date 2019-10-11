#-------------------------------------------------------------------------------
proc getVersion {} {
  set app_version 2.65
  return $app_version
}

#-------------------------------------------------------------------------------
# start window, bind keys
proc guiStartWindow {} {
  global winpos wingeo localName localNameList lastXLS lastXLS1 fout

  wm title . "IFC File Analyzer  (v[getVersion])"

# check that the saved window dimensions do not exceed the screen size
  if {[info exists wingeo]} {
    set gwid [lindex [split $wingeo "x"] 0]
    set ghgt [lindex [split $wingeo "x"] 1]
    if {$gwid > [winfo screenwidth  .]} {set gwid [winfo screenwidth  .]}
    if {$ghgt > [winfo screenheight .]} {set ghgt [winfo screenheight .]}
    set wingeo "$gwid\x$ghgt"
  }

# check that the saved window position is on the screen
  if {[info exists winpos]} {
    set pwid [lindex [split $winpos "+"] 1]
    set phgt [lindex [split $winpos "+"] 2]
    if {$pwid > [winfo screenwidth  .] || $pwid < -10} {set pwid 300}
    if {$phgt > [winfo screenheight .] || $phgt < -10} {set phgt 200}
    set winpos "+$pwid+$phgt"
  }

# check that the saved window position keeps the entire window on the screen
  if {[info exists wingeo] && [info exists winpos]} {
    if {[expr {$pwid+$gwid}] > [winfo screenwidth  .]} {
      set pwid [expr {[winfo screenwidth  .]-$gwid-40}]
      if {$pwid < 0} {set pwid 300}
    }
    if {[expr {$phgt+$ghgt}] > [winfo screenheight  .]} {
      set phgt [expr {[winfo screenheight  .]-$ghgt-40}]
      if {$phgt < 0} {set phgt 200}
    }
    set winpos "+$pwid+$phgt"
  }

# set the window position and dimensions
  if {[info exists winpos]} {catch {wm geometry . $winpos}}
  if {[info exists wingeo]} {catch {wm geometry . $wingeo}}

# yellow background color
  set bgcolor  "#ffffbb"
  catch {option add *Frame.background       $bgcolor}
  catch {option add *Label.background       $bgcolor}
  catch {option add *Checkbutton.background $bgcolor}
  catch {option add *Radiobutton.background $bgcolor}

  ttk::style configure TCheckbutton -background $bgcolor
  ttk::style map       TCheckbutton -background [list disabled $bgcolor]
  ttk::style configure TRadiobutton -background $bgcolor
  ttk::style map       TRadiobutton -background [list disabled $bgcolor]
  ttk::style configure TLabelframe       -background $bgcolor

  font create fontBold {*}[font configure TkDefaultFont]
  font configure fontBold -weight bold
  ttk::style configure TLabelframe.Label -background $bgcolor -font fontBold

# control o,q
  bind . <Control-o> {openFile}
  bind . <Control-d> {openMultiFile}
  bind . <Key-F4>    {openMultiFile 0}
  bind . <Control-q> {exit}

  bind . <Key-F1> {
    set localName [getFirstFile]
    if {$localName != ""} {
      set localNameList [list $localName]
      genExcel
    }
  }

  bind . <Key-F2> {set lastXLS [openXLS $lastXLS 1]}
  if {$lastXLS1 != ""} {bind . <Key-F3> {set lastXLS1 [openXLS $lastXLS1 1]}}

  bind . <MouseWheel> {[$fout.text component text] yview scroll [expr {-%D/30}] units}
  bind . <Up>     {[$fout.text component text] yview scroll -1 units}
  bind . <Down>   {[$fout.text component text] yview scroll  1 units}
  bind . <Left>   {[$fout.text component text] xview scroll -1 units}
  bind . <Right>  {[$fout.text component text] xview scroll  1 units}
  bind . <Prior>  {[$fout.text component text] yview scroll -30 units}
  bind . <Next>   {[$fout.text component text] yview scroll  30 units}
  bind . <Home>   {[$fout.text component text] yview scroll -100000 units}
  bind . <End>    {[$fout.text component text] yview scroll  100000 units}
}

#-------------------------------------------------------------------------------
# buttons and progress bar
proc guiButtons {} {
  global buttons wdir nline nprogfile ftrans tcl_platform mytemp nistVersion

# for some unknown reason, can't put the button in here because it causes an error

  #set ftrans [frame .ftrans1 -bd 2 -background "#E0DFE3"]
  #set buttons(genExcel) [ttk::button $ftrans.generate1 -text "Generate Spreadsheet" -padding 4 \
  #  -state disabled -command {
  #    saveState
  #    if {![info exists localNameList]} {
  #    set localName [getFirstFile]
  #    if {$localName != ""} {
  #      set localNameList [list $localName]
  #      genExcel
  #    }
  #    } elseif {[llength $localNameList] == 1} {
  #    genExcel
  #    } else {
  #    openMultiFile 2
  #    }
  #  }]
  #pack $ftrans.generate1 -side left -padx 10

  if {$nistVersion} {
    catch {
      set l3 [label $ftrans.l3 -relief flat -bd 0]
      $l3 config -image [image create photo -file [file join $wdir images nist.gif]]
      pack $l3 -side right -padx 10
      bind $l3 <ButtonRelease-1> {displayURL https://www.nist.gov}
      tooltip::tooltip $l3 "Click here to learn more about NIST"
    }
  }

  pack $ftrans -side top -padx 10 -pady 10 -fill x

  if {$tcl_platform(osVersion) < 6.0} {
    set fbar [frame .fbar -bd 2 -background "#E0DFE3"]
  } else {
    set fbar [frame .fbar -bd 2 -background "#F0F0F0"]
  }
  set nline 0
  set buttons(pgb) [ttk::progressbar $fbar.pgb -mode determinate -variable nline]
  pack $fbar.pgb -side top -padx 10 -fill x

  set nprogfile 0
  set buttons(pgb1) [ttk::progressbar $fbar.pgb1 -mode determinate -variable nprogfile]
  pack forget $buttons(pgb1)
  #pack $fbar.pgb1 -side top -padx 10 -pady {5 0} -expand true -fill x
  pack $fbar -side bottom -padx 10 -pady {0 10} -fill x

# icon bitmap
  if {$nistVersion} {
    catch {file copy -force [file join $wdir images NIST.ico] $mytemp}
    catch {wm iconbitmap . -default [file join $mytemp NIST.ico]}
  }
}

#-------------------------------------------------------------------------------
# status tab
proc guiStatusTab {} {
  global nb wout fout outputWin statusFont tcl_platform

  set wout [ttk::panedwindow $nb.status -orient horizontal]
  $nb add $wout -text " Status " -padding 2
  set fout [frame $wout.fout -bd 2 -relief sunken -background "#E0DFE3"]

  set outputWin [iwidgets::messagebox $fout.text -maxlines 500000 \
    -hscrollmode dynamic -vscrollmode dynamic -background white]
  pack $fout.text -side top -fill both -expand true
  pack $fout -side top -fill both -expand true

  $outputWin type add black -foreground black -background white
  $outputWin type add red -foreground "#bb0000" -background white
  $outputWin type add green -foreground "#009900" -background white
  $outputWin type add magenta -foreground "#990099" -background white
  $outputWin type add cyan -foreground "#00dddd" -background white
  $outputWin type add blue -foreground blue -background white
  $outputWin type add error -foreground black -background "#ffff99"
  $outputWin type add ifc -foreground black -background "#99ffff"
  $outputWin type add syntax -foreground black -background "#ff9999"

  if {[info exists statusFont]} {
    regsub -all 110 $statusFont 120 statusFont
    regsub -all 130 $statusFont 120 statusFont
    regsub -all 150 $statusFont 140 statusFont
  }

  if {$tcl_platform(osVersion) >= 6.0} {
    if {![info exists statusFont]} {set statusFont [$outputWin type cget black -font]}
    if {[string first "Courier" $statusFont] != -1} {
      regsub "Courier" $statusFont "Consolas" statusFont
      regsub "120" $statusFont "140" statusFont
      saveState
    }
  }

  if {[info exists statusFont]} {
    foreach typ {black red green magenta cyan blue error ifc syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }

  bind . <Key-F6> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error ifc syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }
  bind . <Control-KeyPress-=> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error ifc syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }

  bind . <Key-F5> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error ifc syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }
  bind . <Control-KeyPress--> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error ifc syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }
}

#-------------------------------------------------------------------------------
# file menu
proc guiFileMenu {} {
  global File openFileList lastXLS lastXLS1

  $File add command -label "Open IFC File(s)..." -accelerator "Ctrl+O" -command openFile
  $File add command -label "Open Multiple IFC Files in a Directory..." -accelerator "Ctrl+D, F4" -command {openMultiFile}
  set newFileList {}
  foreach fo $openFileList {if {[file exists $fo]} {lappend newFileList $fo}}
  set openFileList $newFileList

  set llen [llength $openFileList]
  $File add separator
  if {$llen > 0} {
    for {set fi 0} {$fi < $llen} {incr fi} {
      set fo [lindex $openFileList $fi]
      if {$fi != 0} {
        $File add command -label [truncFileName [file nativename $fo] 1] -command [list openFile $fo]
      } else {
        $File add command -label [truncFileName [file nativename $fo] 1] -command [list openFile $fo] -accelerator "F1"
      }
    }
  }
  $File add separator
  $File add command -label "Open Last Spreadsheet" -accelerator "F2" -command {set lastXLS [openXLS $lastXLS 1]}
  if {$lastXLS1 != ""} {
    $File add command -label "Open Last Multiple File Summary Spreadsheet" -accelerator "F3" -command {set lastXLS1 [openXLS $lastXLS1 1]}
  }
  $File add command -label "Exit" -accelerator "Ctrl+Q" -command exit
}

#-------------------------------------------------------------------------------
# overview
proc helpOverview {} {

outputMsg "\nOverview -------------------------------------------------------------------" blue
outputMsg "The IFC File Analyzer reads an IFC file and generates an Excel spreadsheet or CSV
files.  One worksheet or CSV file is generated for each entity type in the IFC
file.  Each worksheet or CSV file lists every entity instance and its attributes.
The types of entities that are Processed can be selected in the Options tab.
Other options are available that add to or modify the information written to the
spreadsheet or CSV files.

IFC2x3 and IFC4 are supported, however, IFC4.0.n addendums and IFC4.n versions 
are not supported.  If the IFC file contains IFC4.0.n entities, those entities
will not be processed and will not be listed as Entity types not processed on the
File Summary worksheet.  See Websites > IFC Documentation

For spreadsheets, a Summary worksheet shows the Count of each entity.  Links on
the Summary and entity worksheets can be used to navigate to other worksheets and
to access IFC entity documentation.

Spreadsheets or CSV files can be selected in the Options tab.  CSV files are
automatically generated if Excel is not installed.

To generate a spreadsheet or CSV files, select an IFC file from the File menu above
and click the Generate button below.  Existing spreadsheet or CSV files are always
overwritten.

Multiple IFC files can be selected or an entire directory structure of IFC files
can also be processed from the File menu.  If multiple IFC files are translated,
then a separate File Summary spreadsheet is also generated.  This is useful to
compare entity usage between different IFC files.

Tooltip help is available for the selections in the tabs.  Hold the mouse over
text in the tabs until a tooltip appears.

Use F6 and F5 to change the font size.  Right-click to save the text."

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# crash recovery
proc helpCrash {} {

set num ""

outputMsg "\nCrash Recovery -------------------------------------------------------------" blue
outputMsg "Sometimes the IFC File Analyzer crashes after an IFC file has been successfully
opened and the processing of entities has started.  Popup dialogs might appear
that say \"Runtime Error!\" or \"ActiveState Basekit has stopped working\" or
\"Fatal Error in Wish - unable to alloc 123456 bytes\".

A crash is most likely due to syntax errors in the IFC file or sometimes due to
limitations of the toolkit used to read IFC files.  To see which type of entity
caused the error, check the Status tab to see which type of entity was last
processed.  A crash can also be caused by insufficient memory to process a very
large IFC file.

Workarounds for these problems:

1 - Processing of the type of entity that caused the error can be deselected in
the Options tab under Process.  However, this will prevent processing of other
entities that do not cause a crash.  Deselecting entity types might also help
with large IFC files.

2 - Run the command-line version 'IFC-File-Analyzer-CL.exe' in a command prompt
window.  The output from reading the IFC file might show error and warning
messages that might have caused the software to crash.  Those messages will be
between the 'Begin ST-Developer output' and 'End ST-Developer output' messages."

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global Help row_limit ifcsvrdir nistVersion

  #$Help add command -label "What's New" -command {whatsNew}
  #$Help add command -label "Check for Update" -command {
  #  set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
  #  outputMsg "The last check for an update was $lastupgrade days ago." red
  #  set os "$tcl_platform(os) $tcl_platform(osVersion)"
  #  regsub -all " " $os "" os
  #  regsub "WindowsNT" $os "" os
  #  if {$pf64 != ""} {append os ".64"}
  #  set url "https://concrete.nist.gov/cgi-bin/ctv/ifa_upgrade.cgi?version=[getVersion]&auto=-$lastupgrade&os=$os"
  #  if {[info exists yrexcel]} {if {$yrexcel != ""} {append url "&yr=[expr {$yrexcel-2000}]"}}
  #  displayURL $url
  #  set upgrade [clock seconds]
  #  saveState
  #}
  #$Help add separator

$Help add command -label "Overview" -command {helpOverview}

# options help

$Help add command -label "Options" -command {
outputMsg "\nOptions --------------------------------------------------------------------" blue
outputMsg "*Process: Select which types of entities are processed.  The tooltip help lists
all the entities associated with that type.  Selectively process only the entities
relevant to your analysis."

outputMsg "\n*Inverse Relationships: For Building Elements, Building Services, and Structural
Analysis entities, some Inverse Relationships are displayed on the worksheets.
The Inverse values are displayed in additional columns of entity worksheets that
are highlighted in light blue."

outputMsg "\n*Expand: The attributes that IfcLocalPlacement, IfcAxis2Placement, or structural
analysis entities refer to will be displayed inline with the entity.  For example,
IfcLocalPlacement refers to an IfcAxis2Placement3D and an optional relative
placement. Those values would be included in addition to the IfcLocalPlacement.
IfcAxis2Placement expands into an IfcCartesianPoint and IfcDirection.  The columns
with the expanded values are color coded.  The expanded columns can be collapsed
on a worksheet."

outputMsg "\n*Output Format: Generate Excel spreadsheets or CSV files.  If Excel is not
installed, CSV files are automatically generated.  Some options are not available
with CSV files."

outputMsg "\n*Count Duplicates: Entities with identical attribute values will be counted and
not duplicated on a worksheet.  This applies to a limited set of entities."

outputMsg "\n*Table: Generate tables for each spreadsheet to facilitate sorting and filtering
(Spreadsheet tab)."

outputMsg "\n*Number Format: Option to not round real numbers."

outputMsg "\n*Maximum Rows: The maximum number of rows for any worksheet can be set lower than
the normal limits for Excel.  This is useful for very large IFC files at the
expense of not processing some entities."

  .tnb select .tnb.status
  update idletasks
}

# count duplicates help
$Help add command -label "Count Duplicates" -command {
outputMsg "\nCount Duplicates -----------------------------------------------------------" blue
outputMsg "When using the Count Duplicates option in the Options tab, entities with identical
attribute values will be counted and notduplicated on a worksheet.  The resulting
entity worksheets can be shorter.  However, counting duplicates in very large IFC
files can be slow and the program might run out of memory.

Some entity attributes might be ignored to check for duplicates.  Duplicate
entities are counted only if there are more than 100 instances of an entity.  The
entity count is displayed in the last column of the worksheet.  The  entity ID
displayed is of the first of the duplicate entities.

If there are no duplicates for an entity type being counted and there are a lot
(> 50000) of that entity type, then the processing can be slow.  This is most
common with Geometry entities.

The list of IFC entities that are counted is displayed in the  Count Duplicates
tooltip on the Options tab."

  .tnb select .tnb.status
  update idletasks
}

# number format help

$Help add command -label "Number Format" -command {
outputMsg "\nNumber Format --------------------------------------------------------------" blue
outputMsg "By default Excel rounds real numbers if there are more than 11 characters in the
number string.

For example, the number 0.12499999999999997 in the IFC file will be displayed as
0.125.  However, double clicking in a cell with a rounded number will show all of
the digits.

This option will display most real numbers exactly as they appear in the IFC file.
This applies only to single real numbers.  Lists of real numbers, such as
cartesian point coordinates, are always displayed exactly as they appear in the
IFC file.

Rounding real numbers might affect how Count Duplicates appears.  If both
0.12499999999999997 and 0.12499999999999993 are rounded to 0.125 they will appear
as two separate values of 0.125 when it would seem that they are identical each
other."

  .tnb select .tnb.status
  update idletasks
}

# display files help

$Help add command -label "Open IFC Files" -command {
outputMsg "\nOpen IFC Files ---------------------------------------------------------" blue
outputMsg "This option is a convenient way to open an IFC file in other applications.  The
pull-down menu will contain applications that can open an IFC file such as IFC
viewers, browsers, and conformance checkers.  If applications are installed in
their default location, then they will appear in the pull-down menu.

The 'Indent IFC File (for debugging)' option rearranges and indents the entities
to show the hierarchy of information in an IFC file.  The 'indented' file is
written to the same directory as the IFC file or to the same user-defined
directory specified in the Spreadsheet tab.  It is useful for debugging IFC files
but is not recommended for large IFC files.

The 'Default IFC Viewer' option will open the IFC file in whatever application is
associated with IFC files.

A text editor will always appear in the menu."

  .tnb select .tnb.status
  update idletasks
}

# multiple files help

$Help add command -label "Multiple IFC Files" -command {
outputMsg "\nMultiple IFC Files --------------------------------------------------------" blue
outputMsg "Multiple IFC files can be selected in the Open File(s) dialog by holding down the
control or shift key when selecting files or an entire directory of IFC files can
be selected with 'Open Multiple IFC Files in a Directory'. Files in subdirectories
of the selected directory can also be processed.

When processing multiple IFC files, a File Summary spreadsheet is generated in
addition to individual spreadsheets for each file.  The File Summary spreadsheet
shows the entity count and totals for all IFC files.  The File Summary spreadsheet
also links to the individual spreadsheets and the IFC file.

If only the File Summary spreadsheet is needed, it can be generated faster by
turning off Processing of most of the entity types and options in the Options tab."

  .tnb select .tnb.status
  update idletasks
}

# large files help

$Help add command -label "Large IFC Files" -command {
outputMsg "\nLarge IFC Files -----------------------------------------------------------" blue
outputMsg "To reduce the amount of time to process large IFC files and to reduce the size of
the resulting spreadsheet, several options are available:
- In the Process section, deselect entity types for which there are usually a lot
    of, such as Geometry and Properties
- In the Options tab, deselect some of the other options
- In the Spreadsheet tab, select the Maximum Rows for any worksheet"

  .tnb select .tnb.status
  update idletasks
}

#$Help add separator
$Help add command -label "Crash Recovery" -command {helpCrash}

#$Help add command -label "Errors" -command {
#outputMsg "\nErrors ---------------------------------------------------------------------" blue
#outputMsg "1 - If sufficient memory is not available to process a very large IFC file, then
#the IFC File Analyzer will stop with an error message that might say
#\"Fatal Error in Wish - unable to alloc 123456 bytes\".  Try processing the IFC
#file on a computer with more memory or deselect some categories of entities to
#Process in the Options tab.
#
#2 - After stopping the program when a large IFC file has been processed, sometimes
#the IFC-File-Analyzer.exe and/or EXCEL.EXE processes will still be running.  The
#Windows Task Manager can be used to kill those processes.  When starting the IFC
#File Analyzer, if there is another instance of IFC-File-Analyzer.exe already
#running, then you will be prompted to close the other instances."
#
#  .tnb select .tnb.status
#  update idletasks
#}

$Help add separator
if {$nistVersion} {
  $Help add command -label "Disclaimers" -command {displayDisclaimer}
  $Help add command -label "NIST Disclaimer" -command {displayURL https://www.nist.gov/disclaimer}
}
$Help add command -label "About" -command {
outputMsg "\nIFC File Analyzer ---------------------------------------------------------" blue
outputMsg "Version:  [getVersion]"
if {$nistVersion} {
outputMsg "Contact:  Robert Lipman, robert.lipman@nist.gov

The IFC File Analyzer was developed at NIST in the former Computer Integrated Building
Processes Group in the Building and Fire Research Laboratory.  The software was first
released in 2008 and development ended in 2014.  Minor updates have been made since 2014.

See Help > Disclaimer and NIST Disclaimer

Credits
- Generating spreadsheets:       Microsoft Excel (https://products.office.com/excel)
- Reading and parsing IFC files: IFCsvr (https://groups.yahoo.com/neo/groups/ifcsvr-users/info)
                                 License agreement C:\\Program Files (x86)\\IFCsvrR300\\doc
                                 IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved"

# debug
  if {$row_limit == 100003 || $env(USERDOMAIN) == "NIST"} {
    outputMsg " "
    outputMsg "Environment variables" red
    foreach id [lsort [array names env]] {
      #outputMsg " $id   $env($id)" green
      foreach id1 [list HOME Program System USER TEMP TMP ROSE EDM] {
        if {[string first $id1 $id] == 0} {outputMsg " $id   $env($id)"; break}
      }
    }
    outputMsg "Registry values" red
    catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
    catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
    catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
    catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}
    catch {outputMsg " Browser   [registry get {HKEY_CURRENT_USER\Software\Classes\http\shell\open\command} {}]"}
    outputMsg "SFA variables" red
      catch {outputMsg " Drive $drive"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp  ([file exists $mytemp])"}
      catch {outputMsg " ifcsvrdir  [file nativename $ifcsvrdir]"}
      if {[info exists virtualDir]} {outputMsg " virtualDir  $virtualDir"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      outputMsg "Other variables" red
      outputMsg " Tcl [info patchlevel]"
      outputMsg " twapi [package versions twapi]"
      outputMsg " $tcl_platform(os) $tcl_platform(osVersion)"
  }

  .tnb select .tnb.status
  update idletasks
}
}
}

#-------------------------------------------------------------------------------
# user-defined list of entities
proc guiUserDefinedEntities {} {
  global fopta opt cb buttons fileDir userEntityFile

  set fopta6 [frame $fopta.6 -bd 0]
  foreach item {{" User-Defined List: " opt(PR_USER)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta6.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }

  set buttons(userentity) [ttk::entry $fopta6.entry -width 50 -textvariable userEntityFile]
  pack $fopta6.entry -side left -anchor w

  set buttons(userentityopen) [ttk::button $fopta6.$cb -text " Browse " -command {
    set typelist {{"All Files" {*}}}
    set uef [tk_getOpenFile -title "Select File of IFC Entities" -filetypes $typelist -initialdir $fileDir]
    if {$uef != "" && [file isfile $uef]} {
      set userEntityFile [file nativename $uef]
      outputMsg "User-defined IFC list: [truncFileName $userEntityFile]" blue
      set fileent [open $userEntityFile r]
      set userentlist {}
      while {[gets $fileent line] != -1} {
        set line [split [string trim $line] " "]
        foreach ent1 $line {lappend userentlist $ent1}
      }
      close $fileent
      set llist [llength $userentlist]
      if {$llist > 0} {
        outputMsg " ($llist) $userentlist"
      } else {
        outputMsg "File does not contain any IFC entity names" red
        set opt(PR_USER) 0
        checkValues
      }
      .tnb select .tnb.status
    }
    checkValues
  }]
  pack $fopta6.$cb -side left -anchor w -padx 10
  incr cb
  foreach item {optPR_USER userentity userentityopen} {
    catch {tooltip::tooltip $buttons($item) "A User-Defined List is a text file with one IFC entity name per line.\nThis allows for more control to process only the required entity types.\nIt is also useful when processing large files that might crash the software."}
  }
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# display result
proc guiDisplayResult {} {
  global buttons cb fopt appNames dispCmds appName dispApps foptf
  global edmWriteToFile edmWhereRules eeWriteToFile

  set foptf [ttk::labelframe $fopt.f -text " Open IFC File in "]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 40]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}
  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]
    catch {
      if {[string first "EDM Model Checker" $appName] == 0} {
        pack $buttons(edmWriteToFile)  -side left -anchor w -padx 5
        pack $buttons(edmWhereRules) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(edmWriteToFile)
        pack forget $buttons(edmWhereRules)
      }
    }
    catch {
      if {[string first "Conformance Checker" $appName] != -1} {
        pack $buttons(eeWriteToFile) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(eeWriteToFile)
      }
    }

# set the app command
    foreach cmd $dispCmds {
      if {$appName == $dispApps($cmd)} {
        set dispCmd $cmd
      }
    }

# put the app name at the top of the list
    for {set i 0} {$i < [llength $dispCmds]} {incr i} {
      if {$dispCmd == [lindex $dispCmds $i]} {
        set dispCmds [lreplace $dispCmds $i $i]
        set dispCmds [linsert $dispCmds 0 $dispCmd]
      }
    }
    set appNames {}
    foreach cmd $dispCmds {
      if {[info exists dispApps($cmd)]} {lappend appNames $dispApps($cmd)}
    }
    $foptf.spinbox configure -values $appNames
  }

  set buttons(appDisplay) [ttk::button $foptf.$cb -text " Open " -state disabled -command {
    displayResult
    saveState
  }]
  pack $foptf.$cb -side left -anchor w -padx {10 0} -pady {0 3}
  incr cb

  foreach item $appNames {
    if {[string first "EDM Model Checker" $item] == 0} {
      foreach item {{" Write results to a file" edmWriteToFile}} {
        regsub -all {[\(\)]} [lindex $item 1] "" idx
        set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
          -variable [lindex $item 1] -command {checkValues}]
        pack forget $buttons($idx)
        #pack $buttons($idx) -side left -anchor w -padx 5
        incr cb
      }
    }
  }
  if {[lsearch -glob $appNames "*Conformance Checker*"] != -1} {
    foreach item {{" Write results to a file" eeWriteToFile}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
        -variable [lindex $item 1] -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

  catch {tooltip::tooltip $foptf "This option is a convenient way to open an IFC file in other applications.\nThe pull-down menu will contain applications that can open an IFC file\nsuch as IFC viewers, browsers, and conformance checkers.  If applications\nare installed in their default location, then they will appear in the\npull-down menu.\n\nThe 'Indent IFC File (for debugging)' option rearranges and indents the\nentities to show the hierarchy of information in an IFC file.  The 'indented'\nfile is written to the same directory as the IFC file or to the same\nuser-defined directory specified in the Spreadsheet tab.\n\nThe 'Default IFC Viewer' option will open the IFC file in whatever\napplication is associated with IFC files."}
  pack $foptf -side top -anchor w -pady {5 2} -padx 10 -fill both

# output format hiding here
  set foptk [ttk::labelframe $fopt.k -text " Output Format "]
  foreach item {{" Open Output files after they have been generated" opt(XL_OPEN)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side bottom -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  foreach item {{" Excel" Excel} {" CSV" CSV}} {
    pack [ttk::radiobutton $foptk.$cb -variable opt(XLSCSV) -text [lindex $item 0] -value [lindex $item 1] -command {checkValues}] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptk -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $foptk "Microsoft Excel is required to generate spreadsheets.  CSV files will be generated if Excel is not installed.\n\nOne CSV file is generated for each entity type.  Some of the options are not available with CSV files."}
}

#-------------------------------------------------------------------------------
# count duplicates
proc guiDuplicates {} {
  global buttons cb foptbf opt countent

  set foptb1 [ttk::labelframe $foptbf.1 -text " Count Duplicates "]
  foreach item {{" Count Duplicate identical entities" opt(COUNT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptb1.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptb1 -side top -anchor w -pady {5 2} -padx 10 -fill both

  set ttmsg ""

  if {[info exists countent(IFC)]} {
    set ttlen 0
    set lchar ""
    foreach item [lsort $countent(IFC)] {
      if {[string range $item 0 3] != $lchar && $lchar != ""} {
        if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
        set ttlen 0
      }
      append ttmsg "$item   "
      incr ttlen [string length $item]
      if {$ttlen > 150} {
        if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
        set ttlen 0
        set ok 0
      }
      set lchar [string range $item 0 3]
    }
  }

  set tmsg "Entities with identical attribute values will be counted and not duplicated on a worksheet.\nThe resulting entity worksheets can be shorter.  However, counting duplicates in very large\nIFC files can be slow and the program might run out of memory."
  append tmsg "\n\nSee Help > Count Duplicates"
  append tmsg "\n\nSpecific values for IFC Properties might be abbreviated with (Number) and (Integer) when counted."
  append tmsg "\n\nThe following IFC entities have Duplicates Counted:\n\n$ttmsg"
  catch {tooltip::tooltip $foptb1 $tmsg}
}

#-------------------------------------------------------------------------------
# sort entities
#proc guiSort {} {
#  global buttons cb foptbf opt
#
#  set foptb2 [ttk::labelframe $foptbf.2 -text " Sort "]
#  foreach item {{" Sort entities by Name attribute" opt(SORT)}} {
#    regsub -all {[\(\)]} [lindex $item 1] "" idx
#    set buttons($idx) [ttk::checkbutton $foptb2.$cb -text [lindex $item 0] \
#      -variable [lindex $item 1] -command {checkValues}]
#    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
#    incr cb
#  }
#  pack $foptb2 -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
#  catch {tooltip::tooltip $foptb2 "Sort the entity worksheet rows by the Name attribute and some other similar text attributes.\n\nFor large IFC files, this option can slow down the processing of the file."}
#}

#-------------------------------------------------------------------------------
# max rows

#proc guiMaxRows {} {
#  global buttons cb fxls row_limit
#
#  set fxlse [ttk::labelframe $fxls.e -text " Maximum Rows for any worksheet"]
#  set rlimit {{" No limit" 10000000} {" 100" 103} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003} {" 100000" 100003}}
#  foreach item $rlimit {
#    pack [ttk::radiobutton $fxlse.$cb -variable row_limit -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
#    incr cb
#  }
#  pack $fxlse -side top -anchor w -pady 5 -padx 10 -fill both
#  set msg "This option will limit the number of rows (entities) written to any one worksheet.\nWithout setting a maximum, the row maximums are:\n\nExcel 2007 and later:  1,048,576 rows\nExcel 2003 and earlier:    65,536 rows\n\nFor large IFC files, setting a low maximum can speed up processing\nat the expense of not processing all of the entities.  This is useful when processing\nGeometry entities."
#  catch {tooltip::tooltip $fxlse $msg}
#}

#-------------------------------------------------------------------------------
# inverse relationships
proc guiInverse {} {
  global buttons cb fopt inverses env opt

  set foptc [ttk::labelframe $fopt.3 -text " Inverse Relationships "]
  set txt " Show some Inverse Relationships for:\n Building Elements, HVAC, Electrical, Building Services, and Structural Analysis"

  regsub -all {[\(\)]} opt(INVERSE) "" idx
  set buttons($idx) [ttk::checkbutton $foptc.$cb -text $txt \
    -variable opt(INVERSE) -command {
      checkValues
      if {$opt(INVERSE)} {set opt(PR_RELA) 1}
    }]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

  #if {$env(USERDOMAIN) == "NIST"} {
  #  regsub -all {[\(\)]} opt(DEBUGINV) "" idx
  #  set buttons($idx) [ttk::checkbutton $foptc.$cb -text "Debug" -variable opt(DEBUGINV)]
  #  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  #  incr cb
  #}

  pack $foptc -side top -anchor w -pady {5 2} -padx 10 -fill both
  set ttmsg "Inverse Relationships\n"
  foreach item [lsort $inverses] {
    set ok 1
    if {$ok} {
      regsub " " $item "  (" item
      append item ")"
      append ttmsg \n$item
    }
  }
  append ttmsg "\n\nInverse Relationships are shown on the entity worksheets.  The Inverse values are\nshown in additional columns of the worksheets that are highlighted in light blue."
  catch {tooltip::tooltip $foptc $ttmsg}
}

#-------------------------------------------------------------------------------
# spreadsheet tab
proc guiSpreadsheet {} {
  global buttons cb env extXLS fileDir fxls maxfiles mydocs nb opt row_limit
  global userWriteDir userXLSFile writeDir writeDirType yrexcel

  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " Spreadsheet " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

  set fxlsz [ttk::labelframe $fxls.z -text " Tables "]
  foreach item {{" Generate Tables for Sorting and Filtering" opt(SORT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsz.$cb -text [lindex $item 0] \
      -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsz -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "Worksheets can be sorted by column values."
  catch {tooltip::tooltip $fxlsz $msg}

  set fxlsa [ttk::labelframe $fxls.a -text " Number Format "]
  foreach item {{" Do not round Real Numbers" opt(XL_FPREC)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsa.$cb -text [lindex $item 0] \
      -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsa -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "See Help > Number Format"
  catch {tooltip::tooltip $fxlsa $msg}

  set fxlsb [ttk::labelframe $fxls.b -text " Maximum Rows for any worksheet"]
  set rlimit {{" 100" 103} {" 500" 503} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003}}
  if {$yrexcel == "" || $yrexcel >= 2007} {
    lappend rlimit {" 100000" 100003}
    lappend rlimit {" Maximum" 1048576}
  } else {
    lappend rlimit {" Maximum" 65536}
  }
  foreach item $rlimit {
    pack [ttk::radiobutton $fxlsb.$cb -variable row_limit -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb -side top -anchor w -pady 5 -padx 10 -fill both
  set msg "This option will limit the number of rows (entities) written to any one worksheet.\nThe Maximum rows depends on the version of Excel.\nFor large IFC files, setting a low maximum can speed up processing at the expense\nof not processing all of the entities.  This is useful when processing Geometry entities."
  append msg "\n\nIf the maximum number of rows is exceeded, then the counts on the summary\nworksheet for Name, Description, etc. might not be correct."
  catch {tooltip::tooltip $fxlsb $msg}

  set fxlsc [ttk::labelframe $fxls.c -text " Other "]
  foreach item {{" On File Summary worksheet, create links to IFC files and spreadsheets (see File > Open Multiple ...)" opt(XL_LINK1)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsc.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsc -side top -anchor w -pady {5 2} -padx 10 -fill both

  set fxlsd [ttk::labelframe $fxls.d -text " Write Spreadsheet to "]
  set buttons(fileDir) [ttk::radiobutton $fxlsd.$cb \
    -text " Same directory as the IFC file" \
    -variable writeDirType -value 0 -command checkValues]
  pack $fxlsd.$cb -side top -anchor w -padx 5 -pady 2
  incr cb

  set fxls1 [frame $fxlsd.1]
  ttk::radiobutton $fxls1.$cb -text " User-defined directory:  " \
    -variable writeDirType -value 2 -command {
    checkValues
    if {[file exists $userWriteDir] && [file isdirectory $userWriteDir]} {
      set writeDir $userWriteDir
    } else {
      set userWriteDir $mydocs
      tk_messageBox -type ok -icon error -title "Invalid Directory" \
        -message "The user-defined directory to write the Spreadsheet to is not valid.\nIt has been set to $userWriteDir"
    }
    focus $buttons(userdir)
  }
  pack $fxls1.$cb -side left -anchor w -padx {5 0}
  catch {tooltip::tooltip $fxls1.$cb "This option can be used when the directory containing the IFC file is\nprotected (read-only) and the Spreadsheet cannot be written to it."}
  incr cb

  set buttons(userentry) [ttk::entry $fxls1.entry -width 38 -textvariable userWriteDir]
  pack $fxls1.entry -side left -anchor w -pady 2
  set buttons(userdir) [ttk::button $fxls1.button -text " Browse " -command {
    set uwd [tk_chooseDirectory -title "Select directory"]
    if {[file isdirectory $uwd]} {
      set userWriteDir $uwd
      set writeDir $userWriteDir
    }
  }]
  pack $fxls1.button -side left -anchor w -padx 10 -pady 2
  pack $fxls1 -side top -anchor w

  set fxls2 [frame $fxlsd.2]
  ttk::radiobutton $fxls2.$cb -text " User-defined file name:  " \
    -variable writeDirType -value 1 -command {
    checkValues
    focus $buttons(userfile)
  }
  pack $fxls2.$cb -side left -anchor w -padx {5 0}
  incr cb

  set buttons(userentry1) [ttk::entry $fxls2.entry -width 35 -textvariable userXLSFile]
  pack $fxls2.entry -side left -anchor w -pady 2
  set buttons(userfile) [ttk::button $fxls2.button -text " Browse " -command {
    if {$extXLS == "xls"} {
      set typelist {{"Excel Files" {".xls"}}}
    } else {
      set typelist {{"Excel Files" {".xlsx"}}}
    }
    set uxf [tk_getSaveFile -title "Save Spreadsheet to" -filetypes $typelist -initialdir $fileDir -defaultextension ".$extXLS"]
    if {$uxf != ""} {
      set userXLSFile $uxf
    } else {
      errorMsg "No file selected"
    }
  }]
  pack $fxls2.button -side left -anchor w -padx 10 -pady 2
  pack $fxls2 -side top -anchor w

  pack $fxlsd -side top -anchor w -pady {5 2} -padx 10 -fill both

  #if {$env(USERDOMAIN) == "NIST"} {
  #  set fxlsz [ttk::labelframe $fxls.z -text " Limit maximum of multiple files "]
  #  foreach item {{" 1" 1} {" 5" 5} {" 10" 10} {" 20" 20} {" 40" 40} {" 100" 100} {" 1000" 1000} {" All" 100000}} {
  #    pack [ttk::radiobutton $fxlsz.$cb -variable maxfiles -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
  #    incr cb
  #  }
  #  pack $fxlsz -side top -anchor w -pady 5 -padx 10 -fill both
  #}
  pack $fxls -side top -fill both -expand true -anchor nw
}

#-------------------------------------------------------------------------------
proc displayDisclaimer {} {

set txt "This software was developed at the National Institute of Standards and Technology by employees of the Federal Government in the course of their official duties. Pursuant to Title 17 Section 105 of the United States Code this software is not subject to copyright protection and is in the public domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for its use by other parties, and makes no guarantees, expressed or implied, about its quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links in this software, NIST does not necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses Microsoft Excel and IFCsvr that are covered by their own Software License Agreements.  The IFCsvr agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc

See Help > NIST Disclaimer and Help > About"

  tk_messageBox -type ok -icon info -title "Disclaimers" -message $txt
}

#-------------------------------------------------------------------------------
# shortcuts
proc setShortcuts {} {
  global mydesk mymenu mytemp tcl_platform

  set progname [info nameofexecutable]
  if {[string first "AppData/Local/Temp" $progname] != -1 || [string first ".zip" $progname] != -1} {
    errorMsg "For the IFC File Analyzer to run properly, it is recommended that you first\n extract all of the files from the ZIP file and run the extracted executable."
    return
  }

  if {[info exists mydesk] || [info exists mymenu]} {
    set ok 1
    set app IFC_Excel
    foreach scut [list "Shortcut to $app.exe.lnk" "$app.exe.lnk" "$app.lnk"] {
      catch {if {[file exists [file join $mydesk $scut]]} {set ok 0; break}}
    }
    if {[file exists [file join $mydesk [file tail [info nameofexecutable]]]]} {set ok 0}

    if {$ok} {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" \
        -message "Do you want to create or overwrite a shortcut to the IFC File Analyzer (v[getVersion]) in the Start Menu and an icon on the Desktop?"]
    } else {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" \
        -message "Do you want to create or overwrite a shortcut to the IFC File Analyzer (v[getVersion]) in the Start Menu"]
    }
    if {$choice == "yes"} {
      outputMsg " "
      catch {
        if {[info exists mymenu]} {
          if {[file exists [file join $mymenu "IFC File Analyzer.lnk"]]} {outputMsg "Existing Start Menu shortcut will be overwritten" red}
          if {$tcl_platform(osVersion) >= 6.2} {
            create_shortcut [file join $mymenu "IFC File Analyzer.lnk"] Description "IFC File Analyzer" TargetPath [info nameofexecutable] IconLocation [info nameofexecutable]
          } else {
            create_shortcut [file join $mymenu "IFC File Analyzer.lnk"] Description "IFC File Analyzer" TargetPath [info nameofexecutable] IconLocation [file join $mytemp NIST.ico]
          }
          outputMsg " Shortcut created in Start Menu to [truncFileName [file nativename [info nameofexecutable]]]"
        }
      }

      if {$ok} {
        catch {
          if {[info exists mydesk]} {
            if {[file exists [file join $mydesk "IFC File Analyzer.lnk"]]} {outputMsg "Existing Desktop shortcut will be overwritten" red}
            if {$tcl_platform(osVersion) >= 6.2} {
              create_shortcut [file join $mydesk "IFC File Analyzer.lnk"] Description "IFC File Analyzer" TargetPath [info nameofexecutable] IconLocation [info nameofexecutable]
            } else {
              create_shortcut [file join $mydesk "IFC File Analyzer.lnk"] Description "IFC File Analyzer" TargetPath [info nameofexecutable] IconLocation [file join $mytemp NIST.ico]
            }
            outputMsg " Shortcut created on Desktop to [truncFileName [file nativename [info nameofexecutable]]]"
          }
        }
      }
    }
  }
}

#-------------------------------------------------------------------------------
# set home, docs, desktop, menu directories
proc setHomeDir {} {
  global env tcl_platform drive myhome mydocs mydesk mymenu mytemp

  set drive "C:/"
  if {[info exists env(SystemDrive)]} {
    set drive $env(SystemDrive)
    append drive "/"
  }
  set myhome $drive

# set based on USERPROFILE and registry entries
  if {[info exists env(USERPROFILE)]} {
    set myhome $env(USERPROFILE)
    catch {
      set reg_personal [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]
      if {[string first "%USERPROFILE%" $reg_personal] == 0} {regsub "%USERPROFILE%" $reg_personal $env(USERPROFILE) mydocs}
    }
    catch {
      set reg_desktop  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]
      if {[string first "%USERPROFILE%" $reg_desktop]  == 0} {regsub "%USERPROFILE%" $reg_desktop  $env(USERPROFILE) mydesk}
    }
    catch {
      set reg_menu [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_menu $env(USERPROFILE) mymenu}
    }
    if {$tcl_platform(osVersion) < 6.0} {
      catch {
        set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local Settings}]
        if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_temp $env(USERPROFILE) mytemp}
        set mytemp [file join $mytemp Temp]
      }
    } else {
      catch {
        set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]
        if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_temp $env(USERPROFILE) mytemp}
        set mytemp [file join $mytemp Temp]
      }
    }
  }

# construct directories from drive and env(USERNAME)
  if {[info exists env(USERNAME)] && $myhome == $drive} {
    set myhome [file join $drive Users $env(USERNAME)]
    if {$tcl_platform(osVersion) < 6.0} {set myhome [file join $drive "Documents and Settings" $env(USERNAME)]}
  }

  if {![info exists mydocs]} {
    set mydocs $myhome
    set docs "Documents"
    if {$tcl_platform(osVersion) < 6.0} {set docs "My Documents"}
    set docs [file join $mydocs $docs]
    if {[file exists $docs]} {if {[file isdirectory $docs]} {set mydocs $docs}}
  }

  if {![info exists mydesk]} {
    set mydesk $myhome
    set desk "Desktop"
    set desk [file join $mydesk $desk]
    if {[file exists $desk]} {if {[file isdirectory $desk]} {set mydesk $desk}}
  }

  if {![info exists mytemp]} {
    set mytemp $myhome
    set temp [file join AppData Local Temp]
    if {$tcl_platform(osVersion) < 6.0} {set temp [file join "Local Settings" Temp]}
    set temp [file join $mytemp $temp]
    if {[file exists $temp]} {if {[file isdirectory $temp]} {set mytemp $temp}}
  }

  set myhome [file nativename $myhome]
  set mydocs [file nativename $mydocs]
  set mydesk [file nativename $mydesk]
  set mytemp [file nativename $mytemp]
  set drive [string range $myhome 0 2]
}
