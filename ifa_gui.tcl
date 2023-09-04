proc getVersion {} {return 3.07}

# see proc installIFCsvr in ifa_proc.tcl for the IFCsvr version

#-------------------------------------------------------------------------------
# start window, bind keys
proc guiStartWindow {} {
  global fout lastXLS lastXLS1 localName localNameList wingeo winpos

  wm title . "IFC File Analyzer [getVersion]"

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
  global buttons ftrans mytemp nline nprogfile wdir

  set ftrans [frame .ftrans1 -bd 2 -background "#F0F0F0"]
  set buttons(genExcel) [ttk::button $ftrans.generate1 -text "Generate Spreadsheet" -padding 4 -state disabled -command {
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

  catch {
    set l3 [label $ftrans.l3 -relief flat -bd 0]
    $l3 config -image [image create photo -file [file join $wdir images nist.gif]]
    pack $l3 -side right -padx 10
    bind $l3 <ButtonRelease-1> {displayURL https://www.nist.gov}
    tooltip::tooltip $l3 "Click here to learn more about NIST"
  }

  pack $ftrans -side top -padx 10 -pady 10 -fill x

  set fbar [frame .fbar -bd 2 -background "#F0F0F0"]
  set nline 0
  set buttons(pgb) [ttk::progressbar $fbar.pgb -mode determinate -variable nline]
  pack $fbar.pgb -side top -padx 10 -fill x

  set nprogfile 0
  set buttons(pgb1) [ttk::progressbar $fbar.pgb1 -mode determinate -variable nprogfile]
  pack forget $buttons(pgb1)
  pack $fbar -side bottom -padx 10 -pady {0 10} -fill x

# icon bitmap
  catch {file copy -force [file join $wdir images NIST.ico] $mytemp}
  catch {wm iconbitmap . -default [file join $mytemp NIST.ico]}
}

#-------------------------------------------------------------------------------
# status tab
proc guiStatusTab {} {
  global fout nb outputWin statusFont wout

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

  if {![info exists statusFont]} {set statusFont [$outputWin type cget black -font]}
  if {[string first "Courier" $statusFont] != -1} {
    regsub "Courier" $statusFont "Consolas" statusFont
    regsub "120" $statusFont "140" statusFont
    saveState
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
  global File openFileList

  $File add command -label "Open IFC File(s)..." -accelerator "Ctrl+O" -command openFile
  $File add command -label "Open Multiple IFC Files in a Directory..." -accelerator "F4" -command {openMultiFile}
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
  $File add command -label "Exit" -accelerator "Ctrl+Q" -command exit
}


#-------------------------------------------------------------------------------
# options tab, process
proc guiProcess {} {
  global allNone buttons cb fopt fopta nb opt type

  set cb 0
  set wopt [ttk::panedwindow $nb.opt -orient horizontal]
  $nb add $wopt -text " Options " -padding 2
  set fopt [frame $wopt.fopt -bd 2 -relief sunken]

  set fopta [ttk::labelframe $fopt.a -text " Process "]

  # option to process user-defined entities
  guiUserDefinedEntities
  set txt1 "Process categories control which entities are written to the Spreadsheet.  See Help > IFC Support\nThe categories are used to group and color-code entities on the File Summary worksheet."
  set txt2 "\n\n"

  set fopta1 [frame $fopta.1 -bd 0]
  foreach item {{" Building Elements" opt(PR_BEAM)} \
                {" HVAC"              opt(PR_HVAC)} \
                {" Electrical"        opt(PR_ELEC)} \
                {" Building Services" opt(PR_SRVC)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists type($tt)]} {
      set ttmsg "$txt1\n\nThere are [llength $type($tt)] [string trim [lindex $item 0]] entities.$txt2"
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta2 [frame $fopta.2 -bd 0]
  foreach item {{" Infrastructure" opt(PR_INFR)} \
                {" Profile"        opt(PR_PROF)} \
                {" Material"       opt(PR_MTRL)} \
                {" Property"       opt(PR_PROP)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    set txt3 $txt2
    if {[lindex $item 0] == " Infrastructure"} {set txt3 "  These entities are supported in IFC4x2 and greater.  See Websites > IFC Infrastructure\n\n"}

    if {[info exists type($tt)]} {
      set ttmsg "$txt1\n\nThere are [llength $type($tt)] [string trim [lindex $item 0]] entities.$txt3"
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta3 [frame $fopta.3 -bd 0]
  foreach item {{" Representation" opt(PR_REPR)} \
                {" Relationship"   opt(PR_RELA)} \
                {" Presentation"   opt(PR_PRES)} \
                {" Analysis"       opt(PR_ANAL)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists type($tt)]} {
      set ttmsg "$txt1\n\nThere are [llength $type($tt)] [string trim [lindex $item 0]] entities.$txt2"
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta4 [frame $fopta.4 -bd 0]
  foreach item {{" Geometry" opt(PR_GEOM)} \
                {" Quantity" opt(PR_QUAN)} \
                {" Unit"     opt(PR_UNIT)} \
                {" Other"    opt(PR_COMM)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta4.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists type($tt)]} {
      set ttmsg "$txt1\n\nThere are [llength $type($tt)] [string trim [lindex $item 0]] entities.$txt2"
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta5 [frame $fopta.5 -bd 0]
  set anbut [list {"All" 0} {"None" 1}]
  foreach item $anbut {
    set bn "allNone[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $fopta5.$cb -variable allNone -text [lindex $item 0] -value [lindex $item 1] \
      -command {
        if {$allNone == 0} {
          foreach item [array names opt] {if {[string first "PR_" $item] == 0 && $item != "PR_USER"} {set opt($item) 1}}
        } elseif {$allNone == 1} {
          foreach item [array names opt] {if {[string first "PR_" $item] == 0} {set opt($item) 0}}
          set opt(PR_BEAM) 1
        }
        checkValues
      }]
    pack $buttons($bn) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  catch {
    tooltip::tooltip $buttons(allNone0) "Select all Process categories"
    tooltip::tooltip $buttons(allNone1) "Deselect most Process categories"
  }
  pack $fopta5 -side left -anchor w -pady 0 -padx 15 -fill y

  pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both
}

#-------------------------------------------------------------------------------
# overview
proc helpOverview {} {
  outputMsg "\nOverview ------------------------------------------------------------------------------------------" blue
  outputMsg "The IFC File Analyzer reads an IFC file and generates an Excel spreadsheet or CSV files.  One
worksheet or CSV file is generated for each entity type in the IFC file.  Each worksheet or CSV
file lists every entity instance and its attributes.  The types of entities that are Processed can
be selected in the Options tab.  Other options are available that add to or modify the information
written to the spreadsheet or CSV files.

For spreadsheets, a Summary worksheet shows the Count of each entity.  Links on the Summary and
entity worksheets can be used to navigate to other worksheets and to access IFC entity
documentation.

Spreadsheets or CSV files can be selected.  CSV files are generated if Excel is not installed.

To generate a spreadsheet or CSV files, select an IFC file from the File menu above and click the
Generate button below.  Existing spreadsheet or CSV files are always overwritten.

Multiple IFC files can be selected or an entire directory structure of IFC files can also be
processed from the File menu.  If multiple IFC files are translated, then a separate File Summary
spreadsheet is also generated.  This is useful to compare entity usage between different IFC files.

Tooltip help is available for the selections in the tabs.  Hold the mouse over text in the tabs
until a tooltip appears.

To run Syntax Checking on an IFC file, use the NIST STEP File Analyzer.

See Help > Function Keys to change the font size in the Status tab.  Right-click to save the text.
See Help > IFC Support
See Help > Disclaimers and NIST Disclaimer"

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# IFC support
proc helpSupport {} {
  global ifcsvrDir

  set schemas {}
  foreach match [lsort [glob -nocomplain -directory $ifcsvrDir *.rose]] {
    set schema [string toupper [file rootname [file tail $match]]]
    if {[string first "IFC" $schema] == 0 && [string first "151" $schema] == -1 && [string first "LONGFORM" $schema] == -1 && [string first "PLATFORM" $schema] == -1 && \
        [string first "2X3_RC" $schema] == -1 && [string first "FINAL" $schema] == -1} {
      regsub -all "X" $schema "x" schema
      lappend schemas $schema
    }
  }
  set schemas [join $schemas " "]
  regsub -all " " $schemas ", " schemas
  set c1 [string last "," $schemas]
  if {$c1 != -1} {set schemas "[string range $schemas 0 $c1] and[string range $schemas $c1+1 end]"}
  outputMsg "\nIFC Support ---------------------------------------------------------------------------------------" blue

  if {$schemas != ""} {
outputMsg "$schemas are supported with the following exceptions.

For IFC4x2 and IFC4x3, all entities related to TEXTURE are not supported and will not be reported
in the spreadsheet.  However, other entities that refer to them might cause a crash.  If necessary,
uncheck Presentation in the Process section.

---------------------------------------------------------------------------------------------------
For IFC4 only, these Geometry entities are not supported and will not be reported in the
spreadsheet.

 IfcCartesianPointList2D  IfcIndexedPolyCurve  IfcIndexedPolygonalFace
 IfcIndexedPolygonalFaceWithVoids  IfcIntersectionCurve  IfcPolygonalFaceSet  IfcSeamCurve
 IfcSphericalSurface  IfcSurfaceCurve  IfcToroidalSurface
 
However, other entities that refer to them might cause a crash.  If necessary, uncheck Profile,
Representation, and Geometry in the Process section.

You can also edit the IFC file and change FILE_SCHEMA(('IFC4')); to FILE_SCHEMA(('IFC4X3')); to
process the geometry entities.  All Process categories can be selected.

---------------------------------------------------------------------------------------------------
Tooltips in the Process section indicate which entities are specific to IFC4 or greater.

Unicode in text strings (\\X2\\ encoding) used for symbols and accented or non-English characters are
not supported.  Those characters will be missing from text strings.

See Websites > IFC Specifications"

  } else {
    errorMsg "No IFC schemas are supported because the IFCsvr toolkit has not been installed."
  }

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global Help ifcsvrKey row_limit

  $Help add command -label "Overview" -command {helpOverview}

# options help
  $Help add command -label "Options" -command {
    outputMsg "\nOptions -------------------------------------------------------------------------------------------" blue
    outputMsg "Process: Select which types of entities are processed.  The tooltip help lists all the entities
associated with that type.  Selectively process only the entities relevant to your analysis.

Inverse Relationships: For many entity types some Inverse Relationships are displayed on the
worksheets.  The Inverse values are displayed in additional columns of entity worksheets that are
highlighted in light blue.

Expand: The attributes that IfcPropertySet, IfcLocalPlacement, IfcAxis2Placement, or structural
analysis entities refer to will be displayed inline with the entity. For example, IfcLocalPlacement
refers to an IfcAxis2Placement3D and an optional relative placement. Those values would be included
in addition to the IfcLocalPlacement. IfcAxis2Placement expands into an IfcCartesianPoint and
IfcDirection.  The columns with the expanded values are color coded.  The expanded columns can be
collapsed on a worksheet.

Generate: Excel spreadsheets or CSV files.  If Excel is not installed, CSV files are automatically
generated.  Some options are not supported with CSV files.

Table: Generate tables for each spreadsheet to facilitate sorting and filtering (Spreadsheet tab).

Number Format: Option to not round real numbers.

Count Duplicates: Entities with identical attribute values will be counted and not duplicated on a
worksheet.  This applies to a limited set of entities.

Maximum Rows: The maximum number of rows for any worksheet can be set lower than the normal limits
for Excel.  This is useful for very large IFC files at the expense of not processing some entities."

    .tnb select .tnb.status
    update idletasks
  }

  $Help add command -label "IFC Support" -command {helpSupport}
  $Help add separator

# open Function Keys help
  $Help add command -label "Function Keys" -command {
    outputMsg "\nFunction Keys -------------------------------------------------------------------------------------" blue
    outputMsg "Function keys can be used as shortcuts for several commands:

F1 - Generate Spreadsheet from the current or last IFC file
F2 - Open current or last Spreadsheet

F3 - Open current or last File Summary Spreadsheet generated from a set of multiple IFC files
F4 - Generate Speadsheets from current or last set of multiple IFC files

F5 - Decrease this font size
F6 - Increase this font size

F8 - Open IFC file in a text editor
Shift-F8 - Open IFC file directory"

    .tnb select .tnb.status
  }

# display files help
  $Help add command -label "Open IFC File in App" -command {
    outputMsg "\nOpen IFC File in App ------------------------------------------------------------------------------" blue
    outputMsg "This option is a convenient way to open an IFC file in other applications.  The pull-down menu will
contain applications that can open an IFC file such as IFC viewers and browsers.  If applications
are installed in their default location, then they will appear in the pull-down menu.

The 'Indent IFC File (for debugging)' option rearranges and indents the entities to show the
hierarchy of information in an IFC file.  The 'indented' file is written to the same directory as
the IFC file or to the same user-defined directory specified in the Spreadsheet tab.  It is useful
for debugging IFC files but is not recommended for large IFC files.

The 'Default IFC Viewer' option will open the IFC file in whatever application is associated with
IFC files.

A text editor will always appear in the menu."

    .tnb select .tnb.status
    update idletasks
  }

# multiple files help
  $Help add command -label "Multiple IFC Files" -command {
    outputMsg "\nMultiple IFC Files --------------------------------------------------------------------------------" blue
    outputMsg "Multiple IFC files can be selected in the Open File(s) dialog by holding down the control or shift
key when selecting files or an entire directory of IFC files can be selected with 'Open Multiple
IFC Files in a Directory'. Files in subdirectories of the selected directory can also be processed.

When processing multiple IFC files, a File Summary spreadsheet is generated in addition to
individual spreadsheets for each file.  The File Summary spreadsheet shows the entity count and
totals for all IFC files.  The File Summary spreadsheet also links to the individual spreadsheets
and the IFC file.

If only the File Summary spreadsheet is needed, it can be generated faster by turning off
Processing of most of the entity types and options in the Options tab."

    .tnb select .tnb.status
    update idletasks
  }

# number format help
  $Help add command -label "Number Format" -command {
    outputMsg "\nNumber Format -------------------------------------------------------------------------------------" blue
    outputMsg "By default Excel rounds real numbers if there are more than 11 characters in the number string.

For example, the number 0.12499999999999997 in the IFC file will be displayed as 0.125.  However,
double clicking in a cell with a rounded number will show all of the digits.

This option will display most real numbers exactly as they appear in the IFC file.  This applies
only to single real numbers.  Lists of real numbers, such as cartesian point coordinates, are
always displayed exactly as they appear in the IFC file.

Rounding real numbers might affect how Count Duplicates appears.  If both 0.12499999999999997 and
0.12499999999999993 are rounded to 0.125 they will appear as two separate values of 0.125 when it
would seem that they are identical each other."

    .tnb select .tnb.status
    update idletasks
  }

# count duplicates help
  $Help add command -label "Count Duplicates" -command {
    outputMsg "\nCount Duplicates ----------------------------------------------------------------------------------" blue
    outputMsg "When using the Count Duplicates option on the Spreadsheet tab, entities with identical attribute
values will be counted and not duplicated on a worksheet.  The resulting entity worksheets might be
shorter.

Some entity attributes might be ignored to check for duplicates.  The entity count is displayed in
the last column of the worksheet.  The entity ID displayed is of the first of the duplicate
entities.

If there are no duplicates for an entity type being counted and there are a lot (> 50000) of that
entity type, then the processing can be slow.  This is most common with Geometry entities.

The list of IFC entities that are counted is displayed in the Count Duplicates tooltip on the
Options tab."

    .tnb select .tnb.status
    update idletasks
  }

# large files help
  $Help add command -label "Large IFC Files" -command {
    outputMsg "\nLarge IFC Files -----------------------------------------------------------------------------------" blue
    outputMsg "The largest IFC file that can be processed for a Spreadsheet is approximately 400 MB.  Processing
larger IFC files might cause a crash.  Popup dialogs might appear that say 'unable to realloc xxx
bytes'.

Try some of these options to reduce the amount of time to process large IFC files that do not cause
a crash and to reduce the size of the resulting spreadsheet:
- Deselect entity types for which there are usually a lot of, such as Geometry and Property
- Use only the User-Defined List option to process specific entity types
- It might be necessary to process only one category of entities at a time to generate multiple
  spreadsheets
- Uncheck the options for Inverse Relationships and Expand
- Set the Maximum Rows for any worksheet"

    .tnb select .tnb.status
    update idletasks
  }

  $Help add command -label "Crash Recovery" -command {
    outputMsg "\nCrash Recovery ------------------------------------------------------------------------------------" blue
    outputMsg "Sometimes this software crashes after an IFC file has been successfully opened and the processing
of entities has started.  Popup dialogs might appear that say 'Runtime Error!' or
'ActiveState Basekit has stopped working' or 'Fatal Error in Wish - unable to alloc 123456 bytes'.

A crash is most likely due to syntax errors in the IFC file or sometimes due to limitations of the
toolkit used to read IFC files.  To see which type of entity caused the error, check the Status tab
to see which type of entity was last processed.  A crash can also be caused by insufficient memory
to process a very large IFC file.

Workarounds for these problems:

1 - Processing of the type of entity that caused the error can be deselected in the Options tab
under Process.  However, this will prevent processing of other entities that do not cause a crash.
The User-Defined List can be used to process only the required entity types.

2 - Use the Syntax Checker in the NIST STEP File Analyzer."

    .tnb select .tnb.status
    update idletasks
  }

  $Help add separator
  $Help add command -label "Disclaimers" -command {
    outputMsg "\nDisclaimers ---------------------------------------------------------------------------------------" blue
    outputMsg "Please see Help > NIST Disclaimer for the Software Disclaimer.

Any mention of commercial products or references to web pages is for information purposes only; it
does not imply recommendation or endorsement by NIST.  For any of the web links, NIST does not
necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses Microsoft Excel and IFCsvr that are covered by their own Software License
Agreements.  See Help > About.

If you are using this software in your own application, please explicitly acknowledge NIST as the
source of the software."
    .tnb select .tnb.status
    update idletasks
  }

  $Help add command -label "NIST Disclaimer" -command {displayURL https://www.nist.gov/disclaimer}
  $Help add command -label "About" -command {
    outputMsg "\nIFC File Analyzer ---------------------------------------------------------------------------------" blue
    outputMsg "Version: [getVersion] ([string trim [clock format $progtime -format "%e %b %Y"]])"

    set winver ""
    if {[catch {
      set winver [registry get {HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion} {ProductName}]
    } emsg]} {
      set winver "$tcl_platform(os) $tcl_platform(osVersion)"
    }
    set sysvar "System: $winver"
    catch {append sysvar ", IFCsvr [registry get $ifcsvrKey {DisplayVersion}]"}
    outputMsg $sysvar
    if {[string first "Windows Server" $winver] != -1 || $tcl_platform(osVersion) < 6.1} {errorMsg " $winver is not supported."}

    outputMsg "\nThe IFC File Analyzer was developed at NIST in the former Computer Integrated Building Processes
Group in the Building and Fire Research Laboratory.  The software was first released in 2008 and
development ended in 2014.  Minor updates have been made since 2014.  IFC4xN versions were added
in 2021.

Credits
- Reading and parsing IFC files:
   IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
   IFCsvr has been modified by NIST to include newer IFC4xN versions.
   The license agreement can be found in C:\\Program Files (x86)\\IFCsvrR300\\doc

See Help > Disclaimers and NIST Disclaimer"

# debug
    if {$row_limit == 100003} {
      outputMsg " "
      outputMsg "SFA variables" red
      catch {outputMsg " Drive $drive"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp  ([file exists $mytemp])"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      catch {outputMsg " scriptName $scriptName"}
      outputMsg " Tcl [info patchlevel], twapi [package versions twapi]"

      outputMsg "Registry values" red
      catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
      catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
      catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
      catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}

      outputMsg "Environment variables" red
      foreach id [lsort [array names env]] {
        foreach id1 [list HOME Program System USER TEMP TMP APP] {if {[string first $id1 $id] == 0} {outputMsg " $id  $env($id)"; break}}
      }
    }

    .tnb select .tnb.status
    update idletasks
  }
}

#-------------------------------------------------------------------------------
# Websites menu
proc guiWebsitesMenu {} {
  global Websites

  $Websites add command -label "IFC File Analyzer"        -command {displayURL https://www.nist.gov/services-resources/software/ifc-file-analyzer}
  $Websites add separator
  $Websites add command -label "Technical Resources"      -command {displayURL https://technical.buildingsmart.org/}
  $Websites add command -label "IFC Specifications"       -command {displayURL https://technical.buildingsmart.org/standards/ifc/ifc-schema-specifications/}
  $Websites add command -label "Software Implementations" -command {displayURL https://technical.buildingsmart.org/resources/software-implementations/}
  $Websites add command -label "Infrastructure Room"      -command {displayURL https://www.buildingsmart.org/standards/rooms/infrastructure/}
  $Websites add command -label "buildingSMART"            -command {displayURL https://www.buildingsmart.org/}
  $Websites add command -label "ISO 16739"                -command {displayURL https://www.iso.org/standard/70303.html}
  $Websites add separator
  $Websites add command -label "Free IFC Software"        -command {displayURL https://www.ifcwiki.org/index.php/Freeware}
  $Websites add command -label "Common BIM Files"         -command {displayURL https://www.wbdg.org/bim/cobie/common-bim-files}
  $Websites add command -label "IFC Format"               -command {displayURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000447.shtml}
  $Websites add command -label "IFC Wikipedia"            -command {displayURL https://en.wikipedia.org/wiki/Industry_Foundation_Classes}
  $Websites add command -label "Source code on GitHub"    -command {displayURL https://github.com/usnistgov/IFA}
}

#-------------------------------------------------------------------------------
# user-defined list of entities
proc guiUserDefinedEntities {} {
  global buttons cb fileDir fopta opt userEntityFile

  set fopta6 [frame $fopta.6 -bd 0]
  foreach item {{" User-Defined List: " opt(PR_USER)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
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
    catch {tooltip::tooltip $buttons($item) "A User-Defined List is a text file with one IFC entity name per line.\nThis allows for more control to process only the required entity types.\nIt is also useful when processing large files that might cause a crash."}
  }
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# display result
proc guiDisplayResult {} {
  global appName appNames buttons cb dispApps dispCmds fopt foptf

  set foptf [ttk::labelframe $fopt.f -text " Open IFC File in App "]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 30]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}
  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]

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

  catch {tooltip::tooltip $foptf "This option is a convenient way to open an IFC file in other applications.\nThe pull-down menu will contain applications that can open an IFC file\nsuch as IFC viewers and file browsers.  If applications\nare installed in their default location, then they will appear in the\npull-down menu.\n\nThe 'Indent IFC File (for debugging)' option rearranges and indents the\nentities to show the hierarchy of information in an IFC file.  The 'indented'\nfile is written to the same directory as the IFC file or to the same\nuser-defined directory specified in the Spreadsheet tab.\n\nThe 'Default IFC Viewer' option will open the IFC file in whatever\napplication is associated with IFC files."}
  pack $foptf -side top -anchor w -pady {5 2} -padx 10 -fill both

# generate
  set foptk [ttk::labelframe $fopt.k -text " Generate "]
  foreach item {{" Spreadsheet" Excel} {" CSV Files" CSV}} {
    pack [ttk::radiobutton $foptk.$cb -variable opt(XLSCSV) -text [lindex $item 0] -value [lindex $item 1] -command {checkValues}] -side left -anchor n -padx 5 -pady {0 2} -ipady 0
    incr cb
  }
  set item {" Open Output Files" opt(XL_OPEN)}
  regsub -all {[\(\)]} [lindex $item 1] "" idx
  set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor n -padx 7 -pady {0 2} -ipady 0
  incr cb
  pack $foptk -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $foptk "Microsoft Excel is required to generate spreadsheets.\n\nCSV files will be generated if Excel is not installed.\nOne CSV file is generated for each entity type.\nSome of the options are not supported with CSV files."}
}

#-------------------------------------------------------------------------------
# inverse relationships and expand
proc guiInverseExpand {} {
  global buttons cb fopt inverses opt

  set foptIE [frame $fopt.d2 -bd 0]
  set foptc [ttk::labelframe $foptIE.3 -text " Inverse Relationships "]
  regsub -all {[\(\)]} opt(INVERSE) "" idx
  set buttons($idx) [ttk::checkbutton $foptc.$cb -text " Show Inverses" -variable opt(INVERSE) -command {
      checkValues
      if {$opt(INVERSE)} {set opt(PR_RELA) 1}
    }]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady {0 2} -ipady 0
  incr cb
  pack $foptc -side left -anchor w -pady {5 2} -padx 10 -fill both

  set ttmsg "Inverse Relationships are shown on entity worksheets in additional\ncolumns that are highlighted in light blue.\n"
  set txt ""
  foreach item [lsort $inverses] {
    set inv [lindex $item 0]
    if {[string first $inv $txt] == -1} {
      append txt "[lindex $item 0]  "
      if {[string length $txt] > 50} {
        append ttmsg "\n$txt"
        set txt ""
      }
    }
  }
  if {$txt != ""} {append ttmsg "\n$txt"}
  catch {tooltip::tooltip $foptc $ttmsg}

  set foptd [ttk::labelframe $foptIE.1 -text " Expand "]
  set foptd1 [frame $foptd.1 -bd 0]
  foreach item {{" Properties"     opt(EX_PROP)} \
                {" LocalPlacement" opt(EX_LP)} \
                {" Axis2Placement" opt(EX_A2P3D)} \
                {" Analysis"       opt(EX_ANAL)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady {0 2} -ipady 0
    incr cb
  }
  pack $foptd1 -side left -anchor w -pady 0 -padx 0 -fill y
  pack $foptd -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

  catch {tooltip::tooltip $foptd "These options expand the selected entity attributes that are referred to on an entity\nbeing processed.\n\n- Properties shows individual property values for IfcPropertySet, IfcElementQuantity,\n   IfcMaterialProperties, IfcProfileProperties, and IfcComplexProperty.\n\n- LocalPlacement shows the attribute values of PlacementRelTo and\n   RelativePlacement for IfcLocalPlacement for every building element.\n- Axis2Placement shows the corresponding attribute values for Location,\n   Axis, and RefDirection.  This option does not work well where building elements\n   of the same type have different levels of coordinate system nesting.\n\n- Analysis applies to structural loads, reactions, and displacements.\n\nFor LocalPlacement and Axis2Placement, the columns used for the expanded\nentities are grouped together and displayed with different colors.  Use the \"-\"\nsymbols above the columns or the \"1\" at the top left of the spreadsheet to\ncollapse the columns."}
  pack $foptIE -side top -anchor w -pady 0 -fill x
}

#-------------------------------------------------------------------------------
# spreadsheet tab
proc guiSpreadsheet {} {
  global buttons cb countent fileDir fxls mydocs nb opt row_limit userWriteDir writeDir writeDirType

  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " Spreadsheet " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

  set fxlsb [ttk::labelframe $fxls.b -text " Maximum Rows "]
  set rlimit {{" 100" 103} {" 500" 503} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003} {" 100000" 100003} {" Maximum" 1048576}}
  foreach item $rlimit {
    pack [ttk::radiobutton $fxlsb.$cb -variable row_limit -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "This option will limit the number of rows (entities) written to any one worksheet.\nFor large IFC files, setting a low maximum can speed up processing at the expense\nof not processing all of the entities.  This is useful when processing Geometry entities."
  append msg "\n\nIf the maximum number of rows is exceeded, then the counts on the summary\nworksheet for Name, Description, etc. might not be correct."
  catch {tooltip::tooltip $fxlsb $msg}

  set fxlsz [ttk::labelframe $fxls.z -text " Formatting "]
  foreach item {{" Generate Tables for sorting and filtering" opt(SORT)} \
                {" Do not round real numbers in spreadsheet cells" opt(XL_FPREC)} \
                {" Count Duplicate identical entities" opt(COUNT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsz.$cb -text [lindex $item 0] -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsz -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {
    tooltip::tooltip $buttons(optSORT) "Worksheets can be sorted by column values."
    tooltip::tooltip $buttons(optXL_FPREC) "See Help > Number Format"

    set ttmsg ""
    if {[info exists countent(IFC)]} {
      set ttlen 0
      set lchar ""
      foreach item [lsort $countent(IFC)] {
        incr ttlen [expr {[string length $item]+2}]
        if {$ttlen <= 120} {
          append ttmsg "$item  "
        } else {
          if {[string index $ttmsg end] != "\n"} {set ttmsg "[string range $ttmsg 0 end-2]\n$item  "}
          set ttlen [expr {[string length $item]+2}]
        }
      }
    }
    set tmsg "Entities with identical attribute values will be counted and not duplicated on a worksheet.  The resulting entity worksheets\nmight be shorter.  See Help > Count Duplicates.  These IFC entities have duplicates counted:\n\n$ttmsg"
    tooltip::tooltip $buttons(optCOUNT) $tmsg
  }

  set fxlsd [ttk::labelframe $fxls.d -text " Write Output to "]
  set buttons(fileDir) [ttk::radiobutton $fxlsd.$cb -text " Same directory as the IFC file" -variable writeDirType -value 0 -command checkValues]
  pack $fxlsd.$cb -side top -anchor w -padx 5 -pady 2
  incr cb

  set fxls1 [frame $fxlsd.1]
  ttk::radiobutton $fxls1.$cb -text " User-defined directory:  " -variable writeDirType -value 2 -command {
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
  catch {tooltip::tooltip $fxls1.$cb "This option can be used when the directory containing the IFC file is\nprotected (read-only) and none of the output can be written to it."}
  incr cb

  set buttons(userentry) [ttk::entry $fxls1.entry -width 50 -textvariable userWriteDir]
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
  pack $fxlsd -side top -anchor w -pady {5 2} -padx 10 -fill both

  set fxlsc [ttk::labelframe $fxls.c -text " Other "]
  foreach item {{" When processing Multiple Files, do not generate links to IFC files and spreadsheets on File Summary worksheet" opt(HIDELINKS)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsc.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsc -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {
    tooltip::tooltip $buttons(optHIDELINKS) "This option is useful when sharing a Spreadsheet with another user."
  }
  pack $fxls -side top -fill both -expand true -anchor nw
}

#-------------------------------------------------------------------------------
# shortcuts
proc setShortcuts {} {
  global mydesk mymenu mytemp

  set progname [info nameofexecutable]
  if {[string first "AppData/Local/Temp" $progname] != -1 || [string first ".zip" $progname] != -1} {
    errorMsg "You should first extract all of the files from the ZIP file and run the extracted software."
    return
  }

  if {[info exists mydesk] || [info exists mymenu]} {
    set progstr "IFC File Analyzer"

    set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" -message "Do you want to create or overwrite shortcuts to the $progstr [getVersion]"]
    if {$choice == "yes"} {
      outputMsg " "
      catch {if {[info exists mymenu]} {twapi::write_shortcut [file join $mymenu "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $mytemp NIST.ico]}}
      catch {if {[info exists mydesk]} {twapi::write_shortcut [file join $mydesk "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $mytemp NIST.ico]}}
    }
  }
}

#-------------------------------------------------------------------------------
# set home, docs, desktop, menu directories
proc setHomeDir {} {
  global drive env mydesk mydocs myhome mymenu mytemp

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
      if {[string first "%USERPROFILE%" $reg_personal] == 0} {set mydocs "$env(USERPROFILE)\\[string range $reg_personal 14 end]"}
    }
    catch {
      set reg_desktop  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]
      if {[string first "%USERPROFILE%" $reg_desktop] == 0} {set mydesk "$env(USERPROFILE)\\[string range $reg_desktop 14 end]"}
    }
    catch {
      set reg_menu [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {set mymenu "$env(USERPROFILE)\\[string range $reg_menu 14 end]"}
    }
    catch {
      set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]
      if {[string first "%USERPROFILE%" $reg_temp] == 0} {set mytemp "$env(USERPROFILE)\\[string range $reg_temp 14 end]"}
      set mytemp [file join $mytemp Temp]
    }
  }

# construct directories from drive and env(USERNAME)
  if {[info exists env(USERNAME)] && $myhome == $drive} {
    set myhome [file join $drive Users $env(USERNAME)]
  }

  if {![info exists mydocs]} {
    set mydocs $myhome
    set docs "Documents"
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
    set temp [file join $mytemp $temp]
    if {[file exists $temp]} {if {[file isdirectory $temp]} {set mytemp $temp}}
  }

  set myhome [file nativename $myhome]
  set mydocs [file nativename $mydocs]
  set mydesk [file nativename $mydesk]
  set mytemp [file nativename $mytemp]
  set drive [string range $myhome 0 2]
}
