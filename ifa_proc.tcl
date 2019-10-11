#-------------------------------------------------------------------------------
proc getTiming {{str ""}} {
  global tlast

  set t [clock clicks -milliseconds]
  if {[info exists tlast]} {outputMsg "Timing: [expr {($t-$tlast)}]  $str" red}
  set tlast $t
}
 
#-------------------------------------------------------------------------------
proc getSchema {fname {msg 0}} {
  
  set schema ""
  set ok 0
  set nline 0
  set stepfile [open $fname r]
  while {[gets $stepfile line] != -1 && $nline < 50} {
    if {$msg} {
      foreach item {"MIME-Version" "Content-Type" "X-MimeOLE" "DOCTYPE HTML" "META content"} {
        if {[string first $item $line] != -1} {
          errorMsg "Syntax Error: The IFC file has probably been saved as an EMAIL or HTML file from\n Outlook or other email client.  The IFC file cannot be translated.  In the email\n client, save the IFC file as a TEXT file and try again.  The first line in the\n IFC file should be 'ISO-10301-21\;'"
        }
      }
    }

    incr nline
    if {[string first "FILE_SCHEMA" $line] != -1} {
      set ok 1
      set fsline $line
    } elseif {[string first "ENDSEC" $line] != -1} {
      set schema [lindex [split $fsline "'"] 1]
      if {$msg} {
        errorMsg "The schema used is: $fsline" red
      }
      break
    } elseif {$ok} {
      append fsline $line
    }
  }
  close $stepfile
  return $schema
}
 
#-------------------------------------------------------------------------------

proc memusage {{str ""}} {
  global anapid lastmem
  
  if {![info exists lastmem]} {set lastmem 0}
  set mem [lindex [twapi::get_process_info $anapid -workingset] 1]
  set dmem [expr {$mem-$lastmem}]
  outputMsg "  $str  dmem [expr {$dmem/1000}]  mem [expr {$mem/1000}]" red
  set lastmem $mem
}
 
#-------------------------------------------------------------------------------

proc processToolTip {ttmsg tt {ttlim 120}} {
  global type ifcProcess
 
  set ttlen 0
  set lchar ""
  set r1 3

  foreach item [lsort $type($tt)] {
    if {[string range $item 0 $r1] != $lchar && $lchar != ""} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
    }
    append ttmsg "$item   "
    incr ttlen [string length $item]
    if {$ttlen > $ttlim} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
      set ok 0
    }
    set lchar [string range $item 0 $r1]
    lappend ifcProcess $item
  }
  return $ttmsg
}
 
#-------------------------------------------------------------------------------
proc checkValues {} {
  global opt buttons appNames appName writeDirType userentlist
  global edmWriteToFile edmWhereRules eeWriteToFile

  if {[info exists buttons(appCombo)]} {
    set ic [lsearch $appNames $appName]
    if {$ic < 0} {set ic 0}
    $buttons(appCombo) current $ic
    catch {
      if {[string first "EDM Model Checker" $appName] == 0} {
        pack $buttons(edmWriteToFile)  -side left -anchor w -padx 5
        pack forget $buttons(edmWhereRules)
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
  }

  if {$opt(XLSCSV) == "CSV"} {
    set opt(INVERSE) 0
    set opt(writeDirType) 0
    set opt(EX_ANAL) 0
    set opt(EX_A2P3D) 0
    set opt(EX_LP) 0
    set opt(COUNT) 0
    $buttons(optEX_ANAL) configure -state disabled
    $buttons(optEX_A2P3D) configure -state disabled
    $buttons(optEX_LP) configure -state disabled
    $buttons(optCOUNT) configure -state disabled
    $buttons(optINVERSE) configure -state disabled
    $buttons(optSORT)    configure -state disabled
    $buttons(optXL_FPREC) configure -state disabled
    $buttons(optXL_LINK1) configure -state disabled
    $buttons(optPR_GUID) configure -state disabled
    $buttons(genExcel)   configure -text "Generate CSV Files"
  } else {
    $buttons(optEX_ANAL) configure -state normal
    $buttons(optEX_A2P3D) configure -state normal
    $buttons(optEX_LP) configure -state normal
    $buttons(optCOUNT) configure -state normal
    $buttons(optINVERSE) configure -state normal
    $buttons(optSORT)    configure -state normal
    $buttons(optXL_FPREC) configure -state normal
    $buttons(optXL_LINK1) configure -state normal
    $buttons(optPR_GUID) configure -state normal
    $buttons(genExcel)   configure -text "Generate Spreadsheet"
  }

# IFC related
  if {$opt(XLSCSV) == "Excel"} {
    if {[info exists buttons(optEX_ANAL)]} {
      if {$opt(PR_ANAL)} {
        $buttons(optEX_ANAL) configure -state normal
      } else {
        $buttons(optEX_ANAL) configure -state disabled
        set opt(EX_ANAL) 0
      }
    }

    if {[info exists buttons(optEX_A2P3D)]} {
      if {$opt(EX_LP)} {
        $buttons(optEX_A2P3D) configure -state normal
      } else {
        $buttons(optEX_A2P3D) configure -state disabled
        set opt(EX_A2P3D) 0
      }
    }
  }
  
# user-defined entity list
  if {[info exists opt(PR_USER)]} {
    if {$opt(PR_USER)} {
      $buttons(userentity) configure -state normal
      $buttons(userentityopen) configure -state normal
    } else {
      $buttons(userentity) configure -state disabled
      $buttons(userentityopen) configure -state disabled
      set userentlist {}
    }
  }
  
  if {$writeDirType == 0} {
    $buttons(userentry) configure -state disabled
    $buttons(userdir) configure -state disabled
    $buttons(userentry1) configure -state disabled
    $buttons(userfile) configure -state disabled
  } elseif {$writeDirType == 1} {
    $buttons(userentry) configure -state disabled
    $buttons(userdir) configure -state disabled
    $buttons(userentry1) configure -state normal
    $buttons(userfile) configure -state normal
  } elseif {$writeDirType == 2} {
    $buttons(userentry) configure -state normal
    $buttons(userdir) configure -state normal
    $buttons(userentry1) configure -state disabled
    $buttons(userfile) configure -state disabled
  }

# make sure there is some entity type to process
  set nopt 0
  foreach idx [array names opt] {
    if {([string first "PR_" $idx] == 0) && \
         [string first "GUID" $idx] == -1} {incr nopt $opt($idx)}
  }
  if {$nopt == 0} {
    set opt(PR_BEAM) 1
    set opt(PR_HVAC) 1
    set opt(PR_ELEC) 1
    set opt(PR_SRVC) 1
  }
}
 
# -------------------------------------------------------------------------------------------------
proc entDocLink {sheet ent r c hlink} {
  global cells worksheet ifcdoc2x3 ifcall2x4 ifcDeprecated fileschema
  
  if {$sheet == "Summary"} {set c 3}

# IFC 2x3 doc link
  if {[string first "IFC4" $fileschema] == -1} {
    if {[info exists ifcdoc2x3([string tolower $ent])]} {
      set ent_link "https://standards.buildingsmart.org/IFC/RELEASE/IFC2x3/TC1/HTML/$ifcdoc2x3([string tolower $ent])/lexical/[string tolower $ent].htm"
      set str "IFC2x3"
      $cells($sheet) Item $r $c "Doc"
      set anchor [$worksheet($sheet) Range [cellRange $r $c]]
      if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
      if {[catch {
        $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $str Documentation"]
      } emsg]} {
        errorMsg "ERROR adding $sheet documentation link: $emsg"
        $cells($sheet) Item $r $c " "
        catch {raise .}
      }
      unset ent_link
    }
  }

# IFC2x4 doc or deprecated link
  if {[string first "IFC4" $fileschema] != -1} {
    set c1 $c
    if {[lsearch $ifcall2x4 $ent] != -1 || [lsearch $ifcDeprecated $ent] != -1} {
      if {[lsearch $ifcall2x4 $ent] != -1} {
        set ent_link "https://standards.buildingsmart.org/IFC/RELEASE/IFC4/FINAL/HTML/link/[string tolower $ent].htm"
        $cells($sheet) Item $r $c1 "Doc"
      } else {
        set ent_link "https://standards.buildingsmart.org/IFC/RELEASE/IFC4/FINAL/HTML/link/annex-f.htm"
        $cells($sheet) Item $r $c1 "Deleted"
      }
      set str "IFC4"
      set anchor [$worksheet($sheet) Range [cellRange $r $c1]]
      if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
      if {[catch {
        $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $str Documentation"]
      } emsg]} {
        errorMsg "ERROR adding $sheet documentation link: $emsg"
        $cells($sheet) Item $r $c1 " "
        catch {raise .}
      }
      unset ent_link
    }
  }

# if ent_link exists put the link except for IFC which is done above
  if {[info exists ent_link]} {
    $cells($sheet) Item $r $c "Doc"
    set anchor [$worksheet($sheet) Range [cellRange $r $c]]
    if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
    if {[catch {
      $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $str Documentation"]
    } emsg]} {
      errorMsg "ERROR adding $sheet documentation link: $emsg"
      $cells($sheet) Item $r $c " "
      catch {raise .}
    }
  }
}

# -------------------------------------------------------------------------------------------------
# set color based on tabcolor variable
proc setColorIndex {ifc {stat 0}} {
  global type tabcolor opt
  
# simple entity, not compound with _and_
  foreach i [array names type] {
    if {[info exist tabcolor($i)]} {
      if {[lsearch $type($i) $ifc] != -1} {
        return $tabcolor($i)
      }
    }
  }
  
# compound entity with _and_
  set c1 [string first "\_and\_" $ifc]
  if {$c1 != -1} {
    set c2 [string last  "\_and\_" $ifc]
    set tc1 "1000"
    set tc2 "1000"
    set tc3 "1000"
    
    foreach i [array names type] {
      if {[info exist tabcolor($i)]} {
        set ifc1 [string range $ifc 0 $c1-1]
        if {[lsearch $type($i) $ifc1] != -1} {
          #outputMsg "1 AND $ifc  $ifc1  $i  $tabcolor($i)"
          set tc1 $tabcolor($i)
        }
        if {$c2 == $c1} {
          set ifc2 [string range $ifc $c1+5 end]
          if {[lsearch $type($i) $ifc2] != -1} {
            #outputMsg "2 AND $ifc  $ifc2  $i  $tabcolor($i)"
            set tc2 $tabcolor($i)
          } 
        } elseif {$c2 != $c1} {
          set ifc2 [string range $ifc $c1+5 $c2-1]
          if {[lsearch $type($i) $ifc2] != -1} {
            #outputMsg "2 AND $ifc  $ifc2  $i  $tabcolor($i)"
            set tc2 $tabcolor($i)
          } 
          set ifc3 [string range $ifc $c2+5 end]
          if {[lsearch $type($i) $ifc3] != -1} {
            #outputMsg "3 AND $ifc  $ifc3  $i  $tabcolor($i)"
            set tc3 $tabcolor($i)
          }
        }
      }
    }
    set tc [expr {min($tc1,$tc2,$tc3)}]

    #outputMsg "TC $tc"
    if {$tc < 1000} {return $tc}
  }

# set color for some IFC entities that do not explicitly have a list
  if {$stat == 0} {
    if {$opt(PR_PROP) && \
      ([string first "Propert" $ifc] != -1 || \
       [string first "IfcDoorStyle" $ifc] == 0 || \
       [string first "IfcWindowStyle" $ifc] == 0)} {
      return 37
    } elseif {$opt(PR_QUAN) && [string first "Quantit" $ifc] != -1} {
      return 44
    } elseif {$opt(PR_MTRL) && [string first "Materia" $ifc] != -1} {
      return 36
    } elseif {$opt(PR_UNIT) && ([string first "Unit"   $ifc] != -1 || [string first "DimensionalExponent" $ifc] != -1)} {
      return 45
    } elseif {$opt(PR_RELA) && \
      ([string first "Relationship" $ifc] != -1 || \
       [string first "IfcRel" $ifc] == 0)} {
      return 39
    }
  } else {
    if { \
      ([string first "Propert" $ifc] != -1 || \
       [string first "IfcDoorStyle" $ifc] == 0 || \
       [string first "IfcWindowStyle" $ifc] == 0)} {
      return 37
    } elseif {[string first "Quantit" $ifc] != -1} {
      return 44
    } elseif {[string first "Materia" $ifc] != -1} {
      return 36
    } elseif {([string first "Unit"   $ifc] != -1 || [string first "DimensionalExponent" $ifc] != -1)} {
      return 45
    } elseif { \
      ([string first "Relationship" $ifc] != -1 || \
       [string first "IfcRel" $ifc] == 0)} {
      return 39
    }
  }
  return -2      
}

# -------------------------------------------------------------------------------------------------

proc getFirstFile {} {
  global openFileList remoteName buttons
  
  set localName [lindex $openFileList 0]
  if {$localName != ""} {
    set remoteName $localName
    outputMsg "\nReady to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
    if {[info exists buttons(appDisplay)]} {$buttons(appDisplay) configure -state normal}
    set fext [string tolower [file extension $localName]]
  }
  return $localName
}

#-------------------------------------------------------------------------------

proc displayURL {url} {
  global pf32

# open in whatever is registered for the file extension, except for .cgi for upgrade url
  if {[string first ".cgi" $url] == -1} {
    if {[catch {
      exec {*}[auto_execok start] "" $url
    } emsg]} {
      if {[string first "is not recognized" $emsg] == -1} {
        if {[string first "UNC" $emsg] == -1} {errorMsg "ERROR opening $url: $emsg"}
      }
    }

# find web browser command  
  } else {
    set webCmd ""
    catch {
      set reg_wb [registry get {HKEY_CURRENT_USER\Software\Classes\http\shell\open\command} {}]
      set reg_wb [lindex [split $reg_wb "\""] 1]
      set webCmd $reg_wb
    }
    if {$webCmd == "" || ![file exists $webCmd]} {set webCmd [file join $pf32 "Internet Explorer" IEXPLORE.EXE]}
    exec $webCmd $url &
  }
}

#-------------------------------------------------------------------------------
proc openFile {{openName ""}} {
  global localName localNameList fileDir buttons wdir mytemp

  if {$openName == ""} {
  
# file types for file select dialog
    set typelist {{"IFC Files" {".ifc" ".ifcZIP"}}}
    lappend typelist {"All Files" {*}}

# file open dialog
    set localNameList [tk_getOpenFile -title "Open IFC File(s)" -filetypes $typelist -initialdir $fileDir -multiple true]
    if {[llength $localNameList] <= 1} {set localName [lindex $localNameList 0]}
    catch {
      set fext [string tolower [file extension $localName]]
    }

# file name passed in as openName
  } else {
    set localName $openName
    set localNameList [list $localName]
  }

# multiple files selected
  if {[llength $localNameList] > 1} {
    set fileDir [file dirname [lindex $localNameList 0]]

    outputMsg "\nReady to process [llength $localNameList] IFC files" green
    $buttons(genExcel) configure -state normal
    if {[info exists buttons(appDisplay)]} {$buttons(appDisplay) configure -state normal}
    focus $buttons(genExcel)

# single file selected
  } elseif {[file exists $localName]} {
    set lcln [string tolower $localName]
  
# check for zipped file
    if {[string first ".stpz" $lcln] != -1 || [string first ".ifczip" $lcln] != -1} {
      if {[catch {
        outputMsg "Unzipping: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue

# copy gunzip to TEMP
        file copy -force [file join $wdir exe gunzip.exe] $mytemp
        set gunzip [file join $mytemp gunzip.exe]

# copy zipped file to TEMP
        if {[regsub ".ifcZIP" $localName ".ifc.Z" ln] == 0} {
          regsub ".ifczip" $localName ".ifc.Z" ln
        }
        set fzip [file join $mytemp [file tail $ln]]
        file copy -force $localName $fzip

# get name of unzipped file
        set ftmp [file join $mytemp [lindex [split [exec $gunzip -Nl $fzip] " "] end]]

# unzip
        if {[file tail $ftmp] != [file tail $fzip]} {outputMsg "Extracting: [file tail $ftmp]" blue}
        exec $gunzip -Nf $fzip

# copy to new stp file
        set fstp [file join [file dirname $localName] [file tail $ftmp]]
        if {[file exists $fstp]} {outputMsg " Overwriting existing IFC file: [truncFileName [file nativename $fstp]]" red}
        file copy -force $ftmp $fstp
        set localName $fstp
        file delete $fzip
        file delete $ftmp
      } emsg]} {
        errorMsg "ERROR unzipping file: $emsg"
      }
    }  
    set fileDir [file dirname $localName]

    outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
    $buttons(genExcel) configure -state normal
    if {[info exists buttons(appDisplay)]} {$buttons(appDisplay) configure -state normal}
    focus $buttons(genExcel)
  
# not found
  } else {
    if {$localName != ""} {errorMsg "File not found: [truncFileName [file nativename $localName]]"}
  }
  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
proc findFile {startDir {recurse 0}} {
  global fileList
  
  set pwd [pwd]
  if {[catch {cd $startDir} err]} {
    errorMsg $err
    return
  }

  set ext ".ifc"

  foreach match [glob -nocomplain -- *] {
    if {[file extension [string tolower $match]] == $ext} {
      lappend fileList [file join $startDir $match]
    }
    if {[info exists ext1]} {
      if {[file extension [string tolower $match]] == $ext1} {
        lappend fileList [file join $startDir $match]
      }
    }
    if {[info exists ext2]} {
      if {[file extension [string tolower $match]] == $ext2} {
        lappend fileList [file join $startDir $match]
      }
    }
    if {[info exists ext3]} {
      if {[file extension [string tolower $match]] == $ext3} {
        lappend fileList [file join $startDir $match]
      }
    }
  }
  if {$recurse} {
    foreach file [glob -nocomplain *] {
      if {[file isdirectory $file]} {findFile [file join $startDir $file] $recurse}
    }
  }
  cd $pwd
}

#-------------------------------------------------------------------------------
proc saveState {} {
  global optionsFile fileDir openFileList opt writeDirType userWriteDir dispCmd dispCmds
  global flag lastXLS lastXLS1 userXLSFile fileDir1 mydocs verite upgrade maxfiles
  global row_limit yrexcel userEntityFile buttons noStyledItem
  global statusFonts upgradeIFCsvr

  if {![info exists buttons]} {return}
  
  if {[catch {
  
    set fileOptions [open $optionsFile w]
    puts $fileOptions "# Options file for the NIST IFC File Analyzer v[getVersion] ([string trim [clock format [clock seconds]]])\n#\n# DO NOT EDIT OR DELETE FROM USER HOME DIRECTORY $mydocs\n# DOING SO WILL CORRUPT OR DELETE THE CURRENT SETTINGS\n#"
    set varlist [list fileDir fileDir1 userWriteDir userEntityFile openFileList dispCmd dispCmds lastXLS lastXLS1 \
                      userXLSFile statusFont maxfiles noStyledItem row_limit upgrade upgradeIFCsvr verite writeDirType yrexcel \
                      flag(CRASH) flag(DISPGUIDE1) flag(FIRSTTIME) flag(THEOREM)]

    foreach var $varlist {
      if {[info exists $var]} {
        set vartmp [set $var]
        if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
          if {$var != "dispCmds" && $var != "openFileList"} {
            regsub -all {\\} $vartmp "/" vartmp
            puts $fileOptions "set $var \"$vartmp\""
          } else {
            regsub -all {\\} $vartmp "/" vartmp
            regsub -all {\[} $vartmp "\\\[" vartmp
            regsub -all {\]} $vartmp "\\\]" vartmp
            for {set i 0} {$i < [llength $vartmp]} {incr i} {
              if {$i == 0} {
                if {[llength $vartmp] > 1} {
                  puts $fileOptions "set $var \"\{[lindex $vartmp $i]\} \\"
                } else {
                  puts $fileOptions "set $var \"\{[lindex $vartmp $i]\}\""
                }
              } elseif {$i == [expr {[llength $vartmp]-1}]} {
                puts $fileOptions "       \{[lindex $vartmp $i]\}\""
              } else {
                puts $fileOptions "       \{[lindex $vartmp $i]\} \\"
              }
            }
          }
        } else {
          if {$vartmp != ""} {
            puts $fileOptions "set $var [set $var]"
          } else {
            puts $fileOptions "set $var \"\""
          }
        }
      }
    }
    
    set winpos "+300+200"
    set wg [winfo geometry .]
    catch {set winpos [string range $wg [string first "+" $wg] end]}
    puts $fileOptions "set winpos \"$winpos\""
    set wingeo [string range $wg 0 [expr {[string first "+" $wg]-1}]]
    puts $fileOptions "set wingeo \"$wingeo\""

    foreach idx [lsort [array names opt]] {
      set var opt($idx)
      set vartmp [set $var]
      if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
        regsub -all {\\} $vartmp "/" vartmp
        puts $fileOptions "set $var \"$vartmp\""
      } else {
        if {$vartmp != ""} {
          puts $fileOptions "set $var [set $var]"
        } else {
          puts $fileOptions "set $var \"\""
        }
      }
    }

    close $fileOptions

  } emsg]} {
    errorMsg "ERROR writing to options file: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------
proc sortEntities {ranrow rancol} {
  global worksheet ifc nsort
  
  set range [$worksheet($ifc) Range [cellRange 4 1] [cellRange $ranrow $rancol]]

  set B3V [[$worksheet($ifc) Range "B3"] Value]
  set C3V [[$worksheet($ifc) Range "C3"] Value]
  
  if {$B3V == "Name" || $C3V == "ProfileName" || $C3V == "LayerSetName" || $C3V == "item_name" || $B3V == "name" || $B3V == "design_part_name"} {
    if {[incr nsort] == 1} {outputMsg " Sorting rows by Name attribute"}

    if {[string range $ifc end-3 end] == "Type"} {
      if {[[$worksheet($ifc) Range "I4"] Value] != ""} {
        set I3V [[$worksheet($ifc) Range "I3"] Value]
        if {$I3V == "PredefinedType"} {
          set I4 [$worksheet($ifc) Range "I4"]
          set sort [$range Sort $I4 [expr 1]]
        }
      }
      if {[[$worksheet($ifc) Range "H4"] Value] != ""} {
        set H3V [[$worksheet($ifc) Range "H3"] Value]
        if {$H3V == "ElementType" || $H3V == "PredefinedType"} {
          set H4 [$worksheet($ifc) Range "H4"]
          set sort [$range Sort $H4 [expr 1]]
        }
      }
    }

    if {[[$worksheet($ifc) Range "G4"] Value] != ""} {
      set G3V [string tolower [[$worksheet($ifc) Range "G3"] Value]]
      if {$G3V == "tag" || $G3V == "longname"} {
        set G4 [$worksheet($ifc) Range "G4"]
        set sort [$range Sort $G4 [expr 1]]
      }
    }

    if {[string first "IfcQuantity" $ifc] == 0 || $ifc == "IfcPropertyBoundedValue"} {
      set E4 [$worksheet($ifc) Range "E4"]
      set sort [$range Sort $E4 [expr 1]]
    }

    set D3V [[$worksheet($ifc) Range "D3"] Value]
    if {$D3V == "ObjectType" || $ifc == "IfcPropertySingleValue"  || $ifc == "IfcPropertyListValue" || \
                                $ifc == "IfcPropertyBoundedValue" || $ifc == "IfcPropertyEnumeratedValue"} {
      set D4 [$worksheet($ifc) Range "D4"]
      set sort [$range Sort $D4 [expr 1]]
    }

    if {$C3V == "Description" || $C3V == "ProfileName" || $C3V == "LayerSetName" || $C3V == "item_name"} {
      set C4 [$worksheet($ifc) Range "C4"]
      set sort [$range Sort $C4 [expr 1]]
    }

    if {[string tolower $B3V] == "name" || $B3V == "design_part_name"} {
      set B4 [$worksheet($ifc) Range "B4"]
      set sort [$range Sort $B4 [expr 1]]
    }
  }
}

#-------------------------------------------------------------------------------
proc displayResult {} {
  global localName dispCmd appName wdir env
  global mytemp
  global File padcmd
  global edmWriteToFile eeWriteToFile
  
  set dispFile $localName
  set idisp [file rootname [file tail $dispCmd]]
  if {[info exists appName]} {if {$appName != ""} {set idisp $appName}}
  outputMsg "Opening IFC file in: $idisp"

# display file
#  (list is programs that CANNOT start up with a file *OR* need specific commands below)
  if {[string first "Conformance"       $idisp] == -1 && \
      [string first "Indent"            $idisp] == -1 && \
      [string first "Default"           $idisp] == -1 && \
      [string first "Express Engine"    $idisp] == -1 && \
      [string first "EDM Model Checker" $idisp] == -1} {

# start up with a file
    if {[catch {
      exec $dispCmd [file nativename $dispFile] &
    } emsg]} {
      errorMsg $emsg
    }

# default viewer associated with file extension
  } elseif {[string first "Default" $idisp] == 0} {
    .tnb select .tnb.status
    if {[catch {
      exec {*}[auto_execok start] "" $dispFile
    } emsg]} {
      errorMsg "No application is associated with IFC files."
      errorMsg " Go to Websites > Free IFC Software  OR  IFC Implementations"        
    }

# indent file
  } elseif {[string first "Indent" $idisp] != -1} {
    .tnb select .tnb.status
    indentFile $dispFile

#-------------------------------------------------------------------------------
# validate file with Express Engine
  } elseif {[string first "Express Engine" $idisp] != -1} {

    .tnb select .tnb.status
    set eefile $dispFile
    outputMsg "Ready to validate:  [truncFileName [file nativename $eefile]] ([expr {[file size $eefile]/1024}] Kb)" blue
    cd [file dirname $eefile]

    set eename [file tail $eefile]
    set eelog  "[file rootname $eename]\_ee.log"
    if {[string tolower [file extension $eename]] == ".ifc"} {set eelog  "[file rootname $eename]\_ifc_ee.log"}
    if {[file exists $eelog]} {file delete $eelog}

    set fschema [getSchema $dispFile]
    if {$fschema == "IFC2X3"} {
      file copy -force [file join $wdir schemas IFC2X3_TC1.lsp] [file join $mytemp IFC2X3_TC1.lsp]
      set schema "\"[file join $mytemp IFC2X3_TC1.lsp]\""
    } elseif {$fschema == "IFC4"} {
      file copy -force [file join $wdir schemas IFC4.lsp] [file join $mytemp IFC4.lsp]
      set schema "\"[file join $mytemp IFC4.lsp]\""
    } else {
      outputMsg "Only IFC2x3 and IFC4 files can be validated with Express Engine." red
      return
    }

    set cmdexp "\"$dispCmd\" --validate -source-schema $schema -input \"$eename\" -input-encoding p21 -messages-file \"$eelog\""
    outputMsg "Running Express Engine"
    outputMsg "Express Engine log file: [truncFileName [file nativename $eelog]]" blue
    if {[catch {eval exec $cmdexp &} err]} {outputMsg " Express Engine error: $err" red}

    if {[string first "TextPad" $padcmd] != -1} {
      outputMsg "Opening log file in editor"
      exec $padcmd $eelog &
    } else {
      outputMsg "Wait until Express Engine has finished and then open the log file"
    }

#-------------------------------------------------------------------------------
# validate file with ST-Developer Conformance Checkers
  } elseif {[string first "Conformance" $idisp] != -1} {
    .tnb select .tnb.status
    set stfile $dispFile
    outputMsg "Ready to validate:  [truncFileName [file nativename $stfile]] ([expr {[file size $stfile]/1024}] Kb)" blue
    cd [file dirname $stfile]

# gui version
    if {[string first "gui" $dispCmd] != -1 && !$eeWriteToFile} {
      if {[catch {exec $dispCmd $stfile &} err]} {outputMsg "Conformance Checker error:\n $err" red}

# non-gui version
    } else {
      set stname [file tail $stfile]
      set stlog  "[file rootname $stname]\_stdev.log"
      if {[string tolower [file extension $stname]] == ".ifc"} {
        set stlog  "[file rootname $stname]\_ifc_stdev.log"
      }
      catch {if {[file exists $stlog]} {file delete -force $stlog}}
      outputMsg "ST-Developer log file: [truncFileName [file nativename $stlog]]" blue

      set c1 [string first "gui" $dispCmd]
      set dispCmd1 $dispCmd
      if {$c1 != -1} {set dispCmd1 [string range $dispCmd 0 $c1-1][string range $dispCmd $c1+3 end]}

      if {[string first "apconform" $dispCmd1] != -1} {
        if {[catch {exec $dispCmd1 -syntax -required -unique -bounds -aggruni -arrnotopt -inverse -strwidth -binwidth -realprec -atttypes -refdom $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      } else {
        if {[catch {exec $dispCmd1 $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      }  
      if {[string first "TextPad" $padcmd] != -1} {
        outputMsg "Opening log file in editor"
        exec $padcmd $stlog &
      } else {
        outputMsg "Wait until the Conformance Checker has finished and then open the log file"
      }
    }

#-------------------------------------------------------------------------------
# EDM Model Checker
  } elseif {[string first "EDM Model Checker" $idisp] != -1} {
    set filename $dispFile
    .tnb select .tnb.status
    outputMsg "Ready to validate:  [truncFileName [file nativename $filename]] ([expr {[file size $filename]/1024}] Kb)" blue
    cd [file dirname $filename]

# write script file to open database
    set edmscript "[file rootname $filename]_edm.scr"
    set scriptfile [open $edmscript w]
    set okschema 1

    if {$env(USERDOMAIN) == "NIST"} {
      set edmdir [split [file nativename $dispCmd] [file separator]]
      set i [lsearch $edmdir "bin"]
      set edmdir [join [lrange $edmdir 0 [expr {$i-1}]] [file separator]]
      set edmdbopen "ACCUMULATING_COMMAND_OUTPUT,OPEN_SESSION"
      
      set fsl [string tolower [getSchema $filename]]
      puts $scriptfile "Database>Open([file nativename [file join $edmdir Db]], $fsl, $fsl, \"$edmdbopen\")"
    }

# create a temporary file if certain characters appear in the name, copy original to temporary and process that one
    if {$okschema} {
      set tmpfile 0
      set fileroot [file rootname [file tail $filename]]
      if {[string is integer [string index $fileroot 0]] || \
        [string first " " $fileroot] != -1 || \
        [string first "." $fileroot] != -1 || \
        [string first "+" $fileroot] != -1 || \
        [string first "%" $fileroot] != -1 || \
        [string first "(" $fileroot] != -1 || \
        [string first ")" $fileroot] != -1} {
        if {[string is integer [string index $fileroot 0]]} {set fileroot "a_$fileroot"}
        regsub -all " " $fileroot "_" fileroot
        regsub -all {[\.()]} $fileroot "_" fileroot
        set edmfile [file join [file dirname $filename] $fileroot]
        append edmfile [file extension $filename]
        file copy -force $filename $edmfile
        set tmpfile 1
      } else {
        set edmfile $filename
      }

# validate everything
      set validate "FULL_VALIDATION,OUTPUT_STEPID"

# write script file if not writing output to file, just import model and validate
      set edmimport "ACCUMULATING_COMMAND_OUTPUT,KEEP_STEP_IDENTIFIERS,DELETING_EXISTING_MODEL,LOG_ERRORS_AND_WARNINGS_ONLY"
      if {$edmWriteToFile == 0} {
        puts $scriptfile "Data>ImportModel(DataRepository, $fileroot, DataRepository, $fileroot\_HeaderModel, \"[file nativename $edmfile]\", \$, \$, \$, \"$edmimport,LOG_TO_STDOUT\")"
        puts $scriptfile "Data>Validate>Model(DataRepository, $fileroot, \$, \$, \$, \"ACCUMULATING_COMMAND_OUTPUT,$validate,FULL_OUTPUT\")"

# write script file if writing output to file, create file names, import model, validate, and exit
      } else {
        if {[file extension $filename] != ".ifc"} {
          set edmlog  "[file rootname $filename]_edm.log"
          set edmloginput "[file rootname $filename]_edm_input.log"
        } else {
          set edmlog  "[file rootname $filename]_ifc_edm.log"
          set edmloginput "[file rootname $filename]_ifc_edm_input.log"
        }
        puts $scriptfile "Data>ImportModel(DataRepository, $fileroot, DataRepository, $fileroot\_HeaderModel, \"[file nativename $edmfile]\", \"[file nativename $edmloginput]\", \$, \$, \"$edmimport,LOG_TO_FILE\")"
        puts $scriptfile "Data>Validate>Model(DataRepository, $fileroot, \$, \"[file nativename $edmlog]\", \$, \"ACCUMULATING_COMMAND_OUTPUT,$validate,FULL_OUTPUT\")"
        puts $scriptfile "Data>Close>Model(DataRepository, $fileroot, \" ACCUMULATING_COMMAND_OUTPUT\")"
        puts $scriptfile "Data>Delete>ModelContents(DataRepository, $fileroot, ACCUMULATING_COMMAND_OUTPUT)"
        puts $scriptfile "Data>Delete>Model(DataRepository, $fileroot, header_section_schema, \"ACCUMULATING_COMMAND_OUTPUT,DELETE_ALL_MODELS_OF_SCHEMA\")"
        puts $scriptfile "Data>Delete>Model(DataRepository, $fileroot, \$, ACCUMULATING_COMMAND_OUTPUT)"
        puts $scriptfile "Data>Delete>Model(DataRepository, $fileroot, \$, \"ACCUMULATING_COMMAND_OUTPUT,CLOSE_MODEL_BEFORE_DELETION\")"
        puts $scriptfile "Exit>Exit()"
      }
      close $scriptfile

# run EDM Model Checker with the script file
      outputMsg "Running EDM Model Checker"
      if {$env(USERDOMAIN) == "NIST"} {
        eval exec {$dispCmd} $edmscript &

# if results are written to a file, open output file from the validation (edmlog) and output file if there are input errors (edmloginput)
        if {$edmWriteToFile} {
          if {[string first "TextPad" $padcmd] != -1} {
            outputMsg "Opening log file in editor"
            exec $padcmd $edmlog &
            after 1000
            if {[file size $edmloginput] > 0} {
              exec $padcmd $edmloginput &
            } else {
              catch {file delete -force $edmloginput}
            }
          } else {
            outputMsg "Wait until the EDM Model Checker has finished and then open the log file"
          }
        }

# attempt to delete the script file
        set nerr 0
        while {[file exists $edmscript]} {
          after 1000
          incr nerr
          catch {file delete $edmscript}
          if {$nerr > 60} {break}
        }

# if using a temporary file, attempt to delete it
        if {$tmpfile} {
          set nerr 0
          while {[file exists $edmfile]} {
            after 1000
            incr nerr
            catch {file delete $edmfile}
            if {$nerr > 60} {break}
          }
        }
      } else {
        outputMsg "In EDM Model Checker, open a database, then manually input the script file with"
        outputMsg " Aux > Command Script > Run > Select > [truncFileName [file nativename $edmscript]]"
        exec $dispCmd &
      }
    }

# all others
  } else {
    outputMsg "You have to manually import the IFC file to $idisp." red
    .tnb select .tnb.status
    exec $dispCmd &
  }
  
    
# add file to menu
  addFileToMenu
  saveState
}

#-------------------------------------------------------------------------------
proc getDisplayPrograms {} {
  global dispApps dispCmds dispCmd appNames appName env pf32 pf64
  global mydesk drive padcmd

  regsub {\\} $pf32 "/" p32
  lappend pflist $p32
  if {$pf64 != "" && $pf64 != $pf32} {
    regsub {\\} $pf64 "/" p64
    lappend pflist $p64
  }

  set lastver 0
  set ok 0
  if {$env(USERDOMAIN) == "NIST"} {
    set edms [glob -nocomplain -directory [file join $drive edm] -join edm* bin Edms.exe]
    foreach match $edms {
      set name "EDM Model Checker"
      if {[string first "edm5" $match] != -1} {
        set num 5 
      } elseif {[string first "edmsix" $match] != -1} {
        set num 6
      }
      set dispApps($match) "$name $num"
    }
  }
  
#-------------------------------------------------------------------------------
# Solibri Model Checker
  foreach pf $pflist {
    foreach match [glob -nocomplain -directory [file join $pf Solibri] -join * bin SMC.exe] {
      if {![info exists dispApps($match)]} {
        set dispApps($match) "Solibri Model Checker"
      }
    }
  }
  foreach pf $pflist {
    foreach match [glob -nocomplain -directory [file join $pf Solibri] -join * "Solibri Model Checker v*.exe"] {
      if {![info exists dispApps($match)]} {
        set dispApps($match) "Solibri Model Checker"
      }
    }
  }
  foreach pf $pflist {
    foreach match [glob -nocomplain -directory [file join $pf Solibri] -join * "Solibri Model Viewer v*.exe"] {
      if {![info exists dispApps($match)]} {
        set dispApps($match) "Solibri Model Viewer"
      }
    }
  }

# simplebim
  foreach pf $pflist {
    foreach match [glob -nocomplain -directory [file join $pf Datacubist] -join * "simplebim*.exe"] {
      if {![info exists dispApps($match)] && [string first "-" $match] == -1} {
        set dispApps($match) "simplebim [string range $match end-4 end-4]"
      }
    }
  }

# IFC viewers
  foreach pf $pflist {
    set applist [list \
      [list [file join $pf "Tekla BIMsight" BIMsight.exe] "Tekla BIMsight"] \
      [list [file join $pf IFCBrowser IfcQuickBrowser.exe] IfcQuickBrowser] \
      [list [file join $pf Kisters 3DViewStation 3DViewStation.exe] 3DViewStation] \
      [list [file join $pf "Data Design System" Viewer Exe DdsViewer.exe] "DDS-CAD Viewer"] \
      [list [file join $pf DDS Viewer Exe DdsViewer.exe] "DDS-CAD Viewer"] \
      [list [file join $pf "Tekla BIMsight" BIMsight.exe] "Tekla BIMsight"] \
      [list [file join $pf Constructivity "Constructivity One" bin Constructivity.exe] "Constructivity Viewer"] \
      [list [file join $pf Constructivity "Constructivity Model Viewer" bin Constructivity.exe] "Constructivity Model Viewer"] \
      [list [file join $pf Constructivity Constructivity bin Constructivity.exe] "Constructivity Model Viewer"] \
      [list [file join $pf Datacomp "BIM Vision" bim_vision_x64.exe] "BIM Vision"] \
      [list [file join $pf KUBUS "BIMcollab ZOOM" "BIMcollab ZOOM.exe"] "BIMcollab ZOOM"] \
      [list [file join $pf "BIM VILLAGE" Beaver Beaver.exe] "BIM Beaver"] \
      [list [file join $pf Areddo Areddo.exe] Areddo] \
      [list [file join $pf Bentley "Bentley View CONNECT Edition" BentleyView BentleyView.exe] BentleyView] \
      [list [file join $pf "StruMIS Ltd" "BIMReview 8.3" BIMReview.exe] BIMReview] \
      [list [file join $pf "CAD Assistant" CADAssistant.exe] "CAD Assistant"] \
    ]
    foreach app $applist {
      if {[file exists [lindex $app 0]]} {
        set name [lindex $app 1]
        set dispApps([lindex $app 0]) $name
      }
    }
  }
  if {[file exists [file join $drive ACCA usBIM.viewer+ usBIM.viewer.exe]]} {
    set name "usBIM.viewer"
    set dispApps([file join $drive ACCA usBIM.viewer+ usBIM.viewer.exe]) $name
  }

  foreach app {FZKViewer "IFC Viewer" IfcQuery} {
    foreach scut [list "Shortcut to $app.exe.lnk" "$app.exe - Shortcut.lnk" "$app.exe.lnk" "$app.lnk"] {
      catch {
        set f1 [file join $mydesk $scut]
        set f2 [file join $drive "Users" "All Users" Desktop $scut]
        set f3 [file join $drive "Users" "Public" "Public Desktop" $scut]
        foreach f [list $f1 $f2 $f3] {
          if {[file exists $f]} {
            set sc [get_shortcut_filename $f]
            if {[string first "javaws" $sc] == -1} {set dispApps($sc) $app}
            break
          }
        }
      }
    }
  }

#-------------------------------------------------------------------------------
  foreach pf $pflist {

# IfcQuickBrowser
    if {[file exists [file join $pf IFCBrowser IfcQuickBrowser.exe]]} {
      set name "IfcQuickBrowser"
      set dispApps([file join $pf IFCBrowser IfcQuickBrowser.exe]) $name
    }

# ST-Developer STEP File Browser and generic Conformance Checker
    set stmatch ""
    foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin stepbrws.exe] {
      if {$stmatch == ""} {
        set stmatch $match
        set lastver [lindex [split [file nativename $match] [file separator]] 3]
      } else {
        set ver [lindex [split [file nativename $match] [file separator]] 3]
        if {$ver > $lastver} {set stmatch $match}
      }
    }
    if {$stmatch != ""} {
      if {![info exists dispApps($stmatch)]} {
        set vn [lindex [lindex [split [file nativename $stmatch] [file separator]] 3] 1]
        set dispApps($stmatch) "STEP File Browser"
      }
    }
    set stmatch ""
    foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin ifcview.exe] {
      if {$stmatch == ""} {
        set stmatch $match
        set lastver [lindex [split [file nativename $match] [file separator]] 3]
      } else {
        set ver [lindex [split [file nativename $match] [file separator]] 3]
        if {$ver > $lastver} {set stmatch $match}
      }
    }
    if {$stmatch != ""} {
      if {![info exists dispApps($stmatch)]} {
        set vn [lindex [lindex [split [file nativename $stmatch] [file separator]] 3] 1]
        set dispApps($stmatch) "IFC Geometry Viewer"
      }
    }

    set stmatch ""
    foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin apconformgui.exe] {
      if {$stmatch == ""} {
        set stmatch $match
        set lastver [lindex [split [file nativename $match] [file separator]] 3]
      } else {
        set ver [lindex [split [file nativename $match] [file separator]] 3]
        if {$ver > $lastver} {set stmatch $match}
      }
    }
    if {$stmatch != ""} {
      if {![info exists dispApps($stmatch)]} {
        set vn [lindex [lindex [split [file nativename $stmatch] [file separator]] 3] 1]
        set dispApps($stmatch) "AP Conformance Checker"
      }
    }

# ST-Developer IFC Check and Browse
    set stmatch ""
    foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin ifccheckgui.exe] {
      if {$stmatch == ""} {
        set stmatch $match
        set lastver [lindex [split [file nativename $match] [file separator]] 3]
      } else {
        set ver [lindex [split [file nativename $match] [file separator]] 3]
        if {$ver > $lastver} {set stmatch $match}
      }
    }
    if {$stmatch != ""} {
      if {![info exists dispApps($stmatch)]} {
        set vn [lindex [lindex [split [file nativename $stmatch] [file separator]] 3] 1]
        set dispApps($stmatch) "IFC Check and Browse"
      }
    }
  }

# Adobe Acrobat Pro for IFC
  foreach pf $pflist {

# Tetra4D in Adobe Acrobat
    for {set i 12} {$i > 9} {incr i -1} {
      foreach match [glob -nocomplain -directory $pf -join Adobe "Acrobat $i.0" Acrobat Acrobat.exe] {
        if {[file exists [file join $pf Adobe "Acrobat $i.0" Acrobat plug_ins 3DPDFConverter 3DPDFConverter.exe]]} {
          if {![info exists dispApps($match)]} {
            set name "Tetra4D Converter"
            set dispApps($match) $name
          }
        }
      }
      set match [file join $pf Adobe "Acrobat $i.0" Acrobat plug_ins 3DPDFConverter 3DReviewer.exe]
      if {![info exists dispApps($match)]} {
        set name "Tetra4D Reviewer"
        set dispApps($match) $name
      }
    }
    for {set i 2030} {$i > 2012} {incr i -1} {
      foreach match [glob -nocomplain -directory $pf -join Adobe "Acrobat $i" Acrobat Acrobat.exe] {
        if {[file exists [file join $pf Adobe "Acrobat $i" Acrobat plug_ins 3DPDFConverter 3DPDFConverter.exe]]} {
          if {![info exists dispApps($match)]} {
            set name "Tetra4D Converter"
            set dispApps($match) $name
          }
        }
      }
      set match [file join $pf Adobe "Acrobat $i" Acrobat plug_ins 3DPDFConverter 3DReviewer.exe]
      if {![info exists dispApps($match)]} {
        set name "Tetra4D Reviewer"
        set dispApps($match) $name
      }
    }
  }
  
# Navisworks
  foreach pf $pflist {
    foreach match [glob -nocomplain -directory [file join $pf Autodesk] -join "Navisworks Manage*" roamer.exe] {
      if {![info exists dispApps($match)]} {
        set dispApps($match) "Navisworks Manage"
      }
    }
  }

#-------------------------------------------------------------------------------
# file indenter
  set dispApps(Indent) "Indent IFC File (for debugging)"

# default viewer
  set dispApps(Default) "Default IFC Viewer"

#-------------------------------------------------------------------------------
# set text editor command and name
  set padcmd ""
  set padnam ""

# Notepad++ or Notepad
  set padcmd [file join $pf32 Notepad++ notepad++.exe]
  if {[file exists $padcmd]} {
    set padnam "Notepad++"
    set dispApps($padcmd) $padnam
  } elseif {[info exists env(windir)]} {
    set padcmd [file join $env(windir) system32 Notepad.exe]
    set padnam "Notepad"
    set dispApps($padcmd) $padnam
  }

#-------------------------------------------------------------------------------
# remove cmd that do not exist in dispCmds and non-executables
  set dispCmds1 {}
  foreach app $dispCmds {
    set fext [file extension $app]
    if {([file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0) && \
        [file tail $app] != "NotePad.exe" && [string first "Analyzer.exe" $app] == -1 && \
        $fext != ".wrl" && $fext != ".ifc" && $fext != ".stp" && \
        $fext != ".step" && $fext != ".p21" && $fext != ".stpnc" && $fext != ".jpg"} {
      lappend dispCmds1 $app
    }
  }
  set dispCmds $dispCmds1
  set fext [file extension $dispCmd]
  if {$fext == ".wrl" || $fext == ".ifc" || $fext == ".stp" || $fext == ".step" || \
      $fext == ".p21" || $fext == ".jpg"} {set dispCmd ""}

# check for cmd in dispApps that does not exist in dispCmds and add to list
  foreach app [array names dispApps] {
    if {[file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0} {
      set notInCmds 1
      foreach cmd $dispCmds {if {[string tolower $cmd] == [string tolower $app]} {set notInCmds 0}}
      if {$notInCmds} {lappend dispCmds $app}
    }
  }

# remove duplicates in dispCmds
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# clean up list of app viewer commands
  if {[info exists dispCmd]} {
    if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0) && [string first "Analyzer.exe" $app] == -1} {
      if {[lsearch $dispCmds $dispCmd] == -1 && $dispCmd != ""} {lappend dispCmds $dispCmd}
    } else {
      if {[llength $dispCmds] > 0} {
        foreach dispCmd $dispCmds {if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0)} {break}}
      } else {
        set dispCmd ""
      }
    }
  } else {
    if {[llength $dispCmds] > 0} {
      set dispCmd [lindex $dispCmds 0]
    }
  }
  for {set i 0} {$i < [llength $dispCmds]} {incr i} {
    if {![file exists [lindex $dispCmds $i]] && [string first "Default" [lindex $dispCmds $i]] == -1 && [string first "Indent" [lindex $dispCmds $i]] == -1} {set dispCmds [lreplace $dispCmds $i $i]}
  }

# put dispCmd at beginning of dispCmds list
  if {[info exists dispCmd]} {
    for {set i 0} {$i < [llength $dispCmds]} {incr i} {
      if {$dispCmd == [lindex $dispCmds $i]} {
        set dispCmds [lreplace $dispCmds $i $i]
        set dispCmds [linsert $dispCmds 0 $dispCmd]
      }
    }
  }

# remove duplicates in dispCmds, again
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# set list of IFC viewer names, appNames
  set appNames {}
  set appName  ""
  foreach cmd $dispCmds {
    if {[info exists dispApps($cmd)]} {
      lappend appNames $dispApps($cmd)
    } else {
      set name [file rootname [file tail $cmd]]
      if {$name == "Edms"} {set name "EDM Model Checker"}

      lappend appNames  $name
      set dispApps($cmd) $name
    }
  }
  if {$dispCmd != ""} {
    if {[info exists dispApps($dispCmd)]} {set appName $dispApps($dispCmd)}
  }
  if {[llength $appNames] > 0 && ![info exists ifc(View)]} {set ifc(View) 1}
}

#-------------------------------------------------------------------------------
proc addFileToMenu {} {
  global openFileList localName File filemenuinc lenlist buttons
  
  if {![info exists buttons]} {return}
  
# change backslash to forward slash, if necessary
  regsub -all {\\} $localName "/" localName

# remove duplicates
  set newlist {}
  set dellist {}
  for {set i 0} {$i < [llength $openFileList]} {incr i} {
    set name [lindex $openFileList $i]
    set ifile [lsearch -all $openFileList $name]
    if {[llength $ifile] == 1 || [lindex $ifile 0] == $i} {
      lappend newlist $name
    } else {
      lappend dellist $i
    }
  }
  set openFileList $newlist
  #foreach idx $dellist {
  #  $File delete [expr {$idx+$filemenuinc}] [expr {$idx+$filemenuinc}]
  #}
  
# check if file name is already in the menu, if so, delete
  set ifile [lsearch $openFileList $localName]
  if {$ifile > 0} {
    set openFileList [lreplace $openFileList $ifile $ifile]
    $File delete [expr {$ifile+$filemenuinc}] [expr {$ifile+$filemenuinc}]
  }

# insert file name at top of list
  set fext [string tolower [file extension $localName]]
  if {$ifile != 0 && ($fext == ".stp" || $fext == ".step" || $fext == ".p21" || $fext == ".ifc")} {
    set openFileList [linsert $openFileList 0 $localName]
    $File insert $filemenuinc command -label [truncFileName [file nativename $localName] 1] \
      -command [list openFile $localName] -accelerator "F1"
    catch {$File entryconfigure 5 -accelerator {}}
  }

# check length of file list, delete from the end of the list
  if {[llength $openFileList] > $lenlist} {
    set openFileList [lreplace $openFileList $lenlist $lenlist]
    $File delete [expr {$lenlist+$filemenuinc}] [expr {$lenlist+$filemenuinc}]
  }

# compare file list and menu list
  set llen [llength $openFileList]
  for {set i 0} {$i < $llen} {incr i} {
    set f1 [file tail [lindex $openFileList $i]]
    set f2 ""
    catch {set f2 [file tail [lindex [$File entryconfigure [expr {$i+$filemenuinc}] -label] 4]]}
    #if {$f1 != $f2 && $f2 != ""} {errorMsg "File list and menu out of synch: $i $f1 $f2"}
  }
  
# save the state so that if the program crashes the file list will be already saved
  saveState
  return
}

#-------------------------------------------------------------------------------
# open a spreadsheet
proc openXLS {filename {check 0} {multiFile 0}} {
  global buttons

  if {[info exists buttons]} {.tnb select .tnb.status}

  if {[file exists $filename]} {

# check if instances of Excel are already running
    if {$check} {checkForExcel}
    outputMsg " "
    
# start Excel
    if {[catch {
      #outputMsg "Starting Excel" green
      ::tcom::ref createobject Excel.Application

# errors
    } emsg]} {
      errorMsg "ERROR starting Excel: $emsg"
    }
    
# open spreadsheet in Excel, works even if Excel not already started above although slower
    if {[catch {
      outputMsg "Opening Spreadsheet: [file tail $filename]"
      exec {*}[auto_execok start] "" $filename

# errors
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {
        if {[string first "The process cannot access the file" $emsg] != -1} {
          outputMsg " The Spreadsheet might already be opened." red
        } else {
          outputMsg " Error opening the Spreadsheet: $emsg" red
        }
        catch {raise .}
      }
    }

  } else {
    if {[file tail $filename] != ""} {errorMsg "Spreadsheet not found: [truncFileName [file nativename $filename]]"}
    set filename ""
  }
  return $filename
}

#-------------------------------------------------------------------------------
proc checkForExcel {{multFile 0}} {
  global lastXLS localName buttons
  
  set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
  if {[llength $pid1] > 0} {
    if {[info exists buttons]} {
      if {!$multFile} {
        set msg "There are at least ([llength $pid1]) Excel spreadsheets already opened.\n\nDo you want to close the open spreadsheets?"
        set dflt yes
        if {[info exists lastXLS] && [info exists localName]} {
          if {[llength $pid1] == 1} {if {[string first [file nativename [file rootname $localName]] [file nativename $lastXLS]] != 0} {set dflt no}}
        }
        set choice [tk_messageBox -type yesno -default $dflt -message $msg -icon question -title "Close Spreadsheets?"]
        if {$choice == "yes"} {
          #outputMsg "Closing Excel" red
          for {set i 0} {$i < 5} {incr i} {
            set nnc 0
            foreach pid $pid1 {
              if {[catch {
                twapi::end_process $pid -force
              } emsg]} {
                incr nnc
              }
            }
            set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
            if {[llength $pid1] == 0} {break}
          }
          #if {$nnc > 0} {errorMsg " Some instances ($nnc) of Excel were not closed.  $emsg" red}
        }
      }
    } else {
      #outputMsg "Closing Excel" red
      foreach pid $pid1 {
        if {[catch {
          twapi::end_process $pid -force
        } emsg]} {
          errorMsg " Some instances of Excel were not closed.  $emsg" red
        }
      }
    }
  }
  return $pid1
}

#-------------------------------------------------------------------------------

proc cellRange {r c} {
  set letters ABCDEFGHIJKLMNOPQRSTUVWXYZ
  
# correct if 'c' is passed in as a letter
  set cf [string first $c $letters]
  if {$cf != -1} {set c [expr {$cf+1}]}

# a much more elegant solution from the Tcl wiki
  set cr ""
  set n $c
  while {[incr n -1] >= 0} {
    set cr [format %c%s [expr {$n%26+65}] $cr]
    set n [expr {$n/26}]
  }

  if {$r > 0} {
    append cr $r
  } else {
    append cr ":$cr"
  }
  
  return $cr
}

#-------------------------------------------------------------------------------

proc colorBadCells {ent} {
  global excel syntaxerr count cells worksheet
      
# color "Bad" (red) for syntax errors
  if {[expr {int([$excel Version])}] >= 12} {
    if {[info exists syntaxerr($ent)]} {
      for {set n 0} {$n < [llength $syntaxerr($ent)]} {incr n} {
        if {[catch {
          set err [lindex $syntaxerr($ent) $n]

# get row and column number
          set r [lindex $err 0]
          set c [lindex $err 1]
          

# values are entity ID (row) and attribute name (column)
          #outputMsg "$ent / $r / $c / [string is integer $c]"
          if {![string is integer $c]} {
            for {set i 2} {$i < 100} {incr i} {
              set val [[$cells($ent) Item 3 $i] Value]
              if {$val == $c} {
                set c $i
                break
              }
            }
            set c1 [expr {$count($ent)+3}]
            for {set i 4} {$i <= $c1} {incr i} {
              set val [[$cells($ent) Item $i 1] Value]
              if {$val == $r} {
                set r $i
                break
              }              
            }
          }
          [$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Style "Bad"
        } emsg]} {
          errorMsg "ERROR setting spreadsheet cell color: $emsg\n  $ent"
          catch {raise .}
        }
      }
    }
  }
}

#-------------------------------------------------------------------------------
# trimNum gets used mostly when processing IFC files

proc trimNum {num {prec 3} {checkcomma 0}} {
  global unq_num comma
  
  set numsav $num
  if {[info exists unq_num($numsav)]} {
    set num $unq_num($numsav)
  } else {
    if {[catch {
      set form "\%."
      append form $prec
      append form "f"
      set num [format $form $num]

      if {[string first "." $num] != -1} {
        for {set i 0} {$i < $prec} {incr i} {
          set num [string trimright $num "0"]
        }
        if {$num == "-0"} {set num 0.}
      }
    } errmsg]} {
      errorMsg "# $errmsg ($numsav reset to 0.0)" red
      set num 0.
    }
    if {$checkcomma && $comma} {regsub -all {\.} $num "," num}
    set unq_num($numsav) $num
  }
  return $num
}

#-------------------------------------------------------------------------------

proc outputMsg {msg {color "black"}} {
  global outputWin

  if {[info exists outputWin]} {
    $outputWin issue "$msg " $color
    update 
  } else {
    puts $msg
  }
}

#-------------------------------------------------------------------------------
proc errorMsg {msg {color ""}} {
  global errmsg outputWin

  if {![info exists errmsg]} {set errmsg ""}
  
  if {[string first $msg $errmsg] == -1} {
    set errmsg "$msg\n$errmsg"
    
# this fix is necessary to handle messages related to inverses
    set c1 [string first "DELETETHIS" $msg]
    if {$c1 != -1} {set msg [string range $msg 0 $c1-1]}
    
    puts $msg
    if {[info exists outputWin]} {
      if {$color == ""} {
        if {[string first "syntax error" [string tolower $msg]] != -1} {
          $outputWin issue "$msg " syntax
        } else {
          set ilevel ""
          catch {set ilevel "  \[[lindex [info level [expr {[info level]-1}]] 0]\]"}
          if {$ilevel == "  \[errorMsg\]"} {set ilevel ""}
          $outputWin issue "$msg$ilevel " error
        }
      } else {
        $outputWin issue "$msg " $color
      }
      update 
    }
    return 1
  } else {
    return 0
  }
}

# -------------------------------------------------------------------------------------------------
proc fixErrorMsg {emsg} {
  set emsg [split $emsg "\n"]
  if {[llength $emsg] > 3} {
    set emsg [join [lrange $emsg 3 end] "\n"]
  } else {
    set emsg ""
  }
  return $emsg
}

# -------------------------------------------------------------------------------------------------
proc truncFileName {fname {compact 0}} {
  global mydocs myhome mydesk

  if {[string first $mydocs $fname] == 0} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydocs] end]"
  } elseif {[string first $mydesk $fname] == 0 && $mydesk != $fname} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydesk] end]"
  #} elseif {[string first $myhome $fname] == 0 && $myhome != $fname} {
  #  set nname "[string range $fname 0 2]...[string range $fname [string length $myhome] end]"
  }

  if {[info exists nname]} {
    if {$nname != "C:\\..."} {set fname $nname}
  }

  if {$compact} {
    catch {
      while {[string length $fname] > 80} {
        set nname $fname
        set s2 0
        if {[string first "\\\\" $nname] == 0} {
          set nname [string range $nname 2 end]
          set s2 1
        }

        set nname [file nativename $nname]
        set sname [split $nname [file separator]]
        if {[llength $sname] <= 3} {break}

        if {[lindex $sname 1] == "..."} {
          set sname [lreplace $sname 2 2]
        } else {
          set sname [lreplace $sname 1 1 "..."]
        }

        set nname ""
        set nitem 0
        foreach item $sname {
          if {$nitem == 0 && [string length $item] == 2 && [string index $item 1] == ":"} {append item "/"}
          set nname [file join $nname $item]
          incr nitem
        }
        if {$s2} {set nname \\\\$nname}
        set fname [file nativename $nname]
      }
    }
  }
  return $fname
}

#-------------------------------------------------------------------------------
# create new file name if current file already exists
proc incrFileName {fn} {
  set fext [file extension $fn]
  set c1 [string last "." $fn]
  for {set i 1} {$i < 100} {incr i} {
    set fn "[string range $fn 0 $c1-1] ($i)$fext"
    catch {[file delete -force -- $fn]}
    if {![file exists $fn]} {break}
  }
  return $fn
}

#-------------------------------------------------------------------------------
# install IFCsvr (or uninstall to reinstall)
proc installIFCsvr {{reinstall 0}} {
  global buttons mydocs mytemp nistVersion upgradeIFCsvr wdir

  set ifcsvr     "ifcsvrr300_setup_1008_en-update.msi"
  set ifcsvrInst [file join $wdir exe $ifcsvr]

# install if not already installed
  outputMsg " "
  
# first time installation
  if {!$reinstall} {
    errorMsg "The IFCsvr toolkit must be installed to read and process IFC files."
    outputMsg "- You might need administrator privileges (Run as administrator) to install the toolkit.
  Antivirus software might respond that there is a security issue with the toolkit.  The
  toolkit is safe to install.  Use the default installation folder for the toolkit.
- To reinstall the toolkit, run the installation file ifcsvrr300_setup_1008_en-update.msi
  in $mytemp  or your home directory or the current directory.
- If there are problems with this procedure, email the Contact in Help > About."

    if {[file exists $ifcsvrInst]} {
      set msg "The IFCsvr toolkit must be installed to read and process IFC files.  After clicking OK the IFCsvr toolkit installation will start."
      append msg "\n\nYou might need administrator privileges (Run as administrator) to install the toolkit.  Antivirus software might respond that there is a security issue with the toolkit.  The toolkit is safe to install.  Use the default installation folder for the toolkit."
      append msg "\n\nIf there are problems with this procedure, email the Contact in Help > About."
      set choice [tk_messageBox -type ok -message $msg -icon info -title "Install IFCsvr"]
      outputMsg "\nWait for the installation to finish before processing an IFC file." red
    } elseif {![info exists buttons]} {
      outputMsg "\nRerun this program after the installation has finished to process an IFC file."
    }

# reinstall
  } else {
    errorMsg "The existing IFCsvr toolkit must be reinstalled to update the toolkit."
    outputMsg "- First REMOVE the current installation of the IFCsvr toolkit."
    outputMsg "    In the IFCsvr Setup Wizard select 'REMOVE IFCsvrR300 ActiveX Component' and Finish" red
    outputMsg "    If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
    if {[info exists buttons]} {
      outputMsg "- Then restart this software or process an IFC file to install the updated IFCsvr toolkit."
    } else {
      outputMsg "- Then run this software again to install the updated IFCsvr toolkit."
    }
    outputMsg "- If there are problems with this procedure, email the Contact in Help > About."

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be reinstalled to update the toolkit."
      append msg "\n\nFirst REMOVE the current installation of the IFCsvr toolkit."
      append msg "\n\nIn the IFCsvr Setup Wizard (after clicking OK below) select 'REMOVE IFCsvrR300 ActiveX Component' and Finish.  If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
      append msg "\n\nThen restart this software or process a STEP file to install the updated IFCsvr toolkit."
      append msg "\n\nIf there are problems with this procedure, email the Contact in Help > About."
      set choice [tk_messageBox -type ok -message $msg -icon warning -title "Reinstall IFCsvr"]
      outputMsg "\nWait for the REMOVE process to finish, then restart this software or process an IFC file to install the updated IFCsvr toolkit." red
    }
  }

# try copying installation file to several locations
  set ifcsvrMsi [file join $mytemp $ifcsvr]
  if {[file exists $ifcsvrInst]} {
    if {[catch {
      file copy -force -- $ifcsvrInst $ifcsvrMsi
    } emsg1]} {
      set ifcsvrMsi [file join $mydocs $ifcsvr]
      if {[catch {
        file copy -force -- $ifcsvrInst $ifcsvrMsi
      } emsg2]} {
        set ifcsvrMsi [file join [pwd] $ifcsvr]
        if {[catch {
          file copy -force -- $ifcsvrInst $ifcsvrMsi
        } emsg3]} {
          errorMsg "ERROR copying the IFCsvr toolkit installation file to a directory."
          outputMsg " $emsg1\n $emsg2\n $emsg3"
        }
      }
    }
  }
  
# delete old installer
  catch {file delete -force -- [file join $mytemp ifcsvrr300_setup_1008_en.msi]}

# ready or not to install
  if {[file exists $ifcsvrMsi]} {
    if {[catch {
      exec {*}[auto_execok start] "" $ifcsvrMsi
      if {!$reinstall} {
        set upgradeIFCsvr [clock seconds]
        saveState
      }
    } emsg]} {
      errorMsg "ERROR installing IFCsvr toolkit: $emsg"
    }
  } else {
    if {[file exists $ifcsvrInst]} {errorMsg "The IFCsvr toolkit cannot be automatically installed."}
    catch {.tnb select .tnb.status}
    update idletasks
    outputMsg "To manually install the IFCsvr toolkit:
- The installation file ifcsvrr300_setup_1008_en-update.msi can be found in either:
  $mytemp or your home directory or the current directory.
- Run the installer and follow the instructions.  Use the default installation folder for IFCsvr.
  You might need administrator privileges (Run as administrator) to install the toolkit.
- If there are problems with the IFCsvr installation, email the Contact in Help > About\n"
    after 1000
    errorMsg "Opening folder: $mytemp"
    if {[catch {
      exec {*}[auto_execok start] [file nativename $mytemp]
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {errorMsg "ERROR opening directory: $emsg"}
    }
  }
}

#-------------------------------------------------------------------------------
# get next unused column
proc getNextUnusedColumn {ent r} {
  global cells
  
  for {set c 30} {$c > 1} {incr c -1} {
    set val [[$cells($ent) Item $r $c] Value]
    if {$val != ""} {
      set nextcol [expr {$c+1}]
      #outputMsg "getNextUnusedColumn $val $nextcol" red
      return $nextcol
    }
  }
}

# -------------------------------------------------------------------------------

proc formatComplexEnt {str {space 0}} {
  set str1 $str
  catch {
    if {[string first "_and_" $str1] != -1} {
      regsub -all "_and_" $str1 ") (" str1
      if {$space == 0} {regsub -all " " $str1 "" str1}
      set str1 "($str1)"
    }
  }
  return $str1
}

# -------------------------------------------------------------------------------
proc get_shortcut_filename {file} {
  set dir [file nativename [file dirname $file]]
  set tail [file nativename [file tail $file]]

  if {![string match ".lnk" [string tolower [file extension $file]]]} {
    return -code error "$file is not a valid shortcut name"
  }

  if {[string match "windows" $::tcl_platform(platform)]} {

# Get Shortcut file as an object
    set oShell [tcom::ref createobject "Shell.Application"]
    set oFolder [$oShell NameSpace $dir]
    set oFolderItem [$oFolder ParseName $tail]
    
# If its a shortcut, do modify
    if {[$oFolderItem IsLink]} {
      set oShellLink [$oFolderItem GetLink]
      set path [$oShellLink Path]
      regsub -all {\\} $path "/" path
      return $path
    } else {
      if {![catch {file readlink $file} new]} {
        set new
      } else {
        set file
      }
    }
  } else {
    if {![catch {file readlink $file} new]} {
      set new
    } else {
      set file
    }
  }
}

# -------------------------------------------------------------------------------
proc create_shortcut {file args} {
  if {![string match ".lnk" [string tolower [file extension $file]]]} {
    append file ".lnk"
  }

  if {[string match "windows" $::tcl_platform(platform)]} {
# Make sure filenames are in nativename format.
    array set opts $args
    foreach item [list IconLocation Path WorkingDirectory] {
      if {[info exists opts($item)]} {
        set opts($item) [file nativename $opts($item)]
      }
    }

    set oShell [tcom::ref createobject "WScript.Shell"]
    set oShellLink [$oShell CreateShortcut [file nativename $file]]
    foreach {opt val} [array get opts] {
      if {[catch {$oShellLink $opt $val} result]} {
        return -code error "Invalid shortcut option $opt or value $value: $result"
      }
    }
    $oShellLink Save
    return 1
  }
  return 0
}
