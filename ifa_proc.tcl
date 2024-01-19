#-------------------------------------------------------------------------------
proc getTiming {{str ""}} {
  global tlast

  set t [clock clicks -milliseconds]
  if {[info exists tlast]} {outputMsg "Timing: [expr {($t-$tlast)}]  $str" red}
  set tlast $t
}

#-------------------------------------------------------------------------------
proc getSchema {fname {limit 0}} {
  set schema ""
  set ok 0
  set IFC4 0
  set IFC4XN 0
  set nline 0
  set ifcfile [open $fname r]

  set ulimit 100
  if {$limit} {set ulimit 100000}

  while {[gets $ifcfile line] != -1 && $nline < $ulimit} {
    incr nline
    if {[string first "FILE_SCHEMA" $line] != -1} {
      set ok 1
      set fsline $line
    } elseif {[string first "ENDSEC" $line] != -1} {
      set schema [lindex [split $fsline "'"] 1]
      set fs1 [string toupper $schema]
      if {$fs1 == "IFC4"} {set IFC4 1}
      if {[string first "IFC4X" $fs1] == 0} {set IFC4XN 1}
      if {!$limit} {break}
    } elseif {$schema != ""} {
      if {[string first "\\X2\\" $line] != -1} {
        errorMsg "Unicode in text strings (\\X2\\ encoding) used for symbols and accented or non-English characters are not supported."
      } elseif {[string first "TEXTURE" $line] != -1 && $IFC4XN} {
        set c1 [string first "TEXTURE" $line]
        set c2 [string first "(" $line]
        if {$c1 < $c2} {errorMsg "Entities related to TEXTURE are not supported and might cause a crash.  See Help > IFC Support"}
      } elseif {[string first "IFCCARTESIANPOINTLIST2D" $line] != -1 && $IFC4} {
        errorMsg "Some IFC4 Geometry entities are not supported and might cause a crash.  See Help > IFC Support"
        break
      }
    } elseif {$ok} {
      append fsline $line
    }
  }

  close $ifcfile
  return $schema
}

#-------------------------------------------------------------------------------
proc processToolTip {ttmsg tt} {
  global ifc4only type

  set ttlim 100
  foreach pr {PR_BEAM PR_HVAC PR_RELA PR_PRES} {if {$tt == $pr} {set ttlim 120}}
  if {$tt == "PR_GEOM" || $tt == "PR_IF23"} {set ttlim 130}

  set txt ""
  foreach ifctype {ifc2x3 ifc4} {
    set ttlen 0
    foreach item [lsort $type($tt)] {
      set ok 0
      switch -- $ifctype {
        ifc2x3 {if {[lsearch $ifc4only $item] == -1} {set ok 1}}
        ifc4   {if {[lsearch $ifc4only $item] != -1} {set ok 1}}
      }
      if {$ok} {
        incr ttlen [expr {[string length $item]+3}]
        set ent $item
        if {$ttlen <= $ttlim} {
          append txt "$ent   "
        } else {
          if {[string index $txt end] != "\n"} {set txt "[string range $txt 0 end-3]\n$ent   "}
          set ttlen [expr {[string length $ent]+3}]
        }
      }
    }
    if {$ifctype == "ifc2x3" && $tt != "PR_INFR" && $tt != "PR_REPR" && $tt != "PR_IF23"} {
      append txt "\n\nThe following entities are supported in IFC4 and/or IFC4X3.\n\n"
    }
  }

  append ttmsg $txt
  return $ttmsg
}

#-------------------------------------------------------------------------------
proc checkValues {} {
  global allNone appName appNames buttons opt userentlist writeDirType

  if {[info exists buttons(appCombo)]} {
    set ic [lsearch $appNames $appName]
    if {$ic < 0} {set ic 0}
    $buttons(appCombo) current $ic
  }

  if {$opt(XLSCSV) == "CSV"} {
    set opt(INVERSE) 0
    set opt(EX_ANAL) 0
    set opt(EX_A2P3D) 0
    set opt(EX_LP) 0
    set opt(COUNT) 0
    $buttons(optEX_ANAL) configure -state disabled
    $buttons(optEX_A2P3D) configure -state disabled
    $buttons(optEX_LP) configure -state disabled
    $buttons(optINVERSE) configure -state disabled
    $buttons(optSORT)    configure -state disabled
    $buttons(optXL_FPREC) configure -state disabled
    $buttons(optHIDELINKS) configure -state disabled
    $buttons(genExcel)   configure -text "Generate CSV Files"
  } else {
    $buttons(optEX_ANAL) configure -state normal
    $buttons(optEX_A2P3D) configure -state normal
    $buttons(optEX_LP) configure -state normal
    $buttons(optINVERSE) configure -state normal
    $buttons(optSORT)    configure -state normal
    $buttons(optXL_FPREC) configure -state normal
    $buttons(optHIDELINKS) configure -state normal
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

  if {[info exists buttons(optEX_PROP)]} {
    if {$opt(PR_PROP)} {
      $buttons(optEX_PROP) configure -state normal
    } else {
      $buttons(optEX_PROP) configure -state disabled
      set opt(EX_PROP) 0
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
  } elseif {$writeDirType == 2} {
    $buttons(userentry) configure -state normal
    $buttons(userdir) configure -state normal
  }

# make sure there is some entity type to process
  set nopt 0
  foreach idx [array names opt] {
    if {[string first "PR_" $idx] == 0} {incr nopt $opt($idx)}
  }
  if {$nopt == 0} {
    set opt(PR_BEAM) 1
    set opt(PR_HVAC) 1
    set opt(PR_ELEC) 1
    set opt(PR_INFR) 1
  }

# configure all and none buttons
  if {[info exists allNone]} {
    if {$allNone == 1} {
      foreach item [array names opt] {
        if {[string first "PR_" $item] == 0 && $item != "PR_BEAM"} {
          if {$opt($item) == 1} {set allNone -1; break}
        }
      }
    } elseif {$allNone == 0} {
      foreach item [array names opt] {
        if {[string first "PR_" $item] == 0} {
          if {$item != "PR_USER"} {if {$opt($item) == 0} {set allNone -1; break}}
        }
      }
    }
  }
}

# -------------------------------------------------------------------------------------------------
proc entDocLink {sheet ent r c hlink} {
  global cells fileschema ifcdoc2x3 worksheet

  if {$sheet == "Summary"} {set c 3}

# IFC2x3 doc link
  if {[string first "IFC4" $fileschema] == -1} {
    if {[info exists ifcdoc2x3([string tolower $ent])]} {
      set ent_link "https://standards.buildingsmart.org/IFC/RELEASE/IFC2x3/TC1/HTML/$ifcdoc2x3([string tolower $ent])/lexical/[string tolower $ent].htm"
      set str "IFC2X3"
      $cells($sheet) Item $r $c "Doc"
      set anchor [$worksheet($sheet) Range [cellRange $r $c]]
      if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
      if {[catch {
        $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $str Documentation"]
      } emsg]} {
        errorMsg "Error adding $sheet documentation link: $emsg"
        $cells($sheet) Item $r $c " "
        catch {raise .}
      }
      unset ent_link
    }
  }

# IFC4 doc or deprecated link
  if {[string first "IFC4" $fileschema] != -1} {
    set fs [string toupper [string range $fileschema 0 5]]
    switch -- $fs {
      IFC4X3 {
        set txt1 "IFC4X3"
        set url1 "https://ifc43-docs.standards.buildingsmart.org/IFC/RELEASE/IFC4x3/HTML/lexical/$ent.htm"
      }
      IFC4 {
        set txt1 "IFC4"
        set url1 "https://standards.buildingsmart.org/IFC/RELEASE/IFC4/FINAL/HTML/link/[string tolower $ent].htm"
      }
    }
    set c1 $c
    set ent_link $url1
    $cells($sheet) Item $r $c1 "Doc"
    set anchor [$worksheet($sheet) Range [cellRange $r $c1]]
    if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
    if {[catch {
      $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $txt1 Documentation"]
    } emsg]} {
      errorMsg "Error adding $sheet documentation link: $emsg"
      $cells($sheet) Item $r $c1 " "
      catch {raise .}
    }
    unset ent_link
  }

# if ent_link exists put the link except for IFC which is done above
  if {[info exists ent_link]} {
    $cells($sheet) Item $r $c "Doc"
    set anchor [$worksheet($sheet) Range [cellRange $r $c]]
    if {$sheet == "Summary"} {$anchor HorizontalAlignment [expr -4108]}
    if {[catch {
      $hlink Add $anchor [join $ent_link] [join ""] [join "$ent $str Documentation"]
    } emsg]} {
      errorMsg "Error adding $sheet documentation link: $emsg"
      $cells($sheet) Item $r $c " "
      catch {raise .}
    }
  }
}

# -------------------------------------------------------------------------------------------------
# set color based on tabcolor variable
proc setColorIndex {ifc {stat 0}} {
  global tabcolor type

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
          set tc1 $tabcolor($i)
        }
        if {$c2 == $c1} {
          set ifc2 [string range $ifc $c1+5 end]
          if {[lsearch $type($i) $ifc2] != -1} {
            set tc2 $tabcolor($i)
          }
        } elseif {$c2 != $c1} {
          set ifc2 [string range $ifc $c1+5 $c2-1]
          if {[lsearch $type($i) $ifc2] != -1} {
            set tc2 $tabcolor($i)
          }
          set ifc3 [string range $ifc $c2+5 end]
          if {[lsearch $type($i) $ifc3] != -1} {
            set tc3 $tabcolor($i)
          }
        }
      }
    }
    set tc [expr {min($tc1,$tc2,$tc3)}]

    if {$tc < 1000} {return $tc}
  }
  return -2
}

# -------------------------------------------------------------------------------------------------
proc getFirstFile {} {
  global buttons openFileList padcmd remoteName

  set localName [lindex $openFileList 0]
  if {$localName != ""} {
    set remoteName $localName
    outputMsg "\nReady to process: [file tail $localName] ([fileSize $localName])" blue

    if {[info exists buttons(appDisplay)]} {
      .tnb select .tnb.status
      $buttons(appDisplay) configure -state normal
      if {$padcmd != ""} {
        bind . <Key-F8> {
          if {[file exists $localName]} {
            outputMsg "\nOpening IFC file: [file tail $localName]"
            exec $padcmd [file nativename $localName] &
          }
        }
        bind . <Shift-F8> {
          if {[file exists $localName]} {
            set dir [file nativename [file dirname $localName]]
            outputMsg "\nOpening IFC file directory: [truncFileName $dir]"
            catch {exec C:/Windows/explorer.exe $dir &}
          }
        }
      }
    }
  }
  return $localName
}

#-------------------------------------------------------------------------------
proc displayURL {url} {

# open in whatever is registered for the file extension
  if {[catch {
    exec {*}[auto_execok start] "" $url
  } emsg]} {
    if {[string first "is not recognized" $emsg] == -1} {
      if {[string first "UNC" $emsg] == -1} {errorMsg "Error opening $url: $emsg"}
    }
  }
}

#-------------------------------------------------------------------------------
proc openFile {{openName ""}} {
  global buttons fileDir localName localNameList mytemp padcmd

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
    if {[string first ".ifczip" $lcln] != -1} {
      if {[catch {
        outputMsg "Unzipping: [file tail $localName] ([fileSize $localName])" blue

        vfs::zip::Mount $localName ifczip
        set files [glob -nocomplain ifczip/*]
        set files [join $files]
        set ftmp [string range $files 7 end]
        outputMsg " Extracting: $ftmp"
        set fifc [file join [file dirname $localName] $ftmp]
        set ok 0
        if {![file exists $fifc]} {
          set ok 1
        } elseif {[file mtime $localName] != [file mtime $fifc]} {
          outputMsg " Overwriting existing file: [truncFileName [file nativename $fifc]]" red
          set ok 1
        } else {
          outputMsg " Using existing uncompressed IFC file" red
        }
        if {$ok} {file copy -force -- $files $fifc}
        set localName $fifc
        catch {file delete -force -- [file join $mytemp "gunzip.exe"]}
      } emsg]} {
        errorMsg "Error unzipping file: $emsg"
      }
    }
    set fileDir [file dirname $localName]

    outputMsg "Ready to process: [file tail $localName] ([fileSize $localName])" blue
    if {[file size $localName] > 429000000} {outputMsg " The file might be too large to generate a Spreadsheet." red}

    if {[info exists buttons]} {
      $buttons(genExcel) configure -state normal
      if {[info exists buttons(appDisplay)]} {$buttons(appDisplay) configure -state normal}
      focus $buttons(genExcel)
      if {$padcmd != ""} {
        bind . <Key-F8> {
          if {[file exists $localName]} {
            outputMsg "\nOpening IFC file: [file tail $localName]"
            exec $padcmd [file nativename $localName] &
          }
        }
        bind . <Shift-F8> {
          if {[file exists $localName]} {
            set dir [file nativename [file dirname $localName]]
            outputMsg "\nOpening IFC file directory: [truncFileName $dir]"
            catch {exec C:/Windows/explorer.exe $dir &}
          }
        }
      }
    }

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
  global buttons dispCmd dispCmds fileDir fileDir1 lastXLS lastXLS1 mydocs openFileList opt optionsFile
  global row_limit statusFont upgradeIFCsvr userEntityFile userWriteDir ifaVersion writeDirType

  if {![info exists buttons]} {return}

  if {[catch {
    set fileOptions [open $optionsFile w]
    puts $fileOptions "# Options file for the NIST IFC File Analyzer v[getVersion] ([string trim [clock format [clock seconds]]])\n# Do not edit or delete from user home directory $mydocs  Doing so might corrupt the current settings or cause errors.\n"

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
    puts $fileOptions " "

    set winpos "+300+200"
    set wg [winfo geometry .]
    catch {set winpos [string range $wg [string first "+" $wg] end]}
    puts $fileOptions "set winpos \"$winpos\""
    set wingeo [string range $wg 0 [expr {[string first "+" $wg]-1}]]
    puts $fileOptions "set wingeo \"$wingeo\""

    set varlist(1) [list statusFont row_limit upgradeIFCsvr ifaVersion writeDirType]
    set varlist(2) [list fileDir fileDir1 userWriteDir userEntityFile lastXLS lastXLS1]
    set varlist(3) [list openFileList dispCmd dispCmds]
    foreach idx {1 2 3} {
      foreach var $varlist($idx) {
        if {[info exists $var]} {
          set vartmp [set $var]
          if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
            regsub -all {\\} $vartmp "/" vartmp
            regsub -all {\[} $vartmp "\\\[" vartmp
            regsub -all {\]} $vartmp "\\\]" vartmp
            if {$var != "dispCmds" && $var != "openFileList"} {
              puts $fileOptions "set $var \"$vartmp\""
            } else {
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
        if {$var == "openFileList"} {puts $fileOptions " "}
      }
      if {$idx < 3} {puts $fileOptions " "}
    }
    close $fileOptions

  } emsg]} {
    errorMsg "Error writing to options file: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------
proc displayResult {} {
  global appName dispCmd File localName

  set dispFile $localName
  set idisp [file rootname [file tail $dispCmd]]
  if {[info exists appName]} {if {$appName != ""} {set idisp $appName}}
  outputMsg "Opening IFC file in: $idisp"

# display file
  if {[string first "Indent" $idisp] == -1 && [string first "Default" $idisp] == -1} {

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
  global appName appNames dispApps dispCmd dispCmds drive padcmd pf32 pf64

  regsub {\\} $pf32 "/" p32
  lappend pflist $p32
  if {$pf64 != "" && $pf64 != $pf32} {
    regsub {\\} $pf64 "/" p64
    lappend pflist $p64
  }

  set lastver 0
  set ok 0

# IFC viewers
  foreach pf $pflist {
    set applist [list \
      [list [file join $pf "CAD Assistant" CADAssistant.exe] "CAD Assistant"] \
      [list [file join $pf "Data Design System" Viewer Exe DdsViewer.exe] "DDS-CAD Viewer"] \
      [list [file join $pf "Geometry Gym" ggIfcTreeViewer ggIFCTreeViewer.exe] "GeometryGym IFC Browser"] \
      [list [file join $pf "Tekla BIMsight" BIMsight.exe] "Tekla BIMsight"] \
      [list [file join $pf Areddo Areddo.exe] Areddo] \
      [list [file join $pf CSTB eveBIM bin eveBIM.exe] eveBIM] \
      [list [file join $pf Datacomp "BIM Vision" bim_vision_x64.exe] "BIM Vision"] \
      [list [file join $pf IFCBrowser IfcQuickBrowser.exe] IfcQuickBrowser] \
      [list [file join $pf Kisters 3DViewStation 3DViewStation.exe] 3DViewStation] \
      [list [file join $pf KUBUS "BIMcollab ZOOM" "BIMcollab ZOOM.exe"] "BIMcollab ZOOM"] \
      [list [file join $pf Solibri SOLIBRI Solibri.exe] "Solibri Anywhere"] \
      [list [file join $pf Trimble "Trimble Connect" TrimbleConnect.exe] "Trimble Connect"] \
   ]
    foreach app $applist {
      if {[file exists [lindex $app 0]]} {
        set name [lindex $app 1]
        set dispApps([lindex $app 0]) $name
      }
    }

    set applist [list \
      [list {*}[glob -nocomplain -directory [file join $pf ODA] -join "Open IFC Viewer*" OpenIFCViewer.exe] "Open IFC Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf Datacubist] -join "*" "simplebim*.exe"] "simplebim"] \
      [list {*}[glob -nocomplain -directory [file join $pf Solibri] -join "*" "Solibri Model Checker v*.exe"] "Solibri Model Checker"] \
      [list {*}[glob -nocomplain -directory [file join $pf Solibri] -join "*" "Solibri Model Viewer v*.exe"] "Solibri Model Viewer"] \
    ]

    foreach app $applist {
      if {[llength $app] == 2} {
        set match [join [lindex $app 0]]
        if {$match != "" && ![info exists dispApps($match)]} {
          set dispApps($match) [lindex $app 1]
        }
      }
    }
  }

  if {[file exists [file join $drive ACCA usBIM.viewer+ usBIM.viewer.exe]]} {
    set name "usBIM.viewer"
    set dispApps([file join $drive ACCA usBIM.viewer+ usBIM.viewer.exe]) $name
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
  foreach pf $pflist {
    set padcmd [file join $pf Notepad++ notepad++.exe]
    if {[file exists $padcmd]} {
      set padnam "Notepad++"
      set dispApps($padcmd) $padnam
      break
    }
  }
  if {$padnam == ""} {
    set padcmd [file join Windows system32 Notepad.exe]
    set padnam "Notepad"
    set dispApps($padcmd) $padnam
  }

#-------------------------------------------------------------------------------
# remove cmd that do not exist in dispCmds and non-executables
  set dispCmds1 {}
  foreach app $dispCmds {
    set fext [file extension $app]
    if {([file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0) && \
        [file tail $app] != "NotePad.exe" && [string first "Analyzer.exe" $app] == -1} {
      lappend dispCmds1 $app
    }
  }
  set dispCmds $dispCmds1

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

# remove old commands
  set ndcs {}
  foreach cmd $dispCmds {
    set ok 1
    foreach bcmd [list ifcview ifccheckgui stepbrws apconformgui] {
      append bcmd ".exe"
      if {[string first $bcmd $cmd] != -1} {set ok 0}
    }
    if {$ok} {lappend ndcs $cmd}
  }
  set dispCmds $ndcs

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
  global buttons File filemenuinc lenlist localName openFileList

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

# check if file name is already in the menu, if so, delete
  set ifile [lsearch $openFileList $localName]
  if {$ifile > 0} {
    set openFileList [lreplace $openFileList $ifile $ifile]
    $File delete [expr {$ifile+$filemenuinc}] [expr {$ifile+$filemenuinc}]
  }

# insert file name at top of list
  set fext [string tolower [file extension $localName]]
  if {$ifile != 0 && $fext == ".ifc"} {
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
  }

# save the state so that if the program crashes the file list will be already saved
  saveState
  return
}

#-------------------------------------------------------------------------------
# file size in KB or MB
proc fileSize {fn} {
  set fs [expr {[file size $fn]/1024}]
  if {$fs < 10000} {
    return "$fs KB"
  } else {
    set fs [expr {round(double($fs)/1024.)}]
    return "$fs MB"
  }
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
      ::tcom::ref createobject Excel.Application

# errors
    } emsg]} {
      errorMsg "Error starting Excel: $emsg"
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
  global buttons lastXLS localName

  set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
  if {[llength $pid1] > 0} {
    if {[info exists buttons]} {
      if {!$multFile} {
        set msg "There are at least ([llength $pid1]) Excel spreadsheets already opened.\n\nDo you want to close the spreadsheets?"
        set dflt yes
        if {[info exists lastXLS] && [info exists localName]} {
          if {[llength $pid1] == 1} {if {[string first [file nativename [file rootname $localName]] [file nativename $lastXLS]] != 0} {set dflt no}}
        }
        set choice [tk_messageBox -type yesno -default $dflt -message $msg -icon question -title "Close Spreadsheets?"]
        if {$choice == "yes"} {
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
        }
      }
    } else {
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
# trimNum gets used mostly when processing IFC files
proc trimNum {num {prec 3}} {
  global unq_num

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
  global mydesk mydocs

  if {[string first $mydocs $fname] == 0} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydocs] end]"
  } elseif {[string first $mydesk $fname] == 0 && $mydesk != $fname} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydesk] end]"
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
    set fn "[string range $fn 0 $c1-1]-$i$fext"
    catch {[file delete -force -- $fn]}
    if {![file exists $fn]} {break}
  }
  return $fn
}

#-------------------------------------------------------------------------------
# check file name for bad characters
proc checkFileName {fn} {
  global mydocs

  set fnt [file tail $fn]
  set fnd [file dirname $fn]
  if {[string first "\[" $fnd] != -1 || [string first "\]" $fnd] != -1} {
    set fn [file nativename [file join $mydocs $fnt]]
    errorMsg "Saving Spreadsheet to the home directory instead of the IFC file directory because of the \[ and \] in the directory name." red
  }
  if {[string first "\[" $fnt] != -1 || [string first "\]" $fnt] != -1} {
    regsub -all {\[} $fn "(" fn
    regsub -all {\]} $fn ")" fn
    errorMsg "\[ and \] are replaced by ( and ) in the Spreadsheet file name." red
  }
  return $fn
}

#-------------------------------------------------------------------------------
# install IFCsvr (or uninstall to reinstall)
proc installIFCsvr {{exit 0}} {
  global buttons ifcsvrKey mydocs mytemp upgradeIFCsvr wdir

# IFCsvr version depends on string entered when IFCsvr is repackaged for new IFC schemas
  set versionIFCsvr 20240111

# if IFCsvr is alreadly installed, get version from registry, decide to reinstall newer version
  if {[catch {

# check IFCsvr CLSID and get version registry value "yyyy.mm.dd" or old format "1.0.0 (NIST Update yyyy-mm-dd)"
    set ifcsvrKey "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{3C8CE0A4-803B-48A6-96A0-A3DDD5AE5596}"
    set verIFCsvr [registry get $ifcsvrKey {DisplayVersion}]

# format version to be yyyymmdd
    set c1 [string first "20" $verIFCsvr]
    if {$c1 != -1} {
      set verIFCsvr [string range $verIFCsvr $c1 end]
      if {[string index $verIFCsvr end] == ")"} {set verIFCsvr [string range $verIFCsvr 0 end-1]}
      regsub -all {\-} $verIFCsvr "" verIFCsvr
      regsub -all {\.} $verIFCsvr "" verIFCsvr
    } else {
      set verIFCsvr 0
    }

# old version, reinstall
    if {$verIFCsvr < $versionIFCsvr} {
      set reinstall 1

# up-to-date, do nothing
    } else {
      set reinstall 2
      set upgradeIFCsvr [clock seconds]
    }

# IFCsvr not installed or can't read registry
  } emsg]} {
    set reinstall 0
  }

# up-to-date
  if {$reinstall == 2} {return}

  set ifcsvr     "ifcsvrr300_setup_1008_en-update.msi"
  set ifcsvrInst [file join $wdir exe $ifcsvr]

  if {[info exists buttons]} {.tnb select .tnb.status}
  outputMsg " "

# first time installation
  if {!$reinstall} {
    errorMsg "The IFCsvr toolkit must be installed to read and process IFC files."
    outputMsg "- You might need administrator privileges (Run as administrator) to install the toolkit.
  Antivirus software might respond that there is a security issue with the toolkit.  The
  toolkit is safe to install.  Use the default installation folder for the toolkit.
- To reinstall the toolkit, run the installation file ifcsvrr300_setup_1008_en-update.msi
  in $mytemp
- After the toolkit is installed, see Help > IFC Support to see which versions of IFC are supported.

- If the software crashes the first time you run it, first uninstall the IFCsvr toolkit.  Then run
  software as Administrator and when prompted, install the IFCsvr toolkit for Everyone, not
  Just Me.  Subsequently, the software does not have to be run as Administrator."

    if {[file exists $ifcsvrInst]} {
      set msg "The IFCsvr toolkit must be installed to read and process IFC files.  After clicking OK the IFCsvr toolkit installation will start."
      append msg "\n\nYou might need administrator privileges (Run as administrator) to install the toolkit.  Antivirus software might respond that there is a security issue with the toolkit.  The toolkit is safe to install.  Use the default installation folder for the toolkit."
      append msg "\n\nAfter the toolkit is installed, see Help > IFC Support to see which versions of IFC are supported."
      append msg "\n\nIf the software crashes the first time you run it, first uninstall the IFCsvr toolkit.  Then run the software as Administrator and when prompted, install the IFCsvr toolkit for Everyone, not Just Me.  Subsequently, the software does not have to be run as Administrator."
      set choice [tk_messageBox -type ok -message $msg -icon info -title "Install IFCsvr"]
      outputMsg "\nWait for the installation to finish before processing an IFC file." red
    } elseif {![info exists buttons]} {
      outputMsg "\nRerun this program after the installation has finished to process an IFC file."
    }

# reinstall
  } else {
    errorMsg "The existing IFCsvr toolkit must be reinstalled to update for new versions of IFC."
    outputMsg "- First REMOVE the current installation of the IFCsvr toolkit."
    outputMsg "    In the IFCsvr Setup Wizard select 'REMOVE IFCsvrR300 ActiveX Component' and Finish" red
    outputMsg "    If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
    if {[info exists buttons]} {
      outputMsg "- Then restart this software or process an IFC file to install the updated IFCsvr toolkit."
    } else {
      outputMsg "- Then run this software again to install the updated IFCsvr toolkit."
    }
    outputMsg "- After the toolkit is reinstalled, see Help > IFC Support to see which versions of IFC are supported."

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be reinstalled to update for new versions of IFC."
      append msg "\n\nFirst REMOVE the current installation of the IFCsvr toolkit."
      append msg "\n\nIn the IFCsvr Setup Wizard (after clicking OK below) select 'REMOVE IFCsvrR300 ActiveX Component' and Finish.  If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
      append msg "\n\nThen restart this software or process an IFC file to install the updated IFCsvr toolkit."
      append msg "\n\nAfter the toolkit is reinstalled, see Help > IFC Support to see which versions of IFC are supported."
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
          errorMsg "Error copying the IFCsvr toolkit installation file to a directory."
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
      set upgradeIFCsvr [clock seconds]
      saveState
      if {$exit} {exit}
    } emsg]} {
      errorMsg "Error installing IFCsvr toolkit: $emsg"
    }

# cannot find the toolkit
  } else {
    if {[file exists $ifcsvrInst]} {errorMsg "The IFCsvr toolkit cannot be automatically installed."}
    catch {.tnb select .tnb.status}
    update idletasks

# manual install instructions
    outputMsg "To manually install the IFCsvr toolkit:
- The installation file ifcsvrr300_setup_1008_en-update.msi can be found in $mytemp
- Run the installer and follow the instructions.  Use the default installation folder for IFCsvr.
  You might need administrator privileges (Run as administrator) to install the toolkit.\n"
    after 1000
    errorMsg "Opening folder: $mytemp"
    if {[catch {
      exec {*}[auto_execok start] [file nativename $mytemp]
      if {$exit} {exit}
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {errorMsg "Error opening directory: $emsg"}
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
