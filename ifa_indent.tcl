proc indentPutLine {line {comment ""} {putstat 1}} {
  global idlist idpatr indentWriteFile lpatr npatr spatr

  if {$putstat != 1} {return 0}
  set stat 1

  set ll [string length $line]
  set t1 [string first "\#" $line]
  set t2 [string first "=" [string range $line [expr {$t1 + 1}] end]]
  set indent [expr {$t1 + $t2 + 2}]

# reset id list
  if {$t1 == 0} {
    set idlist ""
    catch {unset idpatr}
    catch {unset lpatr}
    catch {unset spatr}
    set npatr -1
  }
  set id [string range $line $t1 [expr {$t1+$t2}]]

  if {[string first $id $idlist] != -1} {
    set id1 [string first $id $idpatr]
    if {$id1 != -1} {
      if {$id1 == 0} {
        for {set i 0} {$i < $npatr} {incr i} {
          if {$lpatr($i) == $lpatr($npatr)} {
            errorMsg "Repeating pattern of entities"
            puts $indentWriteFile "(Repeating pattern of entities)"
            set stat -1
            return $stat
          }
        }
        incr npatr
      }
      lappend lpatr($npatr) $id1
      append spatr($npatr) $id
    }
    append idpatr $id
  } else {
    set idpatr ""
  }
  append idlist [string range $line $t1 [expr {$t1+$t2}]]

  if {$indent > 40} {
    errorMsg "Indentation greater than 40 spaces" red
    puts $indentWriteFile "(Indentation greater than 40 spaces)"
    set stat 0
    return $stat
  }

  if {$ll <= 120} {
    puts $indentWriteFile $line

  } else {
    set line1 [string range $line 90 end]
    set p1 [string first ")"  $line1]
    set p2 [string first ");" $line1]
    if {$p1 == $p2} {
      puts $indentWriteFile $line

    } else {
      set line2 [string range $line1 [expr {$p1 + 1}] end]
      if {[string length $line2] > 9} {
        puts $indentWriteFile [string range $line 0 [expr {$p1 + 90}]]
        set line3 ""
        for {set i 0} {$i < $indent} {incr i} {append line3 " "}
        append line3 $line2

        set ll [string length $line3]
        if {$ll <= 120} {
          puts $indentWriteFile $line3

        } else {
          set line4 [string trim [string range $line3 90 end]]
          set p1 [string first ")"  $line4]
          set p2 [string first ");" $line4]
          if {$p1 == $p2} {
            puts $indentWriteFile $line3

          } else {
            set line5 [string range $line4 [expr {$p1 + 1}] end]
            if {[string length $line5] > 9} {
              puts $indentWriteFile [string range $line3 0 [expr {$p1 + 90}]]
              set line6 ""
              for {set i 0} {$i < $indent} {incr i} {append line6 " "}
              append line6 $line5
              puts $indentWriteFile $line6
            } else {
              puts $indentWriteFile $line3
            }
          }
        }
      } else {
        puts $indentWriteFile $line
      }
    }
  }
  return $stat
}

#-------------------------------------------------------------------------------
proc indentSearchLine {line ndent} {
  global comment indentdat2 indentEntity indentMissing indentstat

  incr ndent

  set sline [split $line "\#"]
  for {set i 1} {$i < [llength $sline]} {incr i} {
    set str [lindex $sline $i]
    set id [indentGetID $str]
    if {[string is integer $id]} {
      set str "  "
      for {set j 1} {$j < $ndent} {incr j} {append str "  "}
      if {[info exists indentEntity($id)]} {
        append str $indentEntity($id)
        set cmnt ""
        if {[info exists comment($id)]} {set cmnt $comment($id)}
        if {![info exists indentstat]} {set indentstat 1}
        set indentstat [indentPutLine $str $cmnt $indentstat]
        if {$indentstat ==  0} {break}
        set line1 [string range $indentEntity($id) 1 end]
        if {[string first "\#" $line1] != -1} {
          set ok 1

# check for entities that stop the indentation ($indentdat2)
          foreach idx $indentdat2 {
            if {[string first $idx $line1] != -1} {
              set ok 0
              break
            }
          }
          if {$indentstat == -1} {set ok 0}

          if {$ok} {
            set stat [indentSearchLine $line1 $ndent]
            if {$stat == 0} {break}
          }
        }
      } else {
        if {[lsearch $indentMissing "#$id"] == -1} {lappend indentMissing "#$id"}
      }
    }
  }
  return 1
}

#-------------------------------------------------------------------------------
proc indentFile {ifile} {
  global errmsg indentdat2 indentEntity indentMissing indentPass indentReadFile indentstat indentWriteFile padcmd writeDir writeDirType

# indent on these IFC entities
  set indentdat1 [list \
    IFCACTUATOR IFCAIRTERMINAL IFCAIRTERMINALBOX IFCAIRTOAIRHEATRECOVERY IFCALARM IFCASSET IFCAUDIOVISUALAPPLIANCE \
    IFCBEAM IFCBOILER IFCBUILDINGELEMENTPART IFCBUILDINGELEMENTPROXY IFCBUILDINGSYSTEM IFCBURNER \
    IFCCABLECARRIERFITTING IFCCABLECARRIERSEGMENT IFCCABLEFITTING IFCCABLESEGMENT IFCCHILLER IFCCHIMNEY IFCCOIL IFCCOLUMN IFCCOMMUNICATIONSAPPLIANCE IFCCOMPRESSOR IFCCONDENSER IFCCONTROLLER IFCCOOLEDBEAM IFCCOOLINGTOWER IFCCOVERING IFCCURTAINWALL \
    IFCDAMPER IFCDISCRETEACCESSORY IFCDISTRIBUTIONCHAMBERELEMENT IFCDISTRIBUTIONCIRCUIT IFCDISTRIBUTIONCONTROLELEMENT IFCDISTRIBUTIONFLOWELEMENT IFCDISTRIBUTIONPORT IFCDISTRIBUTIONSYSTEM IFCDOOR IFCDUCTFITTING IFCDUCTSEGMENT IFCDUCTSILENCER \
    IFCELECTRICAPPLIANCE IFCELECTRICDISTRIBUTIONBOARD IFCELECTRICFLOWSTORAGEDEVICE IFCELECTRICGENERATOR IFCELECTRICMOTOR IFCELECTRICTIMECONTROL IFCELEMENTCOMPONENT IFCENERGYCONVERSIONDEVICE IFCENGINE IFCEVAPORATIVECOOLER IFCEVAPORATOR \
    IFCFAN IFCFASTENER IFCFILTER IFCFIRESUPPRESSIONTERMINAL IFCFLOWCONTROLLER IFCFLOWFITTING IFCFLOWINSTRUMENT IFCFLOWMETER IFCFLOWMOVINGDEVICE IFCFLOWSEGMENT IFCFLOWSTORAGEDEVICE IFCFLOWTERMINAL IFCFLOWTREATMENTDEVICE IFCFOOTING IFCFURNISH IFCFURNITURE \
    IFCHEATEXCHANGER IFCHUMIDIFIER \
    IFCINTERCEPTOR IFCINVENTORY \
    IFCJUNCTIONBOX \
    IFCLAMP IFCLIGHTFIXTURE \
    IFCMECHANICAL IFCMECHANICALFASTENER IFCMEDICALDEVICE IFCMEMBER IFCMOTORCONNECTION \
    IFCOCCUPANT IFCOPENING IFCOUTLET IFCOWNERHISTORY \
    IFCPILE IFCPIPEFITTING IFCPIPESEGMENT IFCPLATE IFCPRODUCTDEFINITIONSHAPE IFCPROTECTIVEDEVICE IFCPUMP \
    IFCRAIL IFCRAILING IFCRAMP IFCRAMPFLIGHT IFCREINFORCINGBAR IFCREINFORCINGELEMENT IFCREINFORCINGMESH IFCRELAGGREGATES IFCRELASSIGNSTOGROUP IFCRELASSOCIATES IFCRELCONNECTS IFCRELCONTAINEDINSPATIALSTRUCTURE IFCRELDECLARES IFCRELDEFINESBYOBJECT IFCRELDEFINESBYPROPERTIES IFCRELDEFINESBYTEMPLATE IFCRELFILLSELEMENT IFCRELNESTS IFCRELSEQUENCE IFCRELSPACEBOUNDARY IFCRELVOIDSELEMENT IFCROOF \
    IFCSANITARYTERMINAL IFCSENSOR IFCSHADINGDEVICE IFCSITE IFCSLAB IFCSOLARDEVICE IFCSPACE IFCSPACEHEATER IFCSTACKTERMINAL IFCSTAIR IFCSTAIRFLIGHT IFCSTRUCTURAL IFCSWITCHINGDEVICE IFCSYSTEMFURNITUREELEMENT \
    IFCTANK IFCTENDON IFCTRANSFORMER IFCTUBEBUNDLE \
    IFCUNITARYCONTROLELEMENT IFCUNITARYEQUIPMENT \
    IFCVALVE IFCVIBRATIONISOLATOR \
    IFCWALL IFCWASTETERMINAL IFCWINDOW \
  ]

# stop indenting when these entities are encountered
  set indentdat2 [list \
    IFCACTUATOR IFCAIRTERMINAL IFCAIRTERMINALBOX IFCAIRTOAIRHEATRECOVERY IFCALARM IFCASSET IFCAUDIOVISUALAPPLIANCE IFCAXIS2PLACEMENT3D \
    IFCBEAM IFCBOILER IFCBUILDING IFCBUILDINGELEMENTPART IFCBUILDINGELEMENTPROXY IFCBUILDINGSYSTEM IFCBURNER \
    IFCCABLECARRIERFITTING IFCCABLECARRIERSEGMENT IFCCABLEFITTING IFCCABLESEGMENT IFCCHILLER IFCCHIMNEY IFCCIRCLE IFCCOIL IFCCOLUMN IFCCOMMUNICATIONSAPPLIANCE IFCCOMPRESSOR IFCCONDENSER IFCCONTROLLER IFCCOOLEDBEAM IFCCOOLINGTOWER IFCCOVERING IFCCURTAINWALL \
    IFCDAMPER IFCDISCRETEACCESSORY IFCDISTRIBUTIONCHAMBERELEMENT IFCDISTRIBUTIONCIRCUIT IFCDISTRIBUTIONCONTROLELEMENT IFCDISTRIBUTIONFLOWELEMENT IFCDISTRIBUTIONPORT IFCDISTRIBUTIONSYSTEM IFCDOOR IFCDUCTFITTING IFCDUCTSEGMENT IFCDUCTSILENCER \
    IFCELECTRICAPPLIANCE IFCELECTRICDISTRIBUTIONBOARD IFCELECTRICFLOWSTORAGEDEVICE IFCELECTRICGENERATOR IFCELECTRICMOTOR IFCELECTRICTIMECONTROL IFCELEMENTASSEMBLY IFCELEMENTCOMPONENT IFCENERGYCONVERSIONDEVICE IFCENGINE IFCEVAPORATIVECOOLER IFCEVAPORATOR \
    IFCFACE IFCFAN IFCFASTENER IFCFILTER IFCFIRESUPPRESSIONTERMINAL IFCFLOWCONTROLLER IFCFLOWFITTING IFCFLOWINSTRUMENT IFCFLOWMETER IFCFLOWMOVINGDEVICE IFCFLOWSEGMENT IFCFLOWSTORAGEDEVICE IFCFLOWTERMINAL IFCFLOWTREATMENTDEVICE IFCFOOTING IFCFURNISH IFCFURNITURE \
    IFCGEOMETRICREPRESENTATIONCONTEXT \
    IFCHEATEXCHANGER IFCHUMIDIFIER \
    IFCINTERCEPTOR IFCINVENTORY \
    IFCJUNCTIONBOX \
    IFCLAMP IFCLIGHTFIXTURE IFCLOCALPLACEMENT \
    IFCMECHANICALFASTENER IFCMEDICALDEVICE IFCMEMBER IFCMOTORCONNECTION \
    IFCOCCUPANT IFCOPENING IFCOUTLET IFCOWNERHISTORY \
    IFCPILE IFCPIPEFITTING IFCPIPESEGMENT IFCPLATE IFCPOLYLINE IFCPRODUCTDEFINITIONSHAPE IFCPROJECT IFCPROTECTIVEDEVICE IFCPUMP \
    IFCRAIL IFCRAILING IFCRAMP IFCRAMPFLIGHT IFCREINFORCINGBAR IFCREINFORCINGELEMENT IFCREINFORCINGMESH IFCROOF \
    IFCSANITARYTERMINAL IFCSENSOR IFCSHADINGDEVICE IFCSHELLBASEDSURFACEMODEL IFCSITE IFCSLAB IFCSOLARDEVICE IFCSPACE IFCSPACEHEATER IFCSTACKTERMINAL IFCSTAIR IFCSTAIRFLIGHT IFCSTRUCTURAL IFCSWITCHINGDEVICE IFCSYSTEMFURNITUREELEMENT \
    IFCTANK IFCTENDON IFCTRANSFORMER IFCTUBEBUNDLE \
    IFCUNITARYCONTROLELEMENT IFCUNITARYEQUIPMENT \
    IFCVALVE IFCVIBRATIONISOLATOR \
    IFCWALL IFCWASTETERMINAL IFCWINDOW \
    IFCARBITRARYCLOSEDPROFILEDEF IFCARBITRARYOPENPROFILEDEF IFCARBITRARYPROFILEDEFWITHVOIDS IFCASYMMETRICISHAPEPROFILEDEF IFCCENTERLINEPROFILEDEF\
    IFCCIRCLEHOLLOWPROFILEDEF IFCCIRCLEPROFILEDEF IFCCSHAPEPROFILEDEF IFCDERIVEDPROFILEDEF IFCELLIPSEPROFILEDEF IFCISHAPEPROFILEDEF IFCLSHAPEPROFILEDEF \
    IFCRECTANGLEHOLLOWPROFILEDEF IFCRECTANGLEPROFILEDEF IFCROUNDEDRECTANGLEPROFILEDEF IFCTRAPEZIUMPROFILEDEF IFCTSHAPEPROFILEDEF IFCUSHAPEPROFILEDEF IFCZSHAPEPROFILEDEF \
  ]

  outputMsg "\nIndenting: [truncFileName [file nativename $ifile] 1] ([expr {[file size $ifile]/1024}] Kb)" blue
  outputMsg "Pass 1 of 2"
  set indentPass 1
  set indentReadFile [open $ifile r]

# same directory as file
  if {$writeDirType != 2} {
    set indentFileName [file rootname $ifile]

# user-defined directory
  } else {
    set indentFileName [file join $writeDir [file rootname [file tail $ifile]]]
  }
  append indentFileName "-ifa.txt"
  set indentWriteFile [open $indentFileName w]
  puts $indentWriteFile "Indent File generated by the NIST IFC File Analyzer (v[getVersion])  [clock format [clock seconds]]\nIFC file: [file nativename $ifile]\n"

  set cmnt ""
  foreach var {indentEntity idlist idpatr npatr lpatr spatr indentstat} {
    if {[info exists $var]} {unset $var}
  }
  if {[info exists errmsg]} {unset errmsg}
  set indentMissing {}

# read all entities
  set ihead 1
  while {[gets $indentReadFile line] >= 0} {
    set line [indentCheckLine $line]

    if {[string first "ENDSEC" $line] != -1} {set ihead 0}
    if {$ihead} {puts $indentWriteFile $line}

    if {[string first "\#" $line] == 0} {
      set id [string trim [string range $line 1 [expr {[string first "\=" $line] - 1}]]]
      set indentEntity($id) $line
      if {$cmnt != ""} {set comment($id) $cmnt}
      set cmnt ""
    } elseif {[string first "\/\*" $line] == 0} {
      set cmnt $line
    }
  }
  close $indentReadFile

  outputMsg "Pass 2 of 2"
  set indentPass 2
  set indentReadFile [open $ifile r]

# check for entities that start an indentation ($indentdat1)
  while {[gets $indentReadFile line] >= 0} {
    foreach var {idlist idpatr npatr lpatr spatr} {if {[info exists $var]} {unset $var}}
    set indentstat 1

    set line [indentCheckLine $line]
    set line1 [string range $line 1 end]
    if {[string first "\#" $line1] != -1} {
      foreach idx $indentdat1 {
        if {[string first $idx $line] != -1} {
          puts $indentWriteFile \n
          set stat [indentPutLine $line]
          if {$stat == 0} {break}
          set stat [indentSearchLine $line1 0]
          break
        }
      }
    }
  }
  close $indentReadFile
  close $indentWriteFile

  if {[llength $indentMissing] > 0} {errorMsg "Missing IFC entities: [lsort $indentMissing]"}

  set fs [expr {[file size $indentFileName]/1024}]
  if {$padcmd != "" && $fs < 30000} {
    outputMsg "Opening indented IFC file: [truncFileName [file nativename $indentFileName] 1] ($fs Kb)"
    exec $padcmd $indentFileName &
  } else {
    outputMsg "Indented IFC file written: [truncFileName [file nativename $indentFileName] 1] ($fs Kb)"
  }
}

#-------------------------------------------------------------------------------
proc indentCheckLine {line} {
  global indentPass indentReadFile

  if {[string last ";" $line] == -1 && [string last "*/" $line] == -1} {
    if {[gets $indentReadFile line1] != -1} {
      if {[string length $line] < 900} {
        append line $line1
      } else {
        if {$indentPass == 1} {outputMsg "Long line truncated: [string range $line 0 50] ..." red}
        set iline [string range $line 0 900]
        set c1 [string last "," $iline]
        set iline "[string range $iline 0 $c1] (truncated)"
        if {[string last ";" $line1] != -1} {
          return $iline
        } else {
          while {1} {
            gets $indentReadFile line2
            if {[string last ";" $line2] != -1} {return $iline}
          }
        }
      }
      if {[catch {set line [indentCheckLine $line]} err]} {errorMsg $err}
      return $line
    } else {
      return $line
    }
  } else {
    set line [string trim $line]
    return $line
  }
}

#-------------------------------------------------------------------------------
proc indentGetID {id} {
  set p1 [string first "," $id]
  set p2 [string first "\)" $id]
  if {$p1 != -1 && $p2 != -1} {
    if {$p1 < $p2} {set id [string range $id 0 [expr {$p1-1}]]}
    if {$p1 > $p2} {set id [string range $id 0 [expr {$p2-1}]]}
  } elseif {$p1 != -1} {
    set id [string range $id 0 [expr {$p1-1}]]
  } elseif {$p2 != -1} {
    set id [string range $id 0 [expr {$p2-1}]]
  }
  set id [string trim $id]
  return $id
}
