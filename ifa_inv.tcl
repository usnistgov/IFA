# check 'inverses' with GetUsedIn
proc invFind {objEntity} {
  global env inverses invmsg invs opt

  set DEBUGINV $opt(DEBUGINV)
  if {$DEBUGINV} {outputMsg \ninvFind red}

  set objType [$objEntity Type]
  set objP21ID [$objEntity P21ID]
  set stat ""

  if {[catch {
    foreach inverse $inverses {
      set invEntities [$objEntity GetUsedIn [lindex $inverse 0] [lindex $inverse 1]]
      ::tcom::foreach invEntity $invEntities {
        set invType [$invEntity Type]
        set invRule [lindex $inverse 2]
        set msg " Inverse ($invRule) [formatComplexEnt $invType]:"
        set msgok 0
        set nattr 0

        ::tcom::foreach invAttribute [$invEntity Attributes] {
          set attrName  [string tolower [$invAttribute Name]]
          set attrValue [$invAttribute Value]
          if {$DEBUGINV} {
            incr nattr
            if {$nattr == 1} {outputMsg "Inverse [lindex $inverse 0]  [lindex $inverse 1]  [lindex $inverse 2]" blue}
            outputMsg "  attrName $attrName"
          }

# look for 'relating'
          if {[string first "relating" $attrName] != -1} {
            set stat "relating [lindex $inverse 0] [lindex $inverse 1]"
            if {$DEBUGINV} {outputMsg "relating $attrName [$invAttribute NodeType] [$invAttribute Value]" red}
            if {[string first "handle" $attrValue] != -1} {
              set subType [$attrValue Type]
              if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID]  $objType $objP21ID" red}
              if {[$attrValue P21ID] != $objP21ID} {
                if {$DEBUGINV} {outputMsg "  OK" red}
                lappend invs($invRule) "$subType [$attrValue P21ID]"
                if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                set msgok 1
              }
            }

# look for 'related'
          } elseif {[string first "related" $attrName] != -1} {
            if {[string first "handle" $attrValue] != -1} {
              set stat "related [$invAttribute NodeType] [lindex $inverse 0] [lindex $inverse 1]"
              if {[$invAttribute NodeType] == 18 || [$invAttribute NodeType] == 19} {
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttribute NodeType] [$invAttribute Value]" green}
                set subType [$attrValue Type]
                if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID] $objType $objP21ID" green}
                if {[$attrValue P21ID] != $objP21ID} {
                  if {$DEBUGINV} {outputMsg "  OK" green}
                  lappend invs($invRule) "$subType [$attrValue P21ID]"
                  if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                  set msgok 1
                }

# still needs some work to remove the catch below
# nodetype = 20
              } elseif {[$invAttribute NodeType] == 20 && $attrName != [string tolower [lindex $inverse 1]]} {
                set stat "related [$invAttribute NodeType] A [lindex $inverse 0] [lindex $inverse 1]"
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttribute NodeType] [$invAttribute Value]" magenta}
                if {$DEBUGINV} {outputMsg " $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                if {[catch {
                  ::tcom::foreach aval [$invAttribute Value] {
                    if {$DEBUGINV} {outputMsg "  $aval / [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                    if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                      if {$DEBUGINV} {outputMsg "  OK A  [$aval Type] [$aval P21ID]" magenta}
                      lappend invs($invRule) "[$aval Type] [$aval P21ID]"
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                } emsg1]} {

# still needs some work to remove the catch
                  set stat "related [$invAttribute NodeType] B [lindex $inverse 0] [lindex $inverse 1]"
                  foreach aval [$invAttribute Value] {
                    catch {
                      if {$DEBUGINV} {outputMsg "  [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                      if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                        if {$DEBUGINV} {outputMsg "  OK B" magenta}
                        lappend invs($invRule) "[$aval Type] [$aval P21ID]"
                        if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                        set msgok 1
                      }
                    }
                  }
                }
              }
            }
          }
        }

        if {$msgok} {
          if {[string first "used_in" $msg] != -1} {
            set msg [string range $msg 19 end]
            regsub ":" $msg "." msg
            regsub " " $msg ""  msg
            set msg " Used In: $msg"
          } else {
            set lmsg [split $msg " "]
            set newmsg " Inverse: [string range [lindex $lmsg 3] 0 end-1].[string range [lindex $lmsg 2] 1 end-1] > [formatComplexEnt [lindex $lmsg 4]] [lindex $lmsg 5]"
            set msg $newmsg
          }
        }

        if {$msgok && [string first $msg $invmsg] == -1} {
          append invmsg $msg
          errorMsg $msg blue
        }
      }
    }
  } emsg]} {
    if {$env(USERDOMAIN) != "NIST"} {
      errorMsg "Error processing Inverse for '[$objEntity Type]': $emsg" red
    } else {
      errorMsg "Error processing Inverse for '[$objEntity Type]': $emsg\n ($stat)" red
    }
  }
}

# -------------------------------------------------------------------------------
# report inverses
proc invReport {counting} {
  global cells cellval col colinv ifc invs row

# inverse values and heading
  foreach item [array names invs] {
    catch {foreach idx [array names cellval] {unset cellval($idx)}}

    foreach val $invs($item) {
      set val [split $val " "]
      set val0 [lindex $val 0]
      set val1 "[lindex $val 1] "
      if {[info exists cellval($val0)]} {
        if {[string first $val1 $cellval($val0)] == -1} {
          append cellval($val0) $val1
        }
      } else {
        append cellval($val0) $val1
      }
    }

    set str ""
    set size 0
    catch {set size [array size cellval]}

    if {$size > 0} {
      foreach idx [lsort [array names cellval]] {
        set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
        if {$ncell > 1 || $size > 1} {
          if {$ncell < 30 && !$counting} {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "($ncell) [formatComplexEntInv $idx 1] $cellval($idx)"
          } else {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "($ncell) [formatComplexEntInv $idx 1]"
          }
        } else {
          if {!$counting} {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "(1) [formatComplexEntInv $idx 1] $cellval($idx)"
          } else {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "(1) [formatComplexEntInv $idx 1]"
          }
        }
      }
    }

    set idx "$ifc $item"
    if {[info exists colinv($idx)]} {
      $cells($ifc) Item $row($ifc) $colinv($idx) [string trim $str]
    } else {
      while {[[$cells($ifc) Item 3 $col($ifc)] Value] != ""} {incr col($ifc)}
      $cells($ifc) Item $row($ifc) $col($ifc) [string trim $str]
    }

# heading
    if {[[$cells($ifc) Item 3 $col($ifc)] Value] == ""} {
      if {$item != "used_in"} {
        $cells($ifc) Item 3 $col($ifc) "INV-$item"
        set idx "$ifc $item"
      } else {
        $cells($ifc) Item 3 $col($ifc) "Used In"
        set idx "$ifc $item"
      }
      set colinv($idx) $col($ifc)
    }
  }
}

# -------------------------------------------------------------------------------
proc formatComplexEntInv {str {space 0}} {
  set str1 $str
  catch {
    if {[string first "_and_" $str1] != -1} {
      if {$space == 0} {
        regsub -all "_and_" $str1 ")(" str1
      } else {
        regsub -all "_and_" $str1 ") (" str1
      }
      if {[string first "." $str1] != -1} {
        regsub {\.} $str1 ")." str1
      } else {
        set str1 "$str1)"
      }
      set str1 "($str1"
      if {[string last ")" $str1] < [string last "(" $str1]} {outputMsg $str1 red}
    }
  }
  return $str1
}

# -------------------------------------------------------------------------------
# set column color, border, group for INVERSES and Used In
proc invFormat {rancol} {
  global cells col ifc invGroup row rowmax worksheet

  set igrp1 100
  set igrp2 0
  set i1 [expr {$rancol+1}]

# fix column widths
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($ifc) Item 3 $i] Value]
    if {$val == "Used In" || [string first "INV-" $val] != -1} {
      set range [$worksheet($ifc) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 255]
    }
  }
  [$worksheet($ifc) Columns] AutoFit
  [$worksheet($ifc) Rows] AutoFit

# set colors, borders
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($ifc) Item 3 $i] Value]
    if {[string first "INV-" $val] != -1 || [string first "Used In" $val] != -1} {
      set r1 $row($ifc)
      if {$r1 > $rowmax} {set r1 [expr {$r1-1}]}
      set range [$worksheet($ifc) Range [cellRange 3 $i] [cellRange $r1 $i]]
      if {[string first "INV-" $val] != -1} {
        [$range Interior] ColorIndex [expr 20]
      } else {
        [$range Interior] Color [expr 16768477]
      }
      if {$i < $igrp1} {set igrp1 $i}
      if {$i > $igrp2} {set igrp2 $i}
      set range [$worksheet($ifc) Range [cellRange 4 $i] [cellRange $r1 $i]]
      for {set k 7} {$k <= 12} {incr k} {
        if {$k != 9} {
          catch {[[$range Borders] Item [expr $k]] Weight [expr 1]}
        }
      }
      set range [$worksheet($ifc) Range [cellRange 3 $i] [cellRange 3 $i]]
      catch {
        [[$range Borders] Item [expr 7]]  Weight [expr 1]
        [[$range Borders] Item [expr 10]] Weight [expr 1]
      }
    }
  }

# group
  if {$igrp2 > 0} {
    set grange [$worksheet($ifc) Range [cellRange 1 $igrp1] [cellRange [expr {$row($ifc)+2}] $igrp2]]
    [$grange Columns] Group
    set invGroup($ifc) $igrp1
  }
}

# -------------------------------------------------------------------------------
# decide if inverses should be checked for this entity type
proc invSetCheck {enttyp} {
  global opt type userentlist

  set checkInv 0

# IFC entities
  if {($opt(PR_BEAM) && [lsearch $type(PR_BEAM) $enttyp] != -1) || \
      ($opt(PR_INFR) && [lsearch $type(PR_INFR) $enttyp] != -1) || \
      ($opt(PR_HVAC) && [lsearch $type(PR_HVAC) $enttyp] != -1) || \
      ($opt(PR_ELEC) && [lsearch $type(PR_ELEC) $enttyp] != -1) || \
      ($opt(PR_ANAL) && [lsearch $type(PR_ANAL) $enttyp] != -1) || \
      ($opt(PR_SRVC) && [lsearch $type(PR_SRVC) $enttyp] != -1) || \
      ($opt(PR_COMM) && [lsearch $type(PR_COMM) $enttyp] != -1)} {set checkInv 1}
  if {[info exists userentlist]} {
    if {[lsearch $userentlist $enttyp] != -1} {set checkInv 1}
  }
  if {$enttyp == "IfcPropertySet"} {set checkInv 1}

  return $checkInv
}
