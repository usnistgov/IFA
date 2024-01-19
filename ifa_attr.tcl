proc putAttributes {refEntity} {
  global cells col heading ifc row vals

  ::tcom::foreach refAttribute [$refEntity Attributes] {
    if {$heading($ifc) != 0} {putHeading $refAttribute}

    set subEntity [$refAttribute Value]
    set subName   [$refAttribute Name]

    if {$subEntity != ""} {

# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      if {[$refAttribute NodeType] == 18 || [$refAttribute NodeType] == 19} {
        ::tcom::foreach subAttribute [$subEntity Attributes] {
          incr col($ifc)
          set subValue [join [$subAttribute Value]]
          set subType  [$subAttribute Type]
          $cells($ifc) Item $row($ifc) $col($ifc) $subValue
        }

      } elseif {[string first "handle" $subEntity] == -1} {
        incr col($ifc)
        if {[string first "e-308" $subEntity] == -1} {
          $cells($ifc) Item $row($ifc) $col($ifc) [join $subEntity]
        }

# node type 20=AGGREGATE (ENTITIES), usually SET or LIST, try as a tcom list or regular list
      } elseif {[$refAttribute NodeType] == 20} {
        if {[catch {
          if {[info exists vals]} {unset vals}
          set str ""
          set nval 0
          ::tcom::foreach val [$refAttribute Value] {
            incr nval
            if {$nval != 1} {append str ", "}
            ::tcom::foreach val1 [$val Attributes] {
              set sval [split [$val1 Value] " "]
              set vals($nval,0) [trimNum [lindex $sval 0] 4]
              set vals($nval,1) [trimNum [lindex $sval 1] 4]
              append str "$vals($nval,0) $vals($nval,1)"
            }
          }

          incr col($ifc)
          if {[string length $str] > 1024} {set str "[string range $str 0 1019] ..."}
          $cells($ifc) Item $row($ifc) $col($ifc) $str

        } emsg]} {
          errorMsg "putAttributes: $emsg"
        }

# invalid node type
      } else {
        incr col($ifc)
        errorMsg "putAttributes: Unexpected NodeType [$refAttribute NodeType] ($subName) Expanding $ifc"
      }
    } else {
      incr col($ifc)
    }
  }
}

# -------------------------------------------------------------------------------------------------
proc putValues {refEntity} {
  global cells col heading ifc row

  ::tcom::foreach refAttribute [$refEntity Attributes] {
    if {$heading($ifc) != 0} {putHeading $refAttribute}

    set subEntity [$refAttribute Value]
    incr col($ifc)
    if {[string first "e-308" $subEntity] == -1} {$cells($ifc) Item $row($ifc) $col($ifc) [join $subEntity]}
  }
}

# -------------------------------------------------------------------------------------------------
proc filterHeading {heading} {
  global lastheading opt

  set stat 1
  if {!$opt(EX_A2P3D)} {
    if {($heading == "RefDirection" && ($lastheading == "Axis" || $lastheading == "Location")) || \
        ($heading == "Axis" && $lastheading == "Location") || \
        ($heading == "Location" && ($lastheading == "RelativePlacement" || $lastheading == "Position"))} {
      set stat 0
    }
  }
  return $stat
}

# -------------------------------------------------------------------------------------------------
proc putHeading {refAttribute} {
  global cells colclr heading ifc lastheading

  set refName [$refAttribute Name]
  if {[filterHeading $refName]} {
    $cells($ifc) Item 3 [incr heading($ifc)] $refName

    set inc 0
    if {($refName == "PlacementRelTo" && $refName == $lastheading)} {
      set inc 1
    } elseif {$refName == "RelativePlacement" && $lastheading == "RefDirection"} {
      set inc -1
    } elseif {$refName == "RelativePlacement" && $lastheading != "PlacementRelTo"} {
      set inc -1
    }
    lappend colclr($ifc) "$inc $heading($ifc)"
  }

  set lastheading $refName
}
