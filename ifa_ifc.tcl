proc ifcExpandEntities {refType refEntity counting} {
  global countEnts ifc lpnest opt row type

# expand IfcLocalPlacement
  if {$opt(EX_LP)} {
    if {[lsearch $type(PR_GEOM) $ifc] == -1} {
      if {[string first "IfcStructural" $ifc] == -1 || $opt(EX_ANAL)} {
        if {$refType == "IfcLocalPlacement"} {
          errorMsg " Expanding $refType on: $ifc" green
          ifcLocalPlacement $refEntity
          if {[info exists lpnest($ifc,2)]} {
            if {$lpnest($ifc,2) < $lpnest($ifc,1) && $row($ifc) > 4} {
              if {$opt(EX_LP)}    {set inccol 2}
              if {$opt(EX_A2P3D)} {incr inccol 3}
              incr col($ifc) [expr {$inccol*($lpnest($ifc,1)-$lpnest($ifc,2))}]
              incr lpnest($ifc,3)
            } elseif {$lpnest($ifc,2) > $lpnest($ifc,1)} {
              incr lpnest($ifc,3)
            }
            if {$lpnest($ifc,3) == 1 && $lpnest($ifc,1) != $lpnest($ifc,2)} {
              outputMsg " Varying nesting of IfcLocalPlacement ($lpnest($ifc,2) vs. $lpnest($ifc,1)).  Some cells and headers are not aligned and formatted." red
            }
          }
        }
      }

# expand IfcAxis2Placement
      if {$opt(EX_A2P3D)} {
        if {$refType == "IfcAxis2Placement3D" || $refType == "IfcAxis2Placement2D" || \
            [string first "IfcCartesianTransformationOperator" $refType] == 0 || \
            $refType == "IfcDirection"} {
          if {[lsearch $countEnts $ifc] == -1 || !$counting} {
            errorMsg " Expanding $refType on: $ifc" green
            putAttributes $refEntity
            set subLocDir 1
          }
        }
      }
    }
  }

# expand analysis model entities
  if {$opt(EX_ANAL)} {
    if {$refType == "IfcVertexPoint"} {
      errorMsg " Expanding $refType on: $ifc" green
      putAttributes $refEntity

    } elseif {[string first "IfcStructural" $ifc] == 0} {
      if {[string first "action" [string tolower $ifc]] != -1 && \
          [string first "Load" $refType] != -1} {
        errorMsg " Expanding $refType on: $ifc" green
        putValues $refEntity
      }
    }
  }
}

# -------------------------------------------------------------------------------------------------
proc ifcLocalPlacement {refEntity} {
  global cells col heading ifc lpnest opt row

  if {$row($ifc) == 4} {
    incr lpnest($ifc,1)
  } else {
    incr lpnest($ifc,2)
  }

  ::tcom::foreach refAttribute [$refEntity Attributes] {
    if {$heading($ifc) != 0} {putHeading $refAttribute}

    set subEntity [$refAttribute Value]
    incr col($ifc)
    if {$subEntity != ""} {
      set subType [$subEntity Type]
      $cells($ifc) Item $row($ifc) $col($ifc) "<[$refEntity P21ID]> $subType [$subEntity P21ID]"
      if {$subType == "IfcAxis2Placement3D"} {
        if {$opt(EX_A2P3D)} {putAttributes $subEntity}
      } elseif {$subType == "IfcLocalPlacement"} {
        if {$opt(EX_LP)} {ifcLocalPlacement $subEntity}
      }
    } else {
      $cells($ifc) Item $row($ifc) $col($ifc) "<[$refEntity P21ID]>"
    }
  }
}
