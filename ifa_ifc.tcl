# -------------------------------------------------------------------------------
# which IFC entities are processed depending on options
proc ifcWhichEntities {ok enttyp} {
  global ifcall opt

  if {$ok} {return $ok}

  set ok 0
  if {$opt(PR_PROP) && [lsearch $ifcall $enttyp] != -1 && \
	  (([string first "Propert"        $enttyp] != -1 || \
	    [string first "IfcDoorStyle"   $enttyp] == 0 || \
	    [string first "IfcWindowStyle" $enttyp] == 0) && \
	    [string first "IfcRel" $enttyp] == -1 && [string first "RelationShip" $enttyp] == -1)} {set ok 1}
  if {$opt(PR_QUAN) && [string first "Quantit" $enttyp] != -1} {set ok 1}
  if {$opt(PR_MTRL) && [string first "Materia" $enttyp] != -1 && \
		       [string first "IfcRel" $enttyp] == -1 && [string first "RelationShip" $enttyp] == -1} {set ok 1}
  if {$opt(PR_UNIT) && (([string first "Unit"       $enttyp] != -1 && \
			 [string first "Protective" $enttyp] == -1 && \
			 [string first "Unitary"    $enttyp] == -1) || [string first "DimensionalExponent" $enttyp] != -1)} {set ok 1}
  if {$opt(PR_RELA) && [lsearch $ifcall $enttyp] != -1 && \
    ([string first "Relationship" $enttyp] != -1 || \
     [string first "IfcRel" $enttyp] == 0)} {set ok 1}
  return $ok
}

# -------------------------------------------------------------------------------
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
        if {[string first "IfcAxis2Placement" $refType] == 0 || \
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
      } else {
        if {$opt(EX_LP)}    {ifcLocalPlacement $subEntity}
      }
    } else {
      $cells($ifc) Item $row($ifc) $col($ifc) "<[$refEntity P21ID]>"
    }
  }
}
