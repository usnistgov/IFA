# -------------------------------------------------------------------------------
# which IFC entities are processed depending on options
proc ifcWhichEntities {ok enttyp} {
  global opt ifcall

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
# check for valid GUID
proc ifcCheckGUID {objName ov lastguid} {
  global ifc
	
  if {[string length $ov] == 22} {
    if {[string equal $ov $lastguid]} {
      errorMsg " Duplicate GUID on $ifc"
    } elseif {[string equal -length 8 $ov $lastguid]} {
      errorMsg " GUID on $ifc is similar to other GUIDs and might not be Globally Unique" red
    }
  } else {
    errorMsg " GUID on $ifc must be 22 characters long"
  }
  return $ov
}

# -------------------------------------------------------------------------------
proc ifcFormatPropertySet {} {
  global worksheet ifc row rowmax opt fileschema ifcpset2x3 excel
	
  set nsprop {}
  if {$ifc == "IfcPropertySet"} {
    set hlpset [$worksheet($ifc) Hyperlinks]

# limit checking to first 20000 entities
    set pslim1 20000
    set pslim [expr {$row($ifc)+2}]
    if {$pslim > $rowmax} {incr pslim -1}
   
    if {$pslim > $pslim1} {
      set pslim $pslim1
      outputMsg " Adding some links on IfcPropertySet to IFC documentation" blue
    } else {
      outputMsg " Adding links on IfcPropertySet to IFC documentation" blue
    }

    set pslim [expr {$pslim-2}]	
    set rvals {}
    for {set i 4} {$i <= $pslim} {incr i} {
      set rval ""
      if {[catch {
	if {!$opt(PR_GUID)} {
	  set range [$worksheet($ifc) Range "B$i"]
	} else {
	  set range [$worksheet($ifc) Range "D$i"]
	}
	set rval [$range Value]

# check only new property set names
	if {[lsearch $rvals $rval] == -1} {
	  lappend rvals $rval
	  set ok 0

	  if {[string first "IFC4" $fileschema] == -1} {
	    if {[string first "Pset" $rval] == 0} {
	      if {[info exists ifcpset2x3($rval)]} {
		set docurl "https://standards.buildingsmart.org/IFC/RELEASE/IFC2x3/TC1/HTML/psd/$ifcpset2x3($rval)/$rval.xml"
		$hlpset Add $range [join $docurl] [join ""] [join "$rval IFC2x3 Documentation"]
		set ok 1
	      }
	    } elseif {[string tolower [string range $rval 0 3]] == "pset"} {
	      foreach item [array names ifcpset2x3] {
		if {[string tolower $item] == [string tolower $rval] && !$ok} {
		  set docurl "https://standards.buildingsmart.org/IFC/RELEASE/IFC2x3/TC1/HTML/psd/$ifcpset2x3($item)/$item.xml"
		  $hlpset Add $range [join $docurl] [join ""] [join "$item IFC2x3 Documentation"]
		  set ok 1
		  break
		}
	      }
	    }
	  } else {
	    set docurl "https://standards.buildingsmart.org/IFC/RELEASE/IFC4/FINAL/HTML/link/[string tolower $rval].htm"
	    $hlpset Add $range [join $docurl] [join ""] [join "$rval IFC4 Documentation"]
	    set ok 1
	  }

# property set does not have documentation
	  if {!$ok} {
	    [$range Interior] ColorIndex [expr 19]
	    if {[expr {int([$excel Version])}] >= 12} {
	      for {set k 7} {$k <= 12} {incr k} {
		catch {if {$k != 9 || $i != $pslim} {[[$range Borders] Item [expr $k]] Weight [expr 1]}}
	      }
	    }
	    if {[lsearch $nsprop $rval] == -1 && $rval != ""} {lappend nsprop $rval}
	  }
	}

      } emsg]} {
	errorMsg "ERROR adding IfcPropertySet links: $emsg"
	catch {raise .}
      }
    }
  }
 
  if {[llength $nsprop] > 0} {
    if {$pslim < $pslim1} {
      errorMsg " Non-standard IfcPropertySet Names ([llength $nsprop])" green
    } else {
      errorMsg " Some non-standard IfcPropertySet Names ([llength $nsprop])" green
    }
    foreach item $nsprop {
      set item [string trim [join $item]]
      errorMsg "  $item" green
    }
  }
 
# align values on IfcPropertySingleValue when they are counted
  if {$ifc == "IfcPropertySingleValue" && $opt(COUNT)} {
    for {set i 4} {$i <= [expr {$row($ifc)+2}]} {incr i} {
      set rval ""
      if {[catch {
	set range [$worksheet($ifc) Range "D$i"]
	set rval [$range Value]
	if {$rval == "(Real)" || $rval == "(Integer)"} {$range HorizontalAlignment [expr -4152]}
      } emsg]} {
	errorMsg "ERROR aligning cells on IfcPropertySingleValue: $emsg"
	catch {raise .}
      }
    }
  }
}

# -------------------------------------------------------------------------------
proc ifcExpandEntities {refType refEntity counting} {
  global type ifc lpnest countEnts opt row
	
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
  global cells row col heading ifc opt lpnest

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
