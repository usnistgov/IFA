proc getEntityCSV {objEntity} {
  global nproc row col thisEntType ifc count
  global developer roseLogical
  global ecount badattr
  global csvfile csvdirnam csvstr fcsv

# get entity type
  set thisEntType [$objEntity Type]
  set ifc $thisEntType

  set roseLogical(0) "FALSE"
  set roseLogical(1) "TRUE"
  set roseLogical(2) "UNKNOWN"

  incr nproc

# -------------------------------------------------------------------------------------------------
# csv file for each entity if it does not already exist
  if {![info exists csvfile($thisEntType)]} {
    set msg "[formatComplexEnt $thisEntType] ($ecount($thisEntType))"
    outputMsg $msg
    update idletasks

# open csv file
    set csvfile($thisEntType) 1
    set csvfname [file join $csvdirnam $thisEntType.csv]
    set fcsv [open $csvfname w]
    puts $fcsv "[formatComplexEnt $thisEntType] ($ecount($thisEntType))"
    #outputMsg $fcsv red

# headings in first row
    set csvstr "ID"
    ::tcom::foreach objAttribute [$objEntity Attributes] {append csvstr ",[$objAttribute Name]"}
    puts $fcsv $csvstr
    unset csvstr

    set row($thisEntType) 4
    set count($thisEntType) 0
    update idletasks
  }

# -------------------------------------------------------------------------------------------------
# start filling in the cells
  incr count($thisEntType)
  
# show progress with > 50000 entities
  if {$ecount($thisEntType) >= 50000} {
    set c1 [expr {$count($thisEntType)%20000}]
    if {$c1 == 0} {
      outputMsg " $count($thisEntType) of $ecount($thisEntType) processed"
      update idletasks
    }
  }

# entity ID
  set p21id [$objEntity P21ID]

# -------------------------------------------------------------------------------------------------
# for all attributes of the entity
  set nattr 0
  set csvstr $p21id
  set objAttributes [$objEntity Attributes]
  ::tcom::foreach objAttribute $objAttributes {
    set attrName [$objAttribute Name]
    #outputMsg "$p21id  $attrName  [$objAttribute NodeType]  [info exists badattr($thisEntType)]" red

    if {[catch {
      if {![info exists badattr($thisEntType)]} {
        set objValue [$objAttribute Value]

# look for bad attributes that cause a crash
      } else {
        set ok 1
        foreach ba $badattr($thisEntType) {if {$ba == $attrName} {set ok 0}}
        if {$ok} {
          set objValue [$objAttribute Value]
        } else {
          set objValue "???"
          errorMsg " Skipping '$attrName' attribute on $thisEntType" red
        }
      }

# error getting attribute value
    } emsgv]} {
      set msg "ERROR processing #[$objEntity P21ID]=[$objEntity Type] '$attrName' attribute: $emsgv"
      errorMsg $msg
      set objValue ""
      catch {raise .}
    }

    incr nattr

# -------------------------------------------------------------------------------------------------
# values in rows
    incr col($thisEntType)

# not a handle, just a single value
    if {[string first "handle" $objValue] == -1} {
      set ov $objValue
  
# if value is a boolean, substitute string roseLogical
      if {([$objAttribute Type] == "RoseBoolean" || [$objAttribute Type] == "RoseLogical") && [info exists roseLogical($ov)]} {set ov $roseLogical($ov)}

# check if displaying numbers without rounding
      append csvstr ",$ov"

# -------------------------------------------------------------------------------------------------
# if attribute is reference to another entity
    } else {
      
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      if {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
        set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
#  which causes an error and it has to be processed like a list below
        if {[catch {
          set refType [$refEntity Type]
          set valnotlist 1
        } emsg2]} {

# process like a list which is very unusual
          #if {$developer} {errorMsg " Attribute reference is a List: $emsg2"}
          catch {foreach idx [array names cellval] {unset cellval($idx)}}
          ::tcom::foreach val $refEntity {
            append cellval([$val Type]) "[$val P21ID] "
          }
          set str ""
          set size 0
          catch {set size [array size cellval]}

          if {$size > 0} {
            foreach idx [lsort [array names cellval]] {
              set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
              if {$ncell > 1 || $size > 1} {
                if {$ncell < 30} {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1]  "
                }
              } else {
                append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
              }
            }
          }
          append csvstr ",$str"
          set valnotlist 0
        }

# value is not a list which is the most common
        if {$valnotlist} {
          set str "[formatComplexEnt $refType 1] [$refEntity P21ID]"

# for length measure (and other measures), add the actual measure value
          if {$refType == "IfcMeasureWithUnit"} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "ValueComponent"} {set str "[$refAttribute Value]  ($str)"}
            }
          } elseif {$refType == "IfcMaterial"} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "Name" &&         [$refAttribute Value] != ""} {set str "$str  ([$refAttribute Value])"}
            }
          } elseif {$refType == "IfcMaterialLayerSet"} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "LayerSetName" && [$refAttribute Value] != ""} {set str "$str  ([$refAttribute Value])"}
            }
          } elseif {$refType == "IfcMaterialProfileSet"} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "Name" &&         [$refAttribute Value] != ""} {set str "$str  ([$refAttribute Value])"}
            }
          }
          append csvstr ",$str"
        }

# -------------------------------------------------------------------------------------------------
# node type 20=AGGREGATE (ENTITIES), usually SET or LIST, try as a tcom list or regular list (SELECT type)
      } elseif {[$objAttribute NodeType] == 20} {
        catch {foreach idx [array names cellval]     {unset cellval($idx)}}

        if {[catch {
          ::tcom::foreach val [$objAttribute Value] {

# collect the reference id's (P21ID) for the Type of entity in the SET or LIST
            append cellval([$val Type]) "[$val P21ID] "
          }

        } emsg]} {
          foreach val [$objAttribute Value] {
            append cellval([$val Type]) "[$val P21ID] "
          }
        }

# -------------------------------------------------------------------------------------------------
# format cell values for the SET or LIST
        set str ""
        set size 0
        catch {set size [array size cellval]}

        if {$size > 0} {
          foreach idx [lsort [array names cellval]] {
            set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
            if {$ncell > 1 || $size > 1} {
              if {$ncell < 30} {
                append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
              } else {
                append str "($ncell) [formatComplexEnt $idx 1]  "
              }
            } else {
              append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
            }
          }
        }
        append csvstr ",[string trim $str]"
      }
    }
  }
  #outputMsg "$fcsv $csvstr"

# write to CSV file
  if {[catch {
    puts $fcsv $csvstr
  } emsg]} {
    errorMsg "Error writing to CSV file for: $thisEntType"
  }

# -------------------------------------------------------------------------------------------------
# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {
    if {[info exists $var]} {unset $var}
  }
  update idletasks
  return 1
}
