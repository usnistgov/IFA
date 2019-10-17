# generate an Excel spreadsheet from an IFC file
proc genExcel {{numFile 0}} {
  global all_entity attrsum attrused buttons cellcolors cells cells1 col col1 colclr colinv comma count countent countEnts csvdirnam csvfile
  global ecount entityCount entName env errmsg excel excel1 extXLS fcsv File file_entity fileschema fixent fixprm heading icolor
  global ifc ifcall2x3 ifcall2x4 ifcApplication ignored lastguid lastheading lastXLS lenfilelist localName localNameList lpnest
  global multiFile multiFileDir nline nproc nsheet opt padcmd pcount pcountRow pf32 row row_limit rowmax scriptName startrow
  global timestamp tlast total_entity type types userEntityFile userentlist userXLSFile verexcel workbook workbooks worksheet worksheet1 worksheets
  global writeDir writeDirType ws_last xname xnames yrexcel

  if {[info exists errmsg]} {set errmsg ""}

# check if IFCsvr is installed
  set ifcsvrdir [file join $pf32 IFCsvrR300 dll]
  if {![file exists [file join $ifcsvrdir IFCsvrR300.dll]]} {
    $buttons(genExcel) configure -state disable
    installIFCsvr
    return
  }

  set env(ROSE_RUNTIME) $ifcsvrdir
  set env(ROSE_SCHEMAS) $ifcsvrdir

  if {[info exists buttons]} {$buttons(genExcel) configure -state disable}
  catch {.tnb select .tnb.status}
  set lasttime [clock clicks -milliseconds]

  set multiFile 0
  if {$numFile > 0} {set multiFile 1}

# -------------------------------------------------------------------------------------------------
# connect to IFCsvr
  if {[catch {
    if {![info exists buttons]} {outputMsg "\n*** Begin ST-Developer output"}
    set objIFCsvr [::tcom::ref createobject IFCsvr.R300]
    if {![info exists buttons]} {outputMsg "*** End ST-Developer output"}

# print errors
  } emsg]} {
    errorMsg "\nERROR connecting to IFCsvr: $emsg"
    catch {raise .}
    return 0
  }

# -------------------------------------------------------------------------------------------------
# open IFC file
  if {[catch {
    set nline 0
    outputMsg "\nOpening IFC file"
    set fname $localName

# add file name and size to multi file summary
    if {$numFile != 0 && [info exists cells1(Summary)] && $opt(XLSCSV) == "Excel"} {
      set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
      set fn [string range [file nativename [truncFileName $fname]] $dlen end]
      set fn1 [split $fn "\\"]
      set fn2 [lindex $fn1 end]
      set idx [string first $fn2 $fn]
      if {[string length $fn2] > 40} {
        set div [expr {int([string length $fn2]/2)}]
        set fn2 [string range $fn2 0 $div][format "%c" 10][string range $fn2 [expr {$div+1}] end]
        set fn  [file nativename [string range $fn 0 $idx-1]$fn2]
      }
      regsub -all {\\} $fn [format "%c" 10] fn

      set colsum [expr {$col1(Summary)+1}]
      set range [$worksheet1(Summary) Range [cellRange 4 $colsum]]
      $cells1(Summary) Item 4 $colsum $fn
      if {$opt(XL_LINK1)} {[$worksheet1(Summary) Hyperlinks] Add $range [join $fname] [join ""] [join "Link to IFC file"]}
    }

# open file, count entities
    if {![info exists buttons]} {outputMsg "\n*** Begin ST-Developer output\n*** Check for error or warning messages up to 'End ST-Developer output' below"}
    set objDesign [$objIFCsvr OpenDesign [file nativename $fname]]
    if {![info exists buttons]} {outputMsg "*** End ST-Developer output\n"}

    set entityCount [$objDesign CountEntities "*"]
    outputMsg " $entityCount entities\n"
    if {$entityCount == 0} {errorMsg "There are no entities in the IFC file"}

# add AP, file size, entity count to multi file summary
    if {$numFile != 0 && [info exists cells1(Summary)] && $opt(XLSCSV) == "Excel"} {
      set fsize [expr {[file size $fname]/1024}]
      if {$fsize > 10240} {
        set fsize "[expr {$fsize/1024}] Mb"
      } else {
        append fsize " Kb"
      }
      $cells1(Summary) Item [expr {$startrow-1}] $colsum $fsize
      $cells1(Summary) Item $startrow $colsum $entityCount
    }

# error opening file, report the schema
  } emsg]} {
    errorMsg "ERROR opening IFC file"
    if {$emsg == "invalid command name \"\""} {
     set fext [string tolower [file extension $fname]]
      if {$fext != ".ifc" && $fext != ".ifcZIP"} {
        errorMsg "File extension not supported ([file extension $fname])" red
      } else {
        set fs [getSchema $fname]
        set c1 [string first "\{" $fs]
        if {$c1 != -1} {set fs [string trim [string range $fs 0 $c1-1]]}

        set okSchema 0
        foreach match [lsort [glob -nocomplain -directory $ifcsvrdir *.rose]] {
          set schema [string toupper [file rootname [file tail $match]]]
          set f1 [string range [file rootname [file tail $match]] 0 5]
          if {$f1 == "ifc2x3" || $f1 == "ifc4"} {lappend schemas $schema}
          if {$fs == $schema} {set okSchema 1; break}
        }
        if {!$okSchema} {
          outputMsg " $fs is not supported.  Only IFC2x3 and IFC4 are supported, however, IFC4.0.n addendums are not supported."
          outputMsg " See Websites > IFC Documentation"
        } else {
          set msg "Possible causes of the ERROR:"
          append msg "\n1 - Syntax errors in the IFC file"
          append msg "\n    The file must start with ISO-10303-21; and end with ENDSEC; END-ISO-10303-21;"
          append msg "\n    Try opening the file in some other IFC software, see Websites > Free IFC Software"
          append msg "\n2 - File or directory name contains accented, non-English, or symbol characters"
          append msg "\n     [file nativename $fname]"
          append msg "\n    Change the file or directory name"
          append msg "\n3 - If the problem is not with the IFC file, then restart this software and try again,"
          append msg "\n    or run this software as administrator, or reboot your computer"
          errorMsg $msg red
        }
      }
      if {[info exists padcmd]} {
        if {$padcmd != ""} {
          outputMsg " "
          errorMsg "Opening IFC file in text editor"
          exec $padcmd [file nativename $localName] &
        }
      }
    }

    #outputMsg "\nClosing IFCsvr" green
    if {[info exists errmsg]} {unset errmsg}
    catch {$objDesign Delete}
    catch {unset objDesign}
    catch {unset objIFCsvr}

    catch {raise .}
    return 0
  }

# -------------------------------------------------------------------------------------------------
# connect to Excel
  if {$opt(XLSCSV) == "Excel"} {
    if {[catch {
      set pid1 [checkForExcel $multiFile]
      set excel [::tcom::ref createobject Excel.Application]
      set pidexcel [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]

      set verexcel [expr {int([$excel Version])}]
      if {$verexcel < 12} {
        set extXLS "xls"
        set xlFormat [expr 56]
        set rowmax [expr {2**16}]
      } else {
        set extXLS "xlsx"
        set rowmax [expr {2**20}]
        set xlFormat [expr 51]
      }
      set yrexcel ""
      if {$verexcel == 9} {
        set yrexcel 2000
      } elseif {$verexcel == 10} {
        set yrexcel 2002
      } elseif {$verexcel == 11} {
        set yrexcel 2003
      } elseif {$verexcel == 12} {
        set yrexcel 2007
      } elseif {$verexcel == 14} {
        set yrexcel 2010
      } elseif {$verexcel == 15} {
        set yrexcel 2013
      } elseif {$verexcel == 16} {
        set yrexcel 2016
      }
      if {$verexcel >= 2000 && $verexcel < 2100} {set yrexcel $verexcel}

      if {$verexcel  < 12} {errorMsg " Some spreadsheet features are not available with this older version of Excel."}

# turning off ScreenUpdating saves A LOT of time
      $excel Visible 0
      catch {$excel ScreenUpdating 0}

      set rowmax [expr {$rowmax-2}]
      if {$row_limit < $rowmax} {set rowmax $row_limit}

# error with Excel
    } emsg]} {
      errorMsg "ERROR connecting to Excel: $emsg"
      errorMsg "The IFC File will be written to CSV files.  See the option on the Spreadsheet tab."
      set opt(XLSCSV) "CSV"
      checkValues
      tk_messageBox -type ok -icon error -title "ERROR connecting to Excel" -message "Cannot connect to Excel or Excel is not installed.\nThe IFC file will be written to CSV files.\nSee the option on the Spreadsheet tab."
      catch {raise .}
    }
  }

# -------------------------------------------------------------------------------------------------
# start worksheets
  if {$opt(XLSCSV) == "Excel"} {
    if {[catch {
      set workbooks  [$excel Workbooks]
      set workbook   [$workbooks Add]
      set worksheets [$workbook Worksheets]

# delete all but one worksheet
      catch {$excel DisplayAlerts False}
      set sheetCount [$worksheets Count]
      for {set n $sheetCount} {$n > 1} {incr n -1} {[$worksheets Item [expr $n]] Delete}
      set ws_last [$worksheets Item [$worksheets Count]]
      catch {$excel DisplayAlerts True}
      [$excel ActiveWindow] TabRatio [expr 0.7]

# determine decimal separator
      set sheet [$worksheets Item [expr 1]]
      set cell  [$sheet Cells]

      set A1 12345,67890
      $cell Item 1 A $A1
      set range [$sheet Range "A1"]
      set comma 0
      if {[$range Value] == 12345.6789} {
        set comma 1
        errorMsg "Using comma \",\" as the decimal separator for numbers" red
      }

# print errors
    } emsg]} {
      errorMsg "ERROR opening Excel workbooks and worksheets: $emsg"
      catch {raise .}
      return 0
    }
  }

# -------------------------------------------------------------------------------------------------
# all app names that might appear in header section
# took out UNIGRAPHICS  (put back in)
  set apps {FEMAP IDA-STEP "Autodesk Inventor" "Mechanical Desktop" "Kubotek KeyCreator" \
    "Implementor Forum Team" "I-DEAS" "T-Systems" "CATIA V5" "CATIA V6" "CATIA V4" CATIA "Open CASCADE" \
    "Alias - OpenModel" "Alias OpenModel" "UGS - NX" Theorem THEOREM "Theorem Solutions" "THEOREM SOLUTIONS" \
    "PRO/ENGINEER" CgiStepCamp FiberSim \
    "OneSpace Designer" Datakit DATAKIT "Datakit CrossCad" CADfix "ITI TranscenData" ADT ArchiCAD AxisVM \
    "Bentley Architecture" "Bentley Structural" "Bentley MEP" Bentley bocad Dlubal ETABS CADiE \
    SDS/2 SAP2000 CSI 3D+ CSC Cype "SPACE GASS" VectorWorks Nemetschek Octaga "Revit Building" \
    "Revit Architecture" "Revit Structure" "Revit MEP" Revit Scia Esa.Pt Tekla Allplan \
    "Digital Project" DDS Ramboll "CBS Pro" Kymdata Progman Vizelia "CADWorx Steel" ACTIVe3D \
    MagiCAD ATC IFCtest CADstudio DesignData GTSTRUDL ProSteel Adobe RISA Robot SCI STAAD StruCad \
    CERL Evasys Triforma MicroStation "Fastrak Building Designer" RFEM RSTAB COBIE GT GeorgiaTech \
    EliteCAD "AutoCAD Architecture" "IFC Engine" "Advance Steel" Graitec "Smart Plant 3D" \
    "SmartPlant 3D" "Frameworks Plus" "Frame Works Plus" Intergraph Parabuild CoCreate SteelWorks \
    BuiltWorks "CADS Planner" "House Designer" BlueThink Selvaag AutoCAD "Architectural Desktop" \
    Solibri BLIS PlantVision "Visio " RhinoBIM Prokon BSPro Cis2Export_dll CIS2export "CASE Center" \
    IfcExplorer Constructivity "IFC text editor" "RAM " "SIEMENS PLM Software NX" NX Siemens \
    DATAVISION "CV - CADDS 5" CADDS "ST-ACIS" ACIS CoreTechnologie Creo SolidWorks "CATIA Version 5" \
    "SIEMENS PLM Software NX 8.0" "SIEMENS PLM Software NX 7.0" Patran "Implementor Forum Team" \
    "Mentor Graphics" CADIF EXPRESSO Express2Java "Plant Vision" "Geometry Gym" "FZK Viewer" ADAPT \
    Alibre HiCAD "Alias Studio" "Alias AutoStudio" SolidEdge "Solid Edge" Kubotek PolyTrans Elysium \
    Parasolid PlanetCAD "SIEMENS PLM Software NX 8.5" "SIEMENS PLM Software NX 9.0" "3D_Evolution" \
    "jt_step translator" "Unigraphics" "STEP-NC Explorer" "STEP-NC Maker" "SIEMENS PLM Software NX 7.5" \
    "SIEMENS PLM Software NX 10.0" "SIEMENS PLM Software NX 11.0" "T3D tool generator"}

  set app1 ""
  set timestamp ""

# sort apps by string length
  proc sortlength2 {wordlist} {
    set words {}
    foreach word $wordlist {
      lappend words [list [string length $word] $word]
    }
    set result {}
    foreach pair [lsort -decreasing -integer -index 0 [lsort -ascii -index 1 $words]] {
      lappend result [lindex $pair 1]
    }
    return $result
  }
  set apps [sortlength2 $apps]

# -------------------------------------------------------------------------------------------------
# add header worksheet
  if {[catch {
    if {$opt(XLSCSV) == "Excel"} {
      outputMsg "Generating Header worksheet" blue
    } else {
      outputMsg "Generating Header CSV file" blue
    }
    set hdr "Header"
    if {$opt(XLSCSV) == "Excel"} {
    set worksheet($hdr) [$worksheets Item [expr 1]]
    $worksheet($hdr) Activate
    $worksheet($hdr) Name $hdr
    set ws_last $worksheet($hdr)
    set cells($hdr) [$worksheet($hdr) Cells]

# create directory for CSV files
    } else {
      foreach var {csvdirnam csvfname fcsv} {catch {unset $var}}
      set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-ifa-csv"
      file mkdir $csvdirnam
      set csvfname [file join $csvdirnam $hdr.csv]
      if {[file exists $csvfname]} {file delete -force $csvfname}
      set fcsv [open $csvfname w]
    }

    set fileschema "IFC2X3"
    set row($hdr) 0
    foreach attr {Name FileDirectory FileDescription FileImplementationLevel FileTimeStamp FileAuthor \
                  FileOrganization FilePreprocessorVersion FileOriginatingSystem FileAuthorisation SchemaName} {
      incr row($hdr)
      if {$opt(XLSCSV) == "Excel"} {
        $cells($hdr) Item $row($hdr) 1 $attr
      } else {
        set csvstr $attr
      }
      set objAttr [string trim [join [$objDesign $attr]]]

      if {$attr == "FileDirectory"} {
        if {$opt(XLSCSV) == "Excel"} {
          $cells($hdr) Item $row($hdr) 2 [$objDesign $attr]
        } else {
          append csvstr ",[$objDesign $attr]"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  [$objDesign $attr]"

      } elseif {$attr == "SchemaName"} {
        set sn [getSchema $fname]
        outputMsg "$attr:  $sn" blue
        if {$sn == "IFC4"} {errorMsg "IFC4.0.n addendums and IFC4.n versions are not supported.  See Websites > IFC Documentation"}
        if {$opt(XLSCSV) == "Excel"} {
          $cells($hdr) Item $row($hdr) 2 $sn
        } else {
          append csvstr ",$sn"
          puts $fcsv $csvstr
        }

        set fileschema [string toupper [string range $objAttr 0 5]]
        set stop 0
        if {[string first "AP2" $fileschema] == 0} {
          errorMsg "This is a STEP file that cannot be analyzed by the IFC File Analyzer.  Use the STEP File Analyzer."
          after 1000
          displayURL https://www.nist.gov/services-resources/software/step-file-analyzer
          set stop 1
        }
        if {$objAttr == "STRUCTURAL_FRAME_SCHEMA"} {
          errorMsg "This is a CIS/2 file that cannot be analyzed by the IFC File Analyzer.  Use the STEP File Analyzer."
          after 1000
          displayURL https://www.nist.gov/services-resources/software/step-file-analyzer
          set stop 1
        }

# stop everything and return
        if {$stop} {
          catch {
            $objDesign Delete
            unset objDesign
            unset objIFCsvr
            $excel Quit
            if {[info exists excel]} {unset excel}
            if {[llength $pidexcel] == 1} {twapi::end_process $pidexcel -force}
            raise .
          }
          update idletasks
          return 0
        }

      } else {
        if {$attr == "FileDescription" || $attr == "FileAuthor" || $attr == "FileOrganization"} {
          set str1 "$attr:  "
          set str2 ""
          foreach item [$objDesign $attr] {
            append str1 "[string trim $item], "
            if {$opt(XLSCSV) == "Excel"} {
              append str2 "[string trim $item][format "%c" 10]"
            } else {
              append str2 ",[string trim $item]"
            }
          }
          outputMsg [string range $str1 0 end-2]
          if {$opt(XLSCSV) == "Excel"} {
            $cells($hdr) Item $row($hdr) 2 "'[string trim $str2]"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } else {
            append csvstr [string trim $str2]
            puts $fcsv $csvstr
          }
        } else {
          outputMsg "$attr:  $objAttr"
          if {$opt(XLSCSV) == "Excel"} {
            $cells($hdr) Item $row($hdr) 2 "'$objAttr"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } else {
            append csvstr ",$objAttr"
            puts $fcsv $csvstr
          }
        }

        if {$attr == "FileImplementationLevel"} {if {$objAttr != "2\;1"} {errorMsg "Syntax Error: Implementation Level should be '2\;1'"}}

# add time stamp to multi file summary
        if {$attr == "FileTimeStamp"} {
          set timestamp $objAttr
          if {$numFile != 0 && [info exists cells1(Summary)] && $opt(XLSCSV) == "Excel"} {
            set colsum [expr {$col1(Summary)+1}]
            set range [$worksheet1(Summary) Range [cellRange 5 $colsum]]
            catch {$cells1(Summary) Item 5 $colsum "'[string range $timestamp 2 9]"}
          }
        }
      }
    }

    if {$opt(XLSCSV) == "Excel"} {
      set range [$worksheet($hdr) Range "A:A"]
      [$range Font] Bold [expr 1]
      [$worksheet($hdr) Columns] AutoFit
      [$worksheet($hdr) Rows] AutoFit
      catch {[$worksheet($hdr) PageSetup] Orientation [expr 2]}
      catch {[$worksheet($hdr) PageSetup] PrintGridlines [expr 1]}
    }

    set creo 0
    set appnist 0
    set attr "FilePreprocessorVersion"
    if {[string first "NIST CIS/2" [join [$objDesign $attr]]] != -1} {set appnist 1}

    foreach attr {FilePreprocessorVersion FileOriginatingSystem FileDescription \
                  FileAuthorisation FileOrganization} {

# set the application from the file, apps is a list of all application names defined above
      foreach app $apps {
        if {$app1 == "" && [string first [string tolower $app] [string tolower [join [$objDesign $attr]]]] != -1} {
          set app1 [join [$objDesign $attr]]

# for multiple files, modify the app string to fit in file summary worksheet
          if {$numFile != 0 && [info exists cells1(Summary)]} {
            if {$app == "Architectural Desktop"} {set app "ADT"}
            if {$app == "AutoCAD Architecture"}  {set app "ACA"}
            if {$app == "VectorWorks"}           {set app "Vector Works"}
            if {$app == "Triforma"}              {set app "Bentley"}
            if {$app == "Esa.Pt"}                {set app "Scia"}
            if {$app == "CIS/2 Trans.1.0a"}      {set app "Revit"}
            if {$app == "CASE Center"}           {set app "GTSTRUDL"}
            if {$app == "PRO/ENGINEER"}          {set app "Creo"}
            if {$app == "UGS - NX"}              {set app "UGS-NX"}
            if {$app == "Unigraphics"}           {set app "Siemens NX"}
            if {$app == "UNIGRAPHICS"}           {set app "Unigraphics"}
            if {$app == "SIEMENS PLM Software NX"} {set app "Siemens NX"}
            if {$app == "jt_step translator"}      {set app "Siemens NX"}
            if {$app == "Implementor Forum Team"}  {set app "CAx-IF"}
            if {$app == "3D_Evolution"}          {set app "CT 3D Evolution"}
            if {$app == "CoreTechnologie"}       {set app "CT 3D Evolution"}
            if {$app == "DATAKIT"}               {set app "Datakit"}
            if {$app == "Implementor Forum Team"} {set app "CAx-IF"}
            if {[string first "SIEMENS PLM Software NX" $app] == 0} {set app "Siemens NX[string range $app 23 end]"}
            if {[string first "CATIA Version"           $app] == 0} {set app "CATIA V[string range $app 14 end]"}

            if {$app == "Cis2Export_dll"} {
              set app "Revit"
            } elseif {$app == "CIS2export"} {
              set app "GTSTRUDL"
            }
            if {$app == "IFC text editor"}       {set app "Text editor"}
            if {$app == "Text editor" && [$objDesign FileOrganization] == "Constructivity"} {set app "Constructivity"}
            if {$app == "Text editor" && [$objDesign FileOrganization] == "AEC3"}           {set app "AEC3"}
            if {[string first "T-Systems" [$objDesign FilePreprocessorVersion]] != -1}      {set app "T-Systems"}
            if {[string first "THEOREM"   [$objDesign FilePreprocessorVersion]] != -1}      {set app "Theorem"}

# app > app2, put in spreadsheet
            if {$numFile != 0 && $opt(XLSCSV) == "Excel"} {
            regsub -all " " $app [format "%c" 10] app2
            if {$appnist && $app != "NIST"} {append app2 "[format "%c" 10](SteelVis)"}
            set colsum [expr {$col1(Summary)+1}]
            if {$colsum > 16} {[$excel1 ActiveWindow] ScrollColumn [expr {$colsum-16}]}
            set app2 [string trim $app2]
            $cells1(Summary) Item 6 $colsum $app2
            }
          }
          break
        }
      }
    }

# scroll file summary horizontally
    if {$appnist && $app1 == "" && [info exists cells1(Summary)] && $opt(XLSCSV) == "Excel"} {
      set app2 "NIST"
      set colsum [expr {$col1(Summary)+1}]
      if {$colsum > 16} {[$excel1 ActiveWindow] ScrollColumn [expr {$colsum-16}]}
      $cells1(Summary) Item 6 $colsum $app2
    }
    if {$appnist && $app1 != "NIST"} {append app1 " (SteelVis)"}

# close csv file
    if {$opt(XLSCSV) == "CSV"} {close $fcsv}

# print errors
  } emsg]} {
    errorMsg "ERROR adding Header worksheet: $emsg"
    catch {raise .}
  }

# -------------------------------------------------------------------------------------------------
# set Excel spreadsheet name, delete file if already exists

# user-defined file name
  if {$opt(XLSCSV) == "Excel"} {
    set xlsmsg ""
    if {$writeDirType == 1} {
      if {$userXLSFile != ""} {
        set xname [file nativename $userXLSFile]
      } else {
        append xlsmsg "User-defined Spreadsheet file name is not valid.  Spreadsheet directory and\n file name will be based on the IFC file. (Options tab)"
        set writeDirType 0
      }
    }

    set ifcstp "_ifc"

# same directory as file
    if {$writeDirType == 0} {
      set xname "[file nativename [file join [file dirname $fname] [file rootname [file tail $fname]]]]$ifcstp.$extXLS"

# user-defined directory
    } elseif {$writeDirType == 2} {
      set xname "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]$ifcstp.$extXLS"
    }

# file name too long
    if {[string length $xname] > 218} {
      if {[string length $xlsmsg] > 0} {append xlsmsg "\n\n"}
      append xlsmsg "Pathname of Spreadsheet file is too long for Excel ([string length $xname])"
      set xname "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]$ifcstp.$extXLS"
      if {[string length $xname] < 219} {
        append xlsmsg "\nSpreadsheet file written to User-defined directory (Spreadsheet tab)"
      }
    }

# delete existing file
    if {[file exists $xname]} {
      if {[catch {
        file delete -force $xname
      } emsg]} {
        if {[string length $xlsmsg] > 0} {append xlsmsg "\n"}
        append xlsmsg "Existing Spreadsheet will not be overwritten: [file tail $xname]"
        catch {raise .}
      }
    }
  }

# add file name to menu
  set ok 0
  if {$numFile <= 1} {set ok 1}
  if {[info exists localNameList]} {if {[llength $localNameList] > 1} {set ok 1}}
  if {$ok} {addFileToMenu}

# set types of entities to process
  set types {}
  foreach pr [array names type] {
    set ok 1
    if {[info exists opt($pr)] && $ok} {
      if {$opt($pr)} {set types [concat $types $type($pr)]}
    }
  }

# set entities to count
  set countEnts {}
  if {[info exists countent(IFC)]} {
    set countEnts $countent(IFC)
    if {!$opt(INVERSE)} {
      foreach typ $types {
        set ltyp [expr {[string length $typ]-4}]
        if {[string range $typ end-3 end] == "Type"} {lappend countEnts $typ}
      }
    }
  }

# do not count some entities if expanding some entities or reporting inverses
  if {$opt(EX_ANAL)} {
    set rmcount [list IfcEdge]
    set countEnts [lindex [intersect3 $countEnts $rmcount] 0]
  }
  if {$opt(INVERSE)} {
    set rmcount [list IfcRelAssociatesMaterial IfcRelAssociatesProfileProperties IfcRelConnectsPathElements IfcRelConnectsPorts IfcRelConnectsPortToElement IfcRelConnectsStructuralElement IfcRelDefinesByProperties IfcRelFillsElement IfcRelVoidsElement]
    set countEnts [lindex [intersect3 $countEnts $rmcount] 0]
  }

# -------------------------------------------------------------------------------------------------
# set which entities are processed and which are not
  set ws_proc  {}
  set ws_nproc {}
  set nent 0

# user-defined entity list
  catch {set userentlist {}}
  if {$opt(PR_USER) && [llength $userentlist] == 0 && [info exists userEntityFile]} {
    set userentlist {}
    set fileUserEnt [open $userEntityFile r]
    while {[gets $fileUserEnt line] != -1} {
      set line [split [string trim $line] " "]
      foreach ent $line {
        if {[lsearch $ifcall2x3 $ent] != -1 || \
            [lsearch $ifcall2x4 $ent] != -1} {lappend userentlist $ent}
      }
    }
    close $fileUserEnt
    if {[llength $userentlist] == 0} {
      set opt(PR_USER) 0
      checkValues
    }
  }

# get totals of each entity in file
  set rmcount {}
  set fixlist {}
  if {![info exists objDesign]} {return}

  foreach enttyp [$objDesign EntityTypeNames [expr 2]] {
    set ecount($enttyp) [$objDesign CountEntities "$enttyp"]

    if {$ecount($enttyp) > 0} {
      if {$numFile != 0} {
        set idx [setColorIndex $enttyp 1]
        if {$idx == -2} {set idx 99}
        lappend all_entity "$idx$enttyp"
        lappend file_entity($numFile) "$enttyp $ecount($enttyp)"
        if {![info exists total_entity($enttyp)]} {
          set total_entity($enttyp) $ecount($enttyp)
        } else {
          incr total_entity($enttyp) $ecount($enttyp)
        }
      }

# do not count entities if there are less than 100
      if {$ecount($enttyp) < 100} {lappend rmcount $enttyp}

# some general types of entities
      set ok 0

# user-defined entities
      if {$opt(PR_USER) && [lsearch $userentlist $enttyp] != -1} {set ok 1}

# IFC entities that are translated depending on the options
      set ok [ifcWhichEntities $ok $enttyp]

# handle '_and_' due to a complex entity, enttyp_1 is the first part before the '_and_'
      set enttyp_1 $enttyp
      set c1 [string first "_and_" $enttyp_1]
      if {$c1 != -1} {set enttyp_1 [string range $enttyp_1 0 $c1-1]}

# check for entities that cause crashes
      set nofix 1
      if {[info exists fixent]} {if {[lsearch $fixent $enttyp] != -1} {set nofix 0}}

# add to list of entities to process (ws_proc), uses color index to set the order
      set cidx [setColorIndex $enttyp 1]
      if {([lsearch $types $enttyp_1] != -1 || $ok)} {
        if {$nofix} {
          lappend ws_proc "$cidx$enttyp"
          incr nent $ecount($enttyp)
        } else {
          lappend fixlist $enttyp
          lappend ws_nproc $enttyp
          set ignored($cidx$enttyp) $ecount($enttyp)
        }
      } elseif {[lsearch $types $enttyp] != -1} {
        if {$nofix} {
          lappend ws_proc "$cidx$enttyp"
          incr nent $ecount($enttyp)
        } else {
          lappend fixlist $enttyp
          lappend ws_nproc $enttyp
          set ignored($cidx$enttyp) $ecount($enttyp)
        }
      } else {
        lappend ws_nproc $enttyp
        set ignored($cidx$enttyp) $ecount($enttyp)
      }
    }
  }

# list entities not processed based on fix file
  if {[llength $fixlist] > 0} {
    outputMsg " "
    if {[file exists $cfile]} {
      set ok 0
      foreach item $fixlist {if {[lsearch $fixprm $item] == -1} {set ok 1}}
      if {$ok} {errorMsg "Based on entities list in [truncFileName [file nativename $cfile]]"}
    }
    errorMsg "Worksheets will not be generated for the following entities:"
    foreach item [lsort $fixlist] {outputMsg " $item" red}
  }

# sort ws_proc by color index
  set ws_proc [lsort $ws_proc]

# then strip off the color index
  for {set i 0} {$i < [llength $ws_proc]} {incr i} {
    lset ws_proc $i [string range [lindex $ws_proc $i] 2 end]
  }

  if {[info exists buttons]} {$buttons(pgb) configure -maximum $nent}

# remove entities for list to count
  set countEnts [lindex [intersect3 $countEnts $rmcount] 0]

# -------------------------------------------------------------------------------------------------
# generate worksheet for each entity
  if {$opt(XLSCSV) == "Excel"} {
    outputMsg "\nGenerating IFC Entity worksheets" blue
  } else {
    outputMsg "\nGenerating IFC Entity CSV files" blue
  }
  if {[catch {
    set inverse_ent {}
    set last_ent ""
    set nline 0
    set nproc 0
    set nsheet 0
    set lastheading ""
    set stat 1
    set lastguid ""
    set ntable 0
    set icolor 0

    if {[llength $ws_proc] == 0} {
      errorMsg "No IFC entities were found in the file to Process as selected in the Options tab."
    }
    set tlast [clock clicks -milliseconds]

# loop over list of entities in file
    foreach enttyp $ws_proc {
      set nerr1 0
      set last_ent $enttyp

# decide if inverses should be checked for this entity type
      set checkInv 0
      if {$opt(INVERSE)} {set checkInv [invSetCheck $enttyp]}
      if {$checkInv} {lappend inverse_ent $enttyp}

      ::tcom::foreach objEntity [$objDesign FindObjects [join $enttyp]] {

# process the entity
        if {$enttyp == [$objEntity Type]} {
          incr nline
          if {[expr {$nline%1000}] == 0} {update idletasks}

          if {[catch {
            if {$opt(XLSCSV) == "Excel"} {
              set stat [getEntity $objEntity $enttyp $checkInv]
            } else {
              set stat [getEntityCSV $objEntity]
            }
          } emsg1]} {

# process errors with entity
            if {$stat != 1} {break}
            set msg "ERROR processing "
            if {[info exists objEntity]} {
              if {[string first "handle" $objEntity] != -1} {
                if {[$objEntity Type] != "IfcTrimmedCurve" && [$objEntity Type] != "trimmed_curve"} {
                  append msg "\#[$objEntity P21ID]=[$objEntity Type] (row [expr {$row($ifc)+2}]): $emsg1"

# handle specific errors
                  if {[string first "Unknown error" $emsg1] != -1} {
                    errorMsg $msg
                    catch {raise .}
                    incr nerr1
                    if {$nerr1 > 20} {
                      errorMsg "Processing of $enttyp entities has stopped" red
                      set nline [expr {$nline + $ecount($ifc) - $count($ifc)}]
                      break
                    }

                  } elseif {[string first "Insufficient memory to perform operation" $emsg1] != -1} {
                    errorMsg $msg
                    errorMsg "Several options are available to reduce memory usage:\nUse the option to limit the Maximum Rows"
                    if {$opt(COUNT)}   {errorMsg "Turn off Counting entities and process the file again" red}
                    if {$opt(INVERSE)} {errorMsg "Turn off Inverse Relationships and process the file again" red}
                    if {$opt(EX_LP)} {errorMsg "Turn off Expanding entities and process the file again" red}
                    catch {raise .}
                    break
                  }

# error message for IfcTrimmedCurve and trimmed_curve (causes problems for IFCsvr)
                } else {
                  append msg [$objEntity Type]
                }
                errorMsg $msg
                catch {raise .}
              }
            }
          }
          if {$stat != 1} {
            set nline [expr {$nline + $ecount($ifc) - $count($ifc)}]
            break
          }
        }
      }
      if {$opt(XLSCSV) == "CSV"} {catch {close $fcsv}}
    }

  } emsg2]} {
    set msg "ERROR processing IFC file: "
    if {[info exists objEntity]} {
      if {[string first "handle" $objEntity] != -1} {append msg " \#[$objEntity P21ID]=[$objEntity Type]"}
    }
    append msg " $emsg2"
    append msg "\nProcessing of the IFC file has stopped"
    errorMsg $msg red
    catch {raise .}
  }

# -------------------------------------------------------------------------------------------------
# check fix file
  if {[info exists cfile]} {
    set fixtmp {}
    if {[file exists $cfile]} {
      set fixfile [open $cfile r]
      while {[gets $fixfile line] >= 0} {
        if {[lsearch $fixtmp $line] == -1 && $line != $last_ent} {lappend fixtmp $line}
      }
      close $fixfile
    }

    if {[join $fixtmp] == ""} {
      catch {file delete -force $cfile}
    } else {
      set fixfile [open $cfile w]
      foreach item $fixtmp {puts $fixfile $item}
      close $fixfile
    }
  }

# -------------------------------------------------------------------------------------------------
# quit IFCsvr, but not sure how to do it properly
  if {[catch {
    $objDesign Delete
    unset objDesign
    unset objIFCsvr

# print errors
  } emsg]} {
    errorMsg "ERROR closing IFCsvr: $emsg"
    catch {raise .}
  }

# -------------------------------------------------------------------------------------------------
# add summary worksheet
  if {$opt(XLSCSV) == "Excel"} {
    outputMsg "\nGenerating Summary worksheet" blue
    set sum "Summary"

    set ws_sort {}
    foreach enttyp [lsort [array names worksheet]] {
      if {[string range $enttyp 0 2] == "Ifc" && $enttyp != "Summary" && $enttyp != "Header"} {
        lappend ws_sort "[setColorIndex $enttyp 1]$enttyp"
      }
    }
    set ws_sort [lsort $ws_sort]
    for {set i 0} {$i < [llength $ws_sort]} {incr i} {
      lset ws_sort $i [string range [lindex $ws_sort $i] 2 end]
    }

    if {[catch {
      set worksheet($sum) [$worksheets Add [::tcom::na] $ws_last]
      $worksheet($sum) Activate
      $worksheet($sum) Name $sum
      set cells($sum) [$worksheet($sum) Cells]
      $cells($sum) Item 1 1 "Entity"
      $cells($sum) Item 1 2 "Count"
      set ncol 2
      set attrcol {}
      if {[info exists attrused]} {
        foreach attr $attrsum {
          if {[lsearch $attrused $attr] != -1} {
            incr ncol
            $cells($sum) Item 1 $ncol $attr
            lappend attrcol $attr
          }
        }
      }
      set col($sum) $ncol
      set hlsum [$worksheet($sum) Hyperlinks]

      set nsheet [$worksheets Count]
      [$worksheets Item [expr $nsheet]] -namedarg Move Before [$worksheets Item [expr 1]]

# Summary of entities in column 1 and count in column 2
      set vlink 1
      set row($sum) 1
      foreach enttyp $ws_sort {
        incr row($sum)
        set rws [expr {[lsearch $ws_sort $enttyp]+2}]

# check if entity is compound as opposed to an entity with '_and_'
        set ok 0
        if {[string first "_and_" $enttyp] == -1} {
          set ok 1
        } else {
          foreach item [array names type] {if {[lsearch $type($item) $enttyp] != -1} {set ok 1}}
        }
        if {$ok} {
          $cells($sum) Item $rws 1 $enttyp

# for '_and_' (complex entity) split on multiple lines
# '10' is the ascii character for a linefeed
        } else {
          regsub -all "_and_" $enttyp ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" enttyp_multiline
          set enttyp_multiline "($enttyp_multiline)"
          $cells($sum) Item $rws 1 $enttyp_multiline

          set range [$worksheet($sum) Range $rws:$rws]
          $range VerticalAlignment [expr -4108]
        }

# entity count in column 2
        $cells($sum) Item $rws 2 $ecount($enttyp)

# attribute counts
        foreach attr $attrcol {
          if {[info exists count($enttyp,$attr)]} {
            set ncol [expr {[lsearch $attrcol $attr] + 3}]
            $cells($sum) Item $rws $ncol $count($enttyp,$attr)
          }
        }
      }

# headings for IFC documentation
      switch -- [getSchema $fname] {
        IFC4 {
          $cells($sum) Item 1 [incr col($sum)] "IFC4"
          set anchor [$worksheet($sum) Range [cellRange 1 $col($sum)] [cellRange 1 $col($sum)]]
          [$worksheet($sum) Hyperlinks] Add $anchor "https://standards.buildingsmart.org/IFC/RELEASE/IFC4/FINAL/HTML/" [join ""] [join "IFC4 Documentation"]
        }
        default {
          $cells($sum) Item 1 [incr col($sum)] "IFC2x3"
          set anchor [$worksheet($sum) Range [cellRange 1 $col($sum)] [cellRange 1 $col($sum)]]
          [$worksheet($sum) Hyperlinks] Add $anchor "https://standards.buildingsmart.org/IFC/RELEASE/IFC2x3/TC1/HTML/" [join ""] [join "IFC2x3 Documentation"]
        }
      }
      set doccol 3

# entities not processed
      set rowig [expr {[array size worksheet]+2}]
      $cells($sum) Item $rowig 1 "Entity types not processed ([array size ignored])"

      foreach ent [lsort [array names ignored]] {
        set ent0 [string range $ent 2 end]
        set ok 0
        if {[string first "_and_" $ent] == -1} {
          set ok 1
        } else {
          foreach item [array names type] {if {[lsearch $type($item) $ent0] != -1} {set ok 1}}
        }
        if {$ok} {
          $cells($sum) Item [incr rowig] 1 $ent0
        } else {
# '10' is the ascii character for a linefeed
          regsub -all "_and_" $ent0 ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" ent1
          $cells($sum) Item [incr rowig] 1 "($ent1)"
          set range [$worksheet($sum) Range $rowig:$rowig]
          $range VerticalAlignment [expr -4108]
        }
        $cells($sum) Item $rowig 2 $ignored($ent)
      }
      set row($sum) $rowig
      [$excel ActiveWindow] ScrollRow [expr 1]

# autoformat entire summary worksheet
      set range [$worksheet($sum) Range [cellRange 1 1] [cellRange $row($sum) $col($sum)]]
      $range AutoFormat

# name and link to program website that generated the spreadsheet
      $cells($sum) Item [expr {$row($sum)+2}] 1 "Spreadsheet generated by the NIST IFC File Analyzer (v[getVersion])"
      set anchor [$worksheet($sum) Range [cellRange [expr {$row($sum)+2}] 1]]
      [$worksheet($sum) Hyperlinks] Add $anchor [join "https://www.nist.gov/services-resources/software/ifc-file-analyzer"] [join ""] [join "Link to IFC File Analyzer"]
      $cells($sum) Item [expr {$row($sum)+3}] 1 "[clock format [clock seconds]]"

# print errors
    } emsg]} {
      errorMsg "ERROR adding Summary worksheet: $emsg"
      catch {raise .}
    }

# -------------------------------------------------------------------------------------------------
# format cells on each entity worksheet
    outputMsg "Formatting Worksheets"
    set cellcolors [list 36 35 34 37 39 38 40 24 19 44 45]

    if {[info exists buttons]} {$buttons(pgb) configure -maximum [llength $ws_sort]}
    set nline 0
    set nsort 0
    foreach ifc $ws_sort {
      incr nline
      update idletasks

      if {[catch {
        set counting 0
        if {$opt(COUNT) && [lsearch $countEnts $ifc] != -1} {
          set counting 1
          incr col($ifc)
        }

        $worksheet($ifc) Activate
        [$excel ActiveWindow] ScrollRow [expr 1]

# find extent of columns
        set rancol $col($ifc)
        for {set i 1} {$i < 10} {incr i} {
          if {[[$cells($ifc) Item 3 [expr {$col($ifc)+$i}]] Value] != ""} {
            incr rancol
          } else {
            break
          }
        }

# find extent of rows
        set ranrow [expr {$row($ifc)+2}]
        if {$ranrow > $rowmax} {set ranrow [expr {$rowmax+2}]}
        set ranrow [expr {$ranrow-2}]

# autoformat
        set range [$worksheet($ifc) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
        $range AutoFormat

# freeze panes
        [$worksheet($ifc) Range "B4"] Select
        [$excel ActiveWindow] FreezePanes [expr 1]

# if counting, blank this cell
        if {$counting} {$cells($ifc) Item [expr {$ranrow+1}] 1 " "}

# set A4 as default cell
        [$worksheet($ifc) Range "A4"] Select

# -------------------------------------------------------------------------------------------------
# set column color for expanded entities, depends on colclr variable
        set clr($ifc) 0
        if {[info exists colclr($ifc)]} {
          set grp1 [list [lindex [lindex $colclr($ifc) 0]   1]]
          set grp2 [list [lindex [lindex $colclr($ifc) end] 1]]

          foreach item $colclr($ifc) {
            incr clr($ifc) [lindex $item 0]
            if {$clr($ifc) < 0 || $clr($ifc) >= [llength $cellcolors]} {errorMsg "Color index out of range Expanding: $ifc"}

            set gc [lindex $item 1]
            set r1 [expr {$row($ifc)+2}]
            if {$r1 > $rowmax} {set r1 [expr {$r1-1}]}
            set r1 [expr {$r1-2}]
            set range [$worksheet($ifc) Range [cellRange 3 $gc] [cellRange $r1 $gc]]
            [$range Interior] ColorIndex [expr [lindex $cellcolors $clr($ifc)]]

            if {[expr {int([$excel Version])}] >= 12} {
              set range [$worksheet($ifc) Range [cellRange 4 $gc] [cellRange $r1 $gc]]
              for {set k 7} {$k <= 12} {incr k} {
                if {$k != 9} {
                  catch {[[$range Borders] Item [expr $k]] Weight [expr 1]}
                }
              }
              set range [$worksheet($ifc) Range [cellRange 3 $gc] [cellRange 3 $gc]]
              catch {
                [[$range Borders] Item [expr 7]]  Weight [expr 1]
                [[$range Borders] Item [expr 10]] Weight [expr 1]
              }
            }

            if {[lindex $item 0] == 1} {lappend grp1 [lindex $item 1]}
          }
          for {set i [expr {[llength $colclr($ifc)]-1}]} {$i >= 0} {incr i -1} {
            if {[lindex [lindex $colclr($ifc) $i] 0] == -1} {lappend grp2 [lindex [lindex $colclr($ifc) [expr {$i-1}]] 1]}
          }

          if {$ifc == "IfcExtrudedAreaSolid"} {
            set grp1 {4 8}
            set grp2 {6 8}
          } elseif {$ifc == "IfcGeometricRepresentationContext"} {
            set grp1 {7 11}
            set grp2 {9 11}
          } elseif {$ifc == "IfcEdge"} {
            set grp1 {3 5}
            set grp2 {3 5}
          } elseif {$ifc == "IfcStructuralLinearAction"  || $ifc == "IfcStructuralPointAction"   || \
                    $ifc == "IfcStructuralPointReaction" || $ifc == "IfcStructuralCurveReaction" || \
                    $ifc == "IfcStructuralSurfaceReaction"} {
            set grp1 {6  13}
            set grp2 {10 19}
          }

          if {[llength $grp1] > 0} {
            for {set i [expr {[llength $grp1]-1}]} {$i >= 0} {incr i -1} {
              set grange [$worksheet($ifc) Range [cellRange 1 [lindex $grp1 $i]] [cellRange [expr {$row($ifc)+2}] [lindex $grp2 $i]]]
              [$grange Columns] Group
            }
          }
        }

# set column color for count, if counting entities
        if {$counting} {
          for {set ic 100} {$ic > 2} {incr ic -1} {
            set range [$worksheet($ifc) Range [cellRange 3 $ic] [cellRange 3 $ic]]
            if {[$range Value] == "Count"} {
              set crange [$worksheet($ifc) Range [cellRange 3 $ic] [cellRange $ranrow $ic]]
              [$crange Interior] ColorIndex [expr 19]
              set range  [$worksheet($ifc) Range [cellRange 4 $ic] [cellRange $ranrow $ic]]
              if {[expr {int([$excel Version])}] >= 12} {
                for {set k 7} {$k <= 12} {incr k} {
                  catch {if {$k != 9} {[[$range Borders] Item [expr $k]] Weight [expr 1]}}
                }
              }
              break
            }
          }
          set row3 [expr {$row($ifc)+3}]
          if {$row3 > $ranrow} {
            $cells($ifc) Item $row3 1 ""
            $cells($ifc) Item $row3 2 ""
          }
        }

# set column color, border, group for INVERSES and Used In
        if {$opt(INVERSE)} {if {[lsearch $inverse_ent $ifc] != -1} {invFormat $rancol}}

# -------------------------------------------------------------------------------------------------
# link back to summary on entity worksheets
        set hlink [$worksheet($ifc) Hyperlinks]
        set txt "[formatComplexEnt $ifc]  "
        set row1 [expr {$row($ifc)-3}]
        if {$row1 == $count($ifc) && $row1 == $ecount($ifc)} {
          append txt "($row1)"
        } elseif {$row1 > $count($ifc) && $count($ifc) < $ecount($ifc)} {
          append txt "($count($ifc) of $ecount($ifc))"
        } elseif {$row1 < $ecount($ifc)} {
          if {$count($ifc) == $ecount($ifc)} {
            append txt "($row1 of $ecount($ifc))"
          } else {
            append txt "([expr {$row1-3}] of $count($ifc))"
          }
        }
        $cells($ifc) Item 1 1 $txt
        set range [$worksheet($ifc) Range "A1:H1"]
        $range MergeCells [expr 1]

# link back to summary
        if {$opt(XL_LINK1)} {
          set anchor [$worksheet($ifc) Range "A1"]
          $hlink Add $anchor $xname "Summary!A$rws" "Return to Summary"

# links to documenation on entity worksheet
          entDocLink $ifc $ifc 2 1 $hlink
        }

# check width of columns, wrap text
        if {[catch {
          set widlim 400.
          for {set i 2} {$i <= $rancol} {incr i} {
            if {[[$cells($ifc) Item 3 $i] Value] != ""} {
              set wid [[$cells($ifc) Item 3 $i] Width]
              if {$wid > $widlim} {
                set range [$worksheet($ifc) Range [cellRange -1 $i]]
                $range ColumnWidth [expr {[$range ColumnWidth]/$wid * $widlim}]
                $range WrapText [expr 1]
              }
            }
          }
        } emsg]} {
          errorMsg "ERROR setting column widths: $emsg\n  $ifc"
          catch {raise .}
        }

# -------------------------------------------------------------------------------------------------
# IfcPropertySet link to IFC documentation
        ifcFormatPropertySet

# add table for sorting and filtering
        if {[expr {int([$excel Version])}] >= 12} {
          if {[catch {
            if {$opt(SORT)} {
              if {$ranrow > 5} {
                set range [$worksheet($ifc) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
                set tname [string trim "TABLE-$ifc"]
                [[$worksheet($ifc) ListObjects] Add 1 $range] Name $tname
                [[$worksheet($ifc) ListObjects] Item $tname] TableStyle "TableStyleLight1"
                if {[incr ntable] == 1} {outputMsg " Generating Tables for Sorting" blue}
              }
            }
          } emsg]} {
            errorMsg "ERROR adding Tables for Sorting: $emsg"
            catch {raise .}
          }
        }

# errors
      } emsg]} {
        errorMsg "ERROR formatting Spreadsheet for: $ifc\n$emsg"
        catch {raise .}
      }
    }

    incr col($sum) -3

# -------------------------------------------------------------------------------------------------
# add file name and other info to top of Summary

    set nhrow 0
    if {[catch {
      $worksheet($sum) Activate
      [$worksheet($sum) Range "1:1"] Insert
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Total Entities"
      $cells($sum) Item 1 2 "'$entityCount"
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr nhrow

      if {$timestamp != ""} {
        [$worksheet($sum) Range "1:1"] Insert
        $cells($sum) Item 1 1 "Timestamp"
        $cells($sum) Item 1 2 [join $timestamp]
        set range [$worksheet($sum) Range "B1:K1"]
        $range MergeCells [expr 1]
        incr nhrow
      }

      if {$app1 == "" && [info exists ifcApplication]} {set app1 $ifcApplication}
      if {$app1 != ""} {
        [$worksheet($sum) Range "1:1"] Insert
        $cells($sum) Item 1 1 "Application"
        $cells($sum) Item 1 2 [join $app1]
        set range [$worksheet($sum) Range "B1:K1"]
        $range MergeCells [expr 1]
        incr nhrow
      }

      if {!$opt(COUNT) || $writeDirType != 0} {
        [$worksheet($sum) Range "1:1"] Insert
        $cells($sum) Item 1 1 "Excel File"
        $cells($sum) Item 1 2 [truncFileName $xname]
        set range [$worksheet($sum) Range "B1:K1"]
        $range MergeCells [expr 1]
        incr nhrow
      }

      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "IFC Directory"
      $cells($sum) Item 1 2 [file nativename [file dirname [truncFileName $localName]]]
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr nhrow

      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "IFC File"
      $cells($sum) Item 1 2 [file tail $localName]
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      set anchor [$worksheet($sum) Range "B1"]
      if {$opt(XL_LINK1)} {$hlsum Add $anchor [join $localName] [join ""] [join "Link to IFC file"]}
      incr nhrow

      set range [$worksheet($sum) Range [cellRange 1 1] [cellRange $nhrow 1]]
      [$range Font] Bold [expr 1]

    } emsg]} {
      errorMsg "ERROR adding File Names to Summary: $emsg"
      catch {raise .}
    }

# freeze panes
    [$worksheet($sum) Range "B[expr {$nhrow+3}]"] Select
    [$excel ActiveWindow] FreezePanes [expr 1]
    [$worksheet($sum) Range "A1"] Select

# -------------------------------------------------------------------------------------------------
# add Summary color and hyperlinks
    if {[catch {
      outputMsg " Adding links on Summary to IFC documentation" blue
      set row($sum) [expr {$nhrow+2}]
      set nline 0

      foreach ifc $ws_sort {
        incr nline
        update idletasks

        incr row($sum)
        set nrow [expr {20-$nhrow}]
        if {$row($sum) > $nrow} {[$excel ActiveWindow] ScrollRow [expr {$row($sum)-$nrow}]}

        set rws [expr {[lsearch $ws_sort $ifc]+3+$nhrow}]

# link from summary to entity worksheet
        set anchor [$worksheet($sum) Range "A$rws"]
        set hlsheet $ifc
        if {[string length $ifc] > 31} {
          foreach item [array names entName] {
            if {$entName($item) == $ifc} {set hlsheet $item}
          }
        }
        $hlsum Add $anchor $xname "$hlsheet!A4" "Go to $ifc"

# color entities on summary
        set cidx [setColorIndex $ifc 1]
        if {$cidx > 0} {
          [$anchor Interior] ColorIndex [expr $cidx]
        }

# bold entities for reports
        if {[string first "\[" [$anchor Value]] != -1} {[$anchor Font] Bold [expr 1]}

# add link to entity documentation
        set ncol [expr {$col($sum)-1}]
        entDocLink $sum $ifc $rws $doccol $hlsum
      }

# add links for ignored entities, find row where they start
      set i1 [expr {max([array size worksheet],9)}]
      for {set i $i1} {$i < 1000} {incr i} {
        if {[string first "Entity" [[$cells($sum) Item $i 1] Value]] == 0} {
          set rowig $i
          break
        }
      }
      set range [$worksheet($sum) Range "A$rowig"]
      [$range Font] Bold [expr 1]

      set i1 3
      set range [$worksheet($sum) Range [cellRange $rowig 1] [cellRange $rowig [expr {$col($sum)+$i1}]]]
      catch {[[$range Borders] Item [expr 8]] Weight [expr -4138]}

      foreach ent [lsort [array names ignored]] {
        incr rowig
        set nrow [expr {20-$nhrow}]
        if {$rowig > $nrow} {[$excel ActiveWindow] ScrollRow [expr {$rowig-$nrow}]}
        set ncol [expr {$col($sum)-1}]
        entDocLink $sum $ent $rowig $doccol $hlsum

        set range [$worksheet($sum) Range [cellRange $rowig 1]]
        set cidx [setColorIndex [string range $ent 2 end] 1]
        if {$cidx > 0} {[$range Interior] ColorIndex [expr $cidx]}
      }
      [$worksheet($sum) Columns] AutoFit
      [$worksheet($sum) Rows] AutoFit
      catch {[$worksheet($sum) PageSetup] PrintGridlines [expr 1]}

    } emsg]} {
      errorMsg "ERROR adding Summary links: $emsg"
      catch {raise .}
    }

# select the first tab
    [$worksheets Item [expr 1]] Select
    [$excel ActiveWindow] ScrollRow [expr 1]

    set proctime [expr {([clock clicks -milliseconds] - $lasttime)/1000}]
    outputMsg "Processing time: $proctime seconds"
  }

# -------------------------------------------------------------------------------------------------
# save spreadsheet
  if {$opt(XLSCSV) == "Excel"} {
    if {[catch {
      outputMsg " "
      if {$xlsmsg != ""} {outputMsg $xlsmsg red}
      if {[string first "\[" $xname] != -1} {
        regsub -all {\[} $xname "(" xname
        regsub -all {\]} $xname ")" xname
        outputMsg "In the spreadsheet file name, the characters \'\[\' and \'\]\' have been\n substituted by \'\(\' and \'\)\'" red
      }
      set xlfn $xname

# create new file name if spreadsheet already exists, delete new file name spreadsheets if possible
      if {[file exists $xlfn]} {set xlfn [incrFileName $xlfn]}

      outputMsg "Saving Spreadsheet as:"
      outputMsg " [truncFileName $xlfn 1]" blue
      catch {$excel DisplayAlerts False}
      $workbook -namedarg SaveAs Filename $xlfn FileFormat $xlFormat
      catch {$excel DisplayAlerts True}
      set lastXLS $xlfn
      lappend xnames $xlfn

      catch {$excel ScreenUpdating 1}

# close Excel
      $excel Quit
      if {[info exists excel]} {unset excel}
      set openxl 1
      if {[llength $pidexcel] == 1} {
        catch {twapi::end_process $pidexcel -force}
      } else {
        errorMsg " Excel might not have been closed" red
      }
      update idletasks

# add Link(n) text to multi file summary
      if {$numFile != 0 && [info exists cells1(Summary)]} {
        set colsum [expr {$col1(Summary)+1}]
        if {$opt(XL_LINK1)} {
          $cells1(Summary) Item 3 $colsum "Link ($numFile)"
          set range [$worksheet1(Summary) Range [cellRange 3 $colsum]]
          regsub -all {\\} $xname "/" xls
          [$worksheet1(Summary) Hyperlinks] Add $range [join $xls] [join ""] [join "Link to Spreadsheet"]
        } else {
          $cells1(Summary) Item 3 $colsum "$numFile"
        }
      }

# errors
    } emsg]} {
      errorMsg "ERROR saving Spreadsheet: $emsg"
      if {[string first "The file name or path does not exist" $emsg]} {
        outputMsg " "
        errorMsg "Either copy the IFC file to a different directory and try generating the\n spreadsheet again or use the option to write the spreadsheet to a user-defined\n directory (Spreadsheet tab)."
      }
      catch {raise .}
      set openxl 0
    }

# -------------------------------------------------------------------------------------------------
# open spreadsheet
    set ok 0
    if {$openxl && $opt(XL_OPEN)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {
      openXLS $xlfn
    } elseif {!$opt(XL_OPEN) && $numFile == 0 && [string first "IFC-File-Analyzer.exe" $scriptName] != -1} {
      outputMsg " Use F2 to open the Spreadsheet (see Options tab)" red
    }

# open directory of CSV files
  } else {
    unset csvfile
    outputMsg "\nCSV files written to:"
    outputMsg " [truncFileName [file nativename $csvdirnam]]" blue
    set ok 0
    if {$opt(XL_OPEN)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {
      outputMsg "\nOpening CSV file directory" blue
      set dir [file nativename $csvdirnam]
      if {[string first " " $dir] == -1} {
        if {[catch {
          exec {*}[auto_execok start] $dir
        } emsg]} {
          if {[string first "UNC" $emsg] == -1} {errorMsg "ERROR opening CSV file directory: $emsg"}
        }
      } else {
        exec C:/Windows/explorer.exe $dir &
      }
    }
  }

# -------------------------------------------------------------------------------------------------
# save state
  if {[info exists errmsg]} {unset errmsg}
  saveState
  if {!$multiFile && [info exists buttons]} {$buttons(genExcel) configure -state normal}
  update idletasks

# clean up variables to hopefully release some memory and/or to reset them
  global nrep invGroup
  foreach var {attrused colclr count ignored pcount pcountRow colinv \
               worksheet worksheets workbook workbooks cells \
               heading entName lpnest \
               nrep invGroup} {
    if {[info exists $var]} {unset $var}
  }
  update idletasks
  return 1
}
