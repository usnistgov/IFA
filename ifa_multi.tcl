# process multiple files in a directory
proc openMultiFile {{ask 1}} {
  global all_entity buttons cells1 col1 excel1 extXLS file_entity fileDir fileDir1 fileList lastXLS1 lenfilelist localName localNameList
  global multiFileDir mydocs nprogfile opt row1 startrow total_entity type worksheet1 worksheets1 xlFormat xnames

# select directory of files (default)
  if {$ask == 1} {
    if {![file exists $fileDir1] && [info exists mydocs]} {set fileDir1 $mydocs}
    set multiFileDir [tk_chooseDirectory -title "Select Directory of IFC Files" \
                -mustexist true -initialdir $fileDir1]
    if {[info exists localNameList]} {unset localNameList}

# list of files
  } elseif {$ask == 2} {
    set multiFileDir [file dirname [lindex $localNameList 0]]
    set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
    set fileList [lsort -dictionary $localNameList]
    set lenfilelist [llength $localNameList]

# don't ask for F4
  } elseif {$ask == 0} {
    set multiFileDir $fileDir1
    if {[info exists localNameList]} {unset localNameList}
  }

  if {$multiFileDir != "" && [file isdirectory $multiFileDir]} {
    if {$ask != 2} {
      outputMsg "\nIFC file directory: [truncFileName [file nativename $multiFileDir]]" blue
      set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
      .tnb select .tnb.status
      update idletasks
      set fileDir1 $multiFileDir
      saveState

      set recurse 0
      if {$ask} {
        set choice [tk_messageBox -title "Search Subdirectories?" -type yesno -default yes \
                    -message "Do you want to process IFC files in subdirectories too?" -icon question]
      } else {
        set choice "yes"
      }
      if {$choice == "yes"} {
        set recurse 1
        outputMsg "Searching subdirectories ..." blue
      }

# find all files in directory and subdirectories
      set fileList {}
      findFile $multiFileDir $recurse
      set lenfilelist [llength $fileList]

# list files and size
      foreach file1 $fileList {
        outputMsg "  [string range [file nativename [truncFileName $file1]] $dlen end]  ([fileSize $file1])"
      }
    }

    if {$lenfilelist > 0} {
      if {$ask != 2} {outputMsg "($lenfilelist) IFC files found" green}
      set askstr "Spreadsheets"
      if {$opt(XLSCSV) == "CSV"} {set askstr "CSV files"}

      if {$ask != 2} {
        set choice [tk_messageBox -title "Generate $askstr?" -type yesno -default yes -message "Do you want to Generate $askstr for ($lenfilelist) IFC files ?" -icon question]
      } else {
        set choice "yes"
      }

      if {$choice == "yes"} {
        checkForExcel
        set lasttime [clock clicks -milliseconds]

# save some variables
        checkValues

# start Excel for summary of all files
        if {$lenfilelist > 1 && $opt(XLSCSV) == "Excel"} {
          set fileDir  $multiFileDir
          if {[catch {
            set pid2 [twapi::get_process_ids -name "EXCEL.EXE"]
            set excel1 [::tcom::ref createobject Excel.Application]
            set pidexcel1 [lindex [intersect3 $pid2 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]

            set mf [expr {2**14}]
            set extXLS "xlsx"
            set xlFormat [expr 51]

            set mf [expr {$mf-3}]
            if {$lenfilelist > $mf} {
              errorMsg "Only the first $mf files will be processed due to column limits in Excel."
              set lenfilelist $mf
              set fileList [lrange $fileList 0 [expr {$mf-1}]]
            }
            outputMsg "Starting File Summary spreadsheet" blue
            $excel1 Visible 1

# errors
          } emsg]} {
            errorMsg "ERROR connecting to Excel: $emsg"
          }

# start summary/analysis spreadsheet
          if {[catch {
            set workbooks1  [$excel1 Workbooks]
            set workbook1   [$workbooks1 Add]
            set worksheets1 [$workbook1 Worksheets]

# determine how many worksheets to add for coverage analysis
            set n1 1

# delete 0, 1, or 2 worksheets
            catch {$excel1 DisplayAlerts False}
            set sheetCount [$worksheets1 Count]
            for {set n $sheetCount} {$n > $n1} {incr n -1} {[$worksheets1 Item [expr $n]] Delete}
            catch {$excel1 DisplayAlerts True}

# done starting coverage analysis worksheets
# -------------------------------------------------------------------------

# start file summary worksheet
            set sum "Summary"
            set worksheet1($sum) [$worksheets1 Item [expr 1]]
            $worksheet1($sum) Activate
            $worksheet1($sum) Name "File Summary"
            set cells1($sum) [$worksheet1($sum) Cells]
            $cells1($sum) Item 1 1 "IFC Directory"
            set range [$worksheet1($sum) Range [cellRange 1 2]]
            $cells1($sum) Item 1 2 "[file nativename $multiFileDir]"

# set startrow
            set startrow 8
            $cells1($sum) Item $startrow 1 "Entity"
            set range [$worksheet1($sum) Range "B1:K1"]
            [$range Font] Bold [expr 1]
            $range MergeCells [expr 1]
            set col1($sum) 1

# orientation for file info
            set range [$worksheet1($sum) Range "5:$startrow"]
            $range VerticalAlignment [expr -4107]
            $range HorizontalAlignment [expr -4108]

# vertical orientation for file name
            set range [$worksheet1($sum) Range "4:4"]
            $range Orientation [expr 90]
            $range HorizontalAlignment [expr -4108]

            [$excel1 ActiveWindow] TabRatio [expr 0.3]

# errors
          } emsg]} {
            errorMsg "ERROR opening Excel workbooks and worksheets for file summary: $emsg"
            catch {raise .}
          }
        } elseif {$opt(XLSCSV) == "Excel"} {
          errorMsg "For only one IFC file, no File Summary spreadsheet is generated."
        }

# -------------------------------------------------------------------------------------------------
# loop over all the files and process
        if {[info exists file_entity]}   {unset file_entity}
        if {[info exists total_entity]}  {unset total_entity}
        set xnames {}
        set all_entity {}
        set dirchange {}
        set lastdirname ""
        set nfile 0
        set sum "Summary"

        set nprogfile 0
        pack $buttons(pgb1) -side top -padx 10 -pady {5 0} -expand true -fill x
        $buttons(pgb1) configure -maximum $lenfilelist

# start loop
        foreach file1 $fileList {
          incr nfile

# process the file
          set stat($nfile) 0
          set localName $file1
          outputMsg "\n-------------------------------------------------------------------------------"
          outputMsg "($nfile of $lenfilelist) Ready to process: [file tail $file1] ([fileSize $file1])" green

          if {[catch {
            set stat($nfile) [genExcel $nfile]

# error processing the file
          } emsg]} {
            errorMsg "ERROR processing [file tail $file1]: $emsg"
            catch {raise .}
            set stat($nfile) 0
          }

# set fn from file name (file1), change \ to linefeed
          if {$lenfilelist > 1} {
            set fn [string range [file nativename [truncFileName $file1]] $dlen end]
            regsub -all {\\} $fn [format "%c" 10] fn
            incr col1($sum)

# keep track of changes in directory name to have vertical line when directory changes
            set dirname [file dirname $file1]
            if {$lastdirname != "" && $dirname != $lastdirname} {lappend dirchange [expr {$nfile+1}]}
            set lastdirname $dirname

# done adding coverage analysis results
# -------------------------------------------------------------------------
          }
          incr nprogfile
        }

# -------------------------------------------------------------------------------------------------
# file summary ws, entity names
        if {$lenfilelist > 1 && $opt(XLSCSV) == "Excel"} {
          catch {$excel1 ScreenUpdating 0}
          if {[catch {
            [$excel1 ActiveWindow] ScrollColumn [expr 1]
            set wid 2
            if {$lenfilelist > 16} {set wid 3}
            if {$lenfilelist > 1}  {incr wid}

            set row1($sum) $startrow
            set all_entity [lsort [lrmdups $all_entity]]
            set links [$worksheet1($sum) Hyperlinks]
            set inc1 0

# entity names, split on _and_
            for {set i 0} {$i < [llength $all_entity]} {incr i} {
              set ent [string range [lindex $all_entity $i] 2 end]
              lset all_entity $i $ent
              set ok 0
              if {[string first "_and_" $ent] == -1} {
                set ok 1
              } else {
                foreach item [array names type] {if {[lsearch $type($item) $ent] != -1} {set ok 1}}
              }
              if {$ok} {
                $cells1($sum) Item [incr row1($sum)] 1 $ent
                set ent2 $ent
              } else {
# '10' is the ascii character for a linefeed
                regsub -all "_and_" $ent ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" ent1
                $cells1($sum) Item [incr row1($sum)] 1 "($ent1)"
                set range [$worksheet1($sum) Range $row1($sum):$row1($sum)]
                $range VerticalAlignment [expr -4108]
                set ent2 "($ent1)"
              }
              set entrow($ent) $row1($sum)

# if more than 16 file, repeat entity names on the right
              if {$lenfilelist > 16} {
                $cells1($sum) Item $row1($sum) [expr {$lenfilelist+$wid+$inc1}] $ent2
              }
            }

#-------------------------------------------------------------------------------
# fix wrap for vertical file names
            set range [$worksheet1($sum) Range "4:4"]
            $range WrapText [expr 0]
            $range WrapText [expr 1]

# format file summary worksheet
            set range [$worksheet1($sum) Range [cellRange 3 1] [cellRange [expr {[llength $all_entity]+$startrow}] [expr {$lenfilelist+$wid+$inc1}]]]
            $range AutoFormat
            set range [$worksheet1($sum) Range "5:$startrow"]
            $range VerticalAlignment [expr -4107]
            $range HorizontalAlignment [expr -4108]

# erase some extra horizontal lines that AutoFormat created, but it doesn't work
            if {$lenfilelist > 2} {
              set range [$worksheet1($sum) Range [cellRange [expr {$startrow-1}] 1] [cellRange [expr {$startrow-1}] [expr {$lenfilelist+$wid+$inc1}]]]
              set borders [$range Borders]
              catch {
                [$borders Item [expr 8]] Weight [expr -4142]
                [$borders Item [expr 9]] Weight [expr -4142]
              }
            }

# generated by
            set c [expr {[llength $all_entity]+$startrow+2}]
            $cells1($sum) Item $c 1 "Spreadsheet generated by the NIST IFC File Analyzer (v[getVersion])"
            set anchor [$worksheet1($sum) Range [cellRange $c 1]]
            [$worksheet1($sum) Hyperlinks] Add $anchor [join "https://www.nist.gov/services-resources/software/ifc-file-analyzer"] [join ""] [join "Link to IFC File Analyzer"]
            $cells1($sum) Item [expr {[llength $all_entity]+$startrow+3}] 1 "[clock format [clock seconds]]"

# entity counts
            if {[info exists infiles]} {unset infiles}
            foreach idx [lsort -integer [array names file_entity]] {
              set col1($sum) [expr {$idx+1}]
              set scrollcol 16
              if {$col1($sum) > $scrollcol} {[$excel1 ActiveWindow] ScrollColumn [expr {$col1($sum)-$scrollcol}]}
              foreach item $file_entity($idx) {
                set val [split $item " "]
                $cells1($sum) Item $entrow([lindex $val 0]) $col1($sum) [lindex $val 1]
                incr infiles($entrow([lindex $val 0]))
              }
            }

# entity totals
            set col1($sum) [expr {$lenfilelist+2}]
            $cells1($sum) Item $startrow $col1($sum) "Total[format "%c" 10]Entities"
            foreach idx [array names total_entity] {
              $cells1($sum) Item $entrow($idx) $col1($sum) $total_entity($idx)
            }

# file occurances
            if {$lenfilelist > 1} {
              $cells1($sum) Item $startrow [incr col1($sum)] "Total[format "%c" 10]Files"
            }
            foreach idx [array names infiles] {
              $cells1($sum) Item $idx $col1($sum) $infiles($idx)
            }
            [$excel1 ActiveWindow] ScrollColumn [expr 1]

# bold text
            set range [$worksheet1($sum) Range [cellRange 5 1] [cellRange $startrow [expr {$col1($sum)}]]]
            [$range Font] Bold [expr 1]

            [$worksheet1($sum) Columns] AutoFit

#-------------------------------------------------------------------------------
# color entity names in first column, link to documentation
            for {set i 0} {$i < [llength $all_entity]} {incr i} {
              set ent [lindex $all_entity $i]
              set range [$worksheet1($sum) Range "A$entrow($ent)"]
              set cidx [setColorIndex $ent 1]
              if {$cidx > 0} {
                [$range Interior] ColorIndex [expr $cidx]
                if {$lenfilelist > 16} {
                  set range1 [$worksheet1($sum) Range [cellRange $entrow($ent) [expr {[llength $fileList]+$wid+$inc1}]]]
                  [$range1 Interior] ColorIndex [expr $cidx]
                }
              }

# scroll
              set scrollrow 10
              if {$entrow($ent) > $scrollrow} {[$excel1 ActiveWindow] ScrollRow [expr {$entrow($ent)-$scrollrow}]}
            }

            [$excel1 ActiveWindow] ScrollRow [expr 1]

#-------------------------------------------------------------------------------
# links to IFC file, link to individual spreadsheet
            set nf 1
            set idx -1
            foreach file1 $fileList {
              incr nf
              if {$stat([expr {$nf-1}])} {

# link to file
                if {!$opt(HIDELINKS) && [string first "#" $file1] == -1} {
                  set range [$worksheet1($sum) Range [cellRange 4 $nf]]
                  $links Add $range [join $file1] [join ""] [join "Link to IFC file"]
                }

# link to spreadsheet
                set range [$worksheet1($sum) Range [cellRange 3 $nf]]
                incr idx
                regsub -all {\\} [lindex $xnames $idx] "/" xls
                if {!$opt(HIDELINKS) && [string first "#" $file1] == -1} {$links Add $range [join $xls] [join ""] [join "Link to Spreadsheet"]}

# add vertical border when directory changes from column to column
                if {[lsearch $dirchange $nf] != -1} {
                  set nf1 [expr {$nf-1}]
                  set range [$worksheet1($sum) Range [cellRange 3 $nf1] [cellRange $row1($sum) $nf1]]
                  set borders [$range Borders]
                  catch {[$borders Item [expr -4152]] Weight [expr 2]}
                }
              }
            }
            set range [$worksheet1($sum) Range [cellRange 3 [expr {$lenfilelist+1}]] [cellRange $row1($sum) [expr {$lenfilelist+1}]]]
            set borders [$range Borders]
            catch {[$borders Item [expr -4152]] Weight [expr 2]}

# fix column widths
            for {set i 2} {$i <= [expr {$col1($sum)+20}]} {incr i} {
              set val [[$cells1($sum) Item 6 $i] Value]
              if {$val != ""} {
                set range [$worksheet1($sum) Range [cellRange -1 $i]]
                $range ColumnWidth [expr 255]
              }
            }
            [$worksheet1($sum) Columns] AutoFit
            [$worksheet1($sum) Rows] AutoFit

# freeze panes
            [$worksheet1($sum) Range "B[expr {$startrow+1}]"] Select
            [$excel1 ActiveWindow] FreezePanes [expr 1]
            [$worksheet1($sum) Range "A1"] Select
            catch {[$worksheet1($sum) PageSetup] PrintGridlines [expr 1]}

# errors
          } emsg]} {
            errorMsg "ERROR adding information to File Summary spreadsheet: $emsg"
            catch {raise .}
          }
          catch {$excel1 ScreenUpdating 1}
        }
# -------------------------------------------------------------------------------------------------
# time to generate spreadsheets
        set ptime [expr {([clock clicks -milliseconds] - $lasttime)/1000}]
        if {$ptime < 60} {
          set ptime "$ptime seconds"
        } elseif {$ptime < 3600} {
          set ptime "[trimNum [expr {double($ptime)/60.}] 1] minutes"
        } else {
          set ptime "[trimNum [expr {double($ptime)/3600.}] 1] hours"
        }
        if {$opt(XLSCSV) == "Excel"} {
          outputMsg "\n($nfile) Spreadsheets Generated in $ptime" green
        } else {
          outputMsg "\n($nfile) CSV files Generated in $ptime" green
        }
        outputMsg "-------------------------------------------------------------------------------"

# -------------------------------------------------------------------------------------------------
# save spreadsheet
        if {$lenfilelist > 1 && $opt(XLSCSV) == "Excel"} {
          if {[catch {

# set file name for analysis spreadsheet
            set enddir [lindex [split $multiFileDir "/"] end]
            regsub -all " " $enddir "_" enddir
            set aname [file nativename [file join $multiFileDir $enddir\_IFC\_Summary_$lenfilelist.$extXLS]]
            if {[string length $aname] > 218} {
              errorMsg "Spreadsheet file name is too long for Excel ([string length $aname])."
              set aname [file nativename [file join $mydocs $enddir\_IFC\_Summary_$lenfilelist.$extXLS]]
              if {[string length $aname] < 219} {
                errorMsg " Writing Spreadsheet to the home directory."
              }
            }
            catch {file delete -force $aname}

# check if file exists and create new name
            if {[file exists $aname]} {set aname [incrFileName $aname]}

# save spreadsheet
            set aname [checkFileName $aname]
            outputMsg " "
            outputMsg "Saving File Summary Spreadsheet to:"
            outputMsg " [truncFileName $aname 1]" blue
            if {$xlFormat == 51} {
              $workbook1 -namedarg SaveAs Filename $aname FileFormat $xlFormat
            } else {
              $workbook1 -namedarg SaveAs Filename $aname
            }
            set lastXLS1 $aname

# close Excel
            $excel1 Quit
            if {[llength $pidexcel1] == 1} {catch {twapi::end_process $pidexcel1 -force}}

# errors
          } emsg]} {
            errorMsg "ERROR saving File Summary Spreadsheet: $emsg"
            catch {raise .}
          }

# open spreadsheet
          openXLS $aname 0 1

# unset some variables for the multi-file summary
          foreach var {excel1 worksheets1 worksheet1 cells1 row1 col1} {
            if {[info exists $var]} {unset $var}
          }
        }
        update idletasks

# restore saved variables
        saveState
        $buttons(genExcel) configure -state normal
      }

# no files found
    } elseif {[info exists recurse]} {
      set substr ""
      if {$recurse} {set substr " or subdirectories of"}
      errorMsg "No IFC files were found in the directory$substr:\n  [truncFileName [file nativename $multiFileDir]]"
      set choice [tk_messageBox -title "No IFC files found" -type ok -default ok -icon warning \
        -message "No IFC files were found in the directory$substr\n\n[truncFileName [file nativename $multiFileDir]]"]
    }
  }
  update idletasks
}
