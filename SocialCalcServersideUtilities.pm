#
# SocialCalcServersideUtilites.pm
#

#
# Based in part on the SocialCalc 1.1 code, which is based on wikiCalc 1.0.
# SocialCalc 1.1 was:
#  Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
#  All Rights Reserved.
#  Portions (c) Copyright 2007 Socialtext, Inc.
#  All Rights Reserved.
# This code is also based in part on the SocialCalc JavaScript implementation of 2008.
# 
# The new rights in this new code:
#
# (c) Copyright 2008 Socialtext, Inc.
# All Rights Reserved.
#
# This code has NOT been released under an Open Source license at this point.
# It is not to be used without permission of Socialtext, Inc.
#
# Known problems:
#   Doesn't handle UTF-8 currency symbols (and who knows what else)
#   Doesn't handle hidden rows and columns correctly.
#

   use strict;
   use utf8;
   use Time::Local;

   use SocialCalcServersideNumberFormatting;

   require Exporter;
   our @ISA = qw(Exporter);
   our @EXPORT = qw(CreateSheet ParseSheetSave CreateSheetSave DecodeSpreadsheetSave
      CreateCellHTML CreateCellHTMLSave
      CellToString 
      CreateRenderContext RenderSheet
      CalculateCellSkipData PrecomputeSheetFontsAndLayouts CalculateColWidthData 
      RenderTableTag RenderColGroup RenderSizingRow RenderCell);
   our $VERSION = '1.0.0';

   #
   # CONSTANTS
   #

   #
   # Date/time constants
   #

   our $julian_offset = 2415019;
   our $seconds_in_a_day = 24 * 60 * 60;
   our $seconds_in_an_hour = 60 * 60;

   #
   # SHEET DATA FORMAT
   #
   # When parsed, the sheet is stored in a sheet data structure.
   # The structure is a hash with a format as follows:
   #   
   # Cells:
   #
   #   The sheet's cell data is stored as the value of the "cells" key.
   #   The $sheet{cells} value is a hash with cell coordinates as keys
   #   and a hash with the cell attributes as the value. For example,
   #   $sheet{cells}{A1}{formula} might have a value of "SUM(A1:A10)".
   #
   #   Not all cells need an entry nor a data structure. If they are blank and use
   #   default values for their attributes, then they need not be present.
   #
   #   Cell data is stored in a hash with the following "key: value" pairs:
   #
   #      coord: the column/row as a string, e.g., "A1"
   #      datavalue: the value to be used for computation and formatting for display,
   #                 string or numeric (tolerant of numbers stored as strings)
   #      datatype: if present, v=numeric value, t=text value, f=formula,
   #                or c=constant that is not a simple number (like "$1.20")
   #      formula: if present, the formula (without leading "=") for computation or the constant
   #      valuetype: first char is main type, the following are sub-types.
   #                 Main types are b=blank cell, n=numeric, t=text, e=error
   #                 Examples of using sub-types would be "nt" for a numeric time value, "n$" for currency, "nl" for logical
   #  
   #      The following optional values, if present, are mainly used in rendering, overriding defaults:
   #  
   #      bt, br, bb, bl: number of border's definition
   #      layout: layout (vertical alignment, padding) definition number
   #      font: font definition number
   #      color: text color definition number
   #      bgcolor: background color definition number
   #      cellformat: cell format (horizontal alignment) definition number
   #      nontextvalueformat: custom format definition number for non-text values, e.g., numbers
   #      textvalueformat: custom format definition number for text values
   #      colspan, rowspan: number of cells to span for merged cells (only on main cell)
   #      cssc: custom css classname for cell, as text (no special chars)
   #      csss: custom css style definition
   #      mod: modification allowed flag "y" if present
   #      comment: cell comment string
   #
   # Sheet attributes:
   #   Global values, such as default column width, are stored as "attribs".
   #   For example, $sheet{attribs}{defaultcolwidth} = 80.
   #
   #   The sheet attribute key/values are:
   #
   #      lastcol - number
   #      lastrow - number
   #      defaultcolwidth - number or blank (use system default)
   #      defaultrowheight - not used
   #      defaulttextformat:format# - cell format number (alignment) for sheet default for text values
   #      defaultnontextformat:format# - cell format number for sheet default for non-text values (i.e., numbers)
   #      defaultlayout:layout# - default cell layout number in cell layout list
   #      defaultfont:font# - default font number in sheet font list
   #      defaultnontextvalueformat:valueformat# - default non-text (number) value format number in sheet valueformat list
   #      defaulttextvalueformat:valueformat# - default text value format number in sheet valueformat list
   #      defaultcolor:color# - default number for text color in sheet color list
   #      defaultbgcolor:color# - default number for background color in sheet color list
   #      circularreferencecell:coord - cell coord with a circular reference
   #      recalc:value - on/off (on is default). If not "off", appropriate changes to the sheet cause a recalc
   #      needsrecalc:value - yes/no (no is default). If "yes", formula values are not up to date
   #
   # The Column attributes:
   #    Column attributes are stored in $sheet{colattribs}.
   #       $sheet{colattribs}{width}{col-letter(s)}: width or blank for column
   #       $sheet{colattribs}{hide}{col-letter(s)}: yes/no (default is no)
   #
   # The Row attributes:
   #    Row attributes are stored in $sheet{rowattribs}.
   #       $sheet{rowattribs}{hide}{rownumber}: yes/no (default is no)
   #
   # Names are stored as:
   #    $sheet{names}{"name"} = {desc => "description", definition => "definition"}
   #
   # The lookup value lists are stored as an array of strings and a hash to do reverse lookups.
   #    The lookup value lists are: layout, font, color, borderstyle, cellformat, and valueformat.
   #    $sheet{list-name . "s"}[$num] = value
   #    $sheet{list-name . "hash"}{"value"} = $num
   #
   # The range this was copied from (for partial saves) is in $sheet{copiedfrom}.
   #

   #
   # The following data structures are used as lookups by the save and parse code to make it 
   # more general and easier to extend:
   #

   our %sheetAttribsLongToShort = (lastcol => "c", lastrow => "r", defaultcolwidth => "w", defaultrowheight => "h",
      defaulttextformat => "tf", defaultnontextformat => "ntf", defaulttextvalueformat => "tvf", defaultnontextvalueformat => "ntvf",
      defaultlayout => "layout", defaultfont => "font", defaultcolor => "color", defaultbgcolor => "bgcolor",
      circularreferencecell => "circularreferencecell", recalc => "recalc", needsrecalc => "needsrecalc");

   our %sheetAttribsShortToLong = ("c" => "lastcol", "r" => "lastrow", "w" => "defaultcolwidth",
      "h" => "defaultrowheight", "tf" => "defaulttextformat", "ntf" => "defaultnontextformat",
      "tvf" => "defaulttextvalueformat", "ntvf" => "defaultnontextvalueformat",
      "layout" => "defaultlayout", "font" => "defaultfont", "color" => "defaultcolor",
      "bgcolor" => "defaultbgcolor", "circularreferencecell" => "circularreferencecell",
      "recalc" => "recalc", "needsrecalc" => "needsrecalc");

   our %sheetAttribsStyle = (lastcol => 1, lastrow => 1, defaultcolwidth => 2, defaultrowheight => 1,
      defaulttextformat => 1, defaultnontextformat => 1, defaulttextvalueformat => 1, defaultnontextvalueformat => 1,
      defaultlayout => 1, defaultfont => 1, defaultcolor => 1, defaultbgcolor => 1,
      circularreferencecell => 2, recalc => 2, needsrecalc => 2);

   our %cellAttribsStyle = (v => "v", t => "t", vt => "vt", vtf => "vtf", vtc => "vtc",
      e => "decode", l => "plain", f => "plain", c => "plain", bg => "plain", cf => "plain", cvf => "plain",
      ntvf => "plain", tvf => "plain", colspan => "plain", rowspan => "plain", cssc => "plain",
      csss => "decode", mod => "plain", comment => "decode", b => "b");

   our %cellAttribTypeLong = (e => "errors", l => "layout", f => "font", c => "color", bg => "bgcolor", cf => "cellformat",
      ntvf => "nontextvalueformat", tvf => "textvalueformat", colspan => "colspan", rowspan => "rowspan", cssc => "cssc",
      csss => "csss", mod => "mod", comment => "comment");

   our %vlistNames = (layout => "layout", font => "font", color => "color", border => "borderstyle",
      cellformat => "cellformat", valueformat => "valueformat");

#
# $newsheet = CreateSheet();
#
# Returns a new sheet structure with all of the parts initialized.
#

sub CreateSheet {

   my $sheet = {
      version => "", # start with not set
      cells => {},
      attribs => {lastcol => 1, lastrow => 1, defaultlayout => 0},
      rowattribs => {hide => {}, height => {}},
      colattribs => {hide => {}, width => {}},
      names => {},
      layouts => [],
      layouthash => {},
      fonts => [],
      fonthash => {},
      colors => [],
      colorhash => {},
      borderstyles => [],
      borderstylehash => {},
      cellformats => [],
      cellformathash => {},
      valueformats => [],
      valueformathash => {}
      };

   return $sheet;

   }

#
# $errorstring = ParseSheetSave($sheet, $str);
#
# Parses the string $str in SocialCalc Sheet Save format and sets the
# values in $sheet.
#

sub ParseSheetSave {

   my ($sheet, $str) = @_;

   my $maxcol = 1;
   my $maxrow = 1;
   $sheet->{attribs}{lastcol} = 0;
   $sheet->{attribs}{lastrow} = 0;

   my @lines = split(/\r\n|\n/, $str);

   my ($linetype, $value, $coord, $type, $rest, $cell, $cr, $style, $valuetype, 
       $formula, $attrib, $num, $name, $desc);

   foreach my $line (@lines) {
      ($linetype, $rest) = split(/:/, $line, 2);

      if ($linetype eq "version") {
         $sheet->{version} = $rest;
         }

      elsif ($linetype eq "cell") {
         ($coord, $type, $rest) = split(/:/, $rest, 3);
         $coord = uc($coord);
         $cell = {'coord' => $coord} if $type; # start with minimal cell
         $sheet->{cells}{$coord} = $cell;
         $cr = CoordToCR($coord);
         $maxcol = $cr->{col} if $cr->{col} > $maxcol;
         $maxrow = $cr->{row} if $cr->{row} > $maxrow;
         while ($type) {
            $style = $cellAttribsStyle{$type};
            last if !$style; # process known, non-null types
            if ($style eq "v") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $cell->{datavalue} = DecodeFromSave($value);
               $cell->{datatype} = "v";
               $cell->{valuetype} = "n";
               }
            elsif ($style eq "t") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $cell->{datavalue} = DecodeFromSave($value);
               $cell->{datatype} = "t";
               $cell->{valuetype} = "t"; # !! should be Constants.textdatadefaulttype
               }
            elsif ($style eq "vt") {
               ($valuetype, $value, $type, $rest) = split(/:/, $rest, 4);
               $cell->{datavalue} = DecodeFromSave($value);
               if (substr($valuetype,0,1) eq "n") {
                  $cell->{datatype} = "n";
                  }
               else {
                  $cell->{datatype} = "t";
                  }
               $cell->{valuetype} = $valuetype;
               }
            elsif ($style eq "vtf") {
               ($valuetype, $value, $formula, $type, $rest) = split(/:/, $rest, 5);
               $cell->{datavalue} = DecodeFromSave($value);
               $cell->{formula} = DecodeFromSave($formula);
               $cell->{datatype} = "f";
               $cell->{valuetype} = $valuetype;
               }
            elsif ($style eq "vtc") {
               ($valuetype, $value, $formula, $type, $rest) = split(/:/, $rest, 5);
               $cell->{datavalue} = DecodeFromSave($value);
               $cell->{formula} = DecodeFromSave($formula);
               $cell->{datatype} = "c";
               $cell->{valuetype} = $valuetype
               }
            elsif ($style eq "b") {
               my ($t, $r, $b, $l);
               ($t, $r, $b, $l, $type, $rest) = split(/:/, $rest, 6);
               $cell->{bt} = $t;
               $cell->{br} = $r;
               $cell->{bb} = $b;
               $cell->{bl} = $l;
               }
            elsif ($style eq "plain") {
               $attrib = $cellAttribTypeLong{$type};
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $cell->{$attrib} = $value;
               }
            elsif ($style eq "decode") {
               $attrib = $cellAttribTypeLong{$type};
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $cell->{$attrib} = DecodeFromSave($value);
               }
            else {
               last;
               }
            }
         }

      elsif ($linetype eq "col") {
         ($coord, $type, $rest) = split(/:/, $rest, 3);
         $coord = uc($coord); # normalize to upper case
         while ($type) {
            if ($type eq "w") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{colattribs}{width}{$coord} = $value;
               }
            elsif ($type eq "hide") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{colattribs}{hide}{$coord} = $value;
               }
            else {
               last;
               }
            }
         }

      elsif ($linetype eq "row") {
         ($coord, $type, $rest) = split(/:/, $rest, 3);
         while ($type) {
            if ($type eq "h") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{rowattribs}{height}{$coord} = $value;
               }
            elsif ($type eq "hide") {
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{rowattribs}{hide}{$coord} = $value;
               }
            else {
               last;
               }
            }
         }

      elsif ($linetype eq "sheet") {
         ($type, $rest) = split(/:/, $rest, 2);
         while ($type) {
            $attrib = $sheetAttribsShortToLong{$type};
            $style = $sheetAttribsStyle{$attrib};
            last if !$style; # process known, non-null types
            if ($style==1) { # plain number
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{attribs}{$attrib} = $value-0;
               }
            elsif ($style==2) { # text
               ($value, $type, $rest) = split(/:/, $rest, 3);
               $sheet->{attribs}{$attrib} = $value;
               }
            else {
               last;
               }
            }
         }

      elsif ($linetype eq "name") {
         ($name, $desc, $value) = split(/:/, $rest, 3);
         $name = uc (DecodeFromSave($name));
         $sheet->{names}{$name}{desc} = DecodeFromSave($desc);
         $sheet->{names}{$name}{definition} = DecodeFromSave($value);
         }

      elsif ($vlistNames{$linetype}) { # if one of the value lists, process
         $style = $vlistNames{$linetype}; # get base name
         ($num, $value) = split(/:/, $rest, 2);
         $value = DecodeFromSave($value);
         $sheet->{$style . "s"}->[$num] = $value;
         $sheet->{$style . "hash"}{$value} = $num;
         }

      elsif ($linetype eq "") { # skip blank lines
         }

      }

   $sheet->{attribs}{lastcol} ||= $maxcol || 1;
   $sheet->{attribs}{lastrow} ||= $maxrow || 1;

   }


#
# $outstr = CreateSheetSave($sheet, $range);
#
# Returns the sheet as a string in SocialCalc Sheet Save format.
# If $range is present, it is in the form similar to "A1:B7",
# and just those cells will be used.
#

sub CreateSheetSave {

   my ($sheet, $range) = @_;
   my $outstr;

   my $prange;

   if ($range) {
      $prange = ParseRange($range);
      }
   else {
      $prange = ();
      $prange->{cr1}{row} = 1;
      $prange->{cr1}{col} = 1;
      $prange->{cr2}{row} = $sheet->{attribs}{lastrow};
      $prange->{cr2}{col} = $sheet->{attribs}{lastcol};
      }

   my $cr1 = $prange->{cr1};
   my $cr2 = $prange->{cr2};

   for (my $row=$cr1->{row}; $row <= $cr2->{row}; $row++) {
      for (my $col=$cr1->{col}; $col <= $cr2->{col}; $col++) {
         my $coord = CRToCoord($col, $row);
         my $cell = $sheet->{cells}{$coord};
         next if !$cell;
         my $line = CellToString($cell);
         next if !$line; # ignore completely empty cells
         $line = "cell:$coord$line\n";
         $outstr .= $line;
         }
      }

   for (my $col=1; $col <= $sheet->{attribs}{lastcol}; $col++) {
      my $coord = NumberToCol($col);
      if ($sheet->{colattribs}{width}{$coord}) {
         $outstr .= "col:$coord:w:$sheet->{colattribs}{width}{$coord}\n";
         }
      if ($sheet->{colattribs}{hide}{$coord}) {
         $outstr .= "col:$coord:hide:$sheet->{colattribs}{hide}{$coord}\n";
         }
      }

   for (my $row=1; $row <= $sheet->{attribs}{lastrow}; $row++) {
      if ($sheet->{rowattribs}{height}{$row}) {
         $outstr .= "row:$row:h:$sheet->{rowattribs}{height}{$row}\n";
         }
      if ($sheet->{rowattribs}{hide}{$row}) {
         $outstr .= "row:$row:hide:$sheet->{rowattribs}{hide}{$row}\n";
         }
      }

   my $line = "sheet";
   for my $attrib (keys %sheetAttribsLongToShort) {
      my $value = EncodeForSave($sheet->{attribs}{$attrib});
      if ($value) {
         $line .= ":$sheetAttribsLongToShort{$attrib}:$value";
         }
      }
   $outstr .= "$line\n";

   foreach my $name (keys %{$sheet->{names}}) {
      $outstr .= "name:" . uc(EncodeForSave($name)) . ":" .
                   EncodeForSave($sheet->{names}{$name}{desc}) . ":" .
                   EncodeForSave($sheet->{names}{$name}{definition}) . "\n";
      }

   foreach my $type (keys %vlistNames) { # check each value list
      my $vlist = $sheet->{$vlistNames{$type} . "s"};
      for (my $i=1; $i < @$vlist; $i++) { # output all values present
         my $value = EncodeForSave($vlist->[$i]);
         $outstr .= "$type:$i:$value\n";
         }
      }

   if ($range) {
      $outstr .= "copiedfrom:$range\n";
      }

   return $outstr;

   }


#
# $str = CellToString($cell)
#
# Returns string in cell save format.
#

sub CellToString {

   my $cell = shift @_;

   my $str = "";

   return $str if (!$cell);

   my $value = EncodeForSave($cell->{datavalue});
   if ($cell->{datatype} eq "v") {
      if ($cell->{valuetype} eq "n") {
         $str .= ":v:$value";
         }
      else {
         $str .= ":vt:$cell->{valuetype}:$value";
         }
      }
   elsif ($cell->{datatype} eq "t") {
      if ($cell->{valuetype} eq "t") { # !!! should be Constants.textdatadefaulttype
         $str .= ":t:$value";
         }
      else {
         $str .= ":vt:$cell->{valuetype}:$value";
         }
      }
   else {
      my $formula = EncodeForSave($cell->{formula});
      if ($cell->{datatype} eq "f") {
         $str .= ":vtf:$cell->{valuetype}:$value:$formula";
         }
      elsif ($cell->{datatype} eq "c") {
         $str .= ":vtc:$cell->{valuetype}:$value:$formula";
         }
      }
   if ($cell->{errors}) {
      $str .= ":e:" . EncodeForSave($cell->{errors});
      }
   my $t = $cell->{bt} || "";
   my $r = $cell->{br} || "";
   my $b = $cell->{bb} || "";
   my $l = $cell->{bl} || "";
   if ($t || $r || $b || $l) {
      $str .= ":b:$t:$r:$b:$l";
      }
   if ($cell->{layout}) {
      $str .= ":l:$cell->{layout}";
      }
   if ($cell->{font}) {
      $str .= ":f:$cell->{font}";
      }
   if ($cell->{color}) {
      $str .= ":c:$cell->{color}";
      }
   if ($cell->{bgcolor}) {
      $str .= ":bg:$cell->{bgcolor}";
      }
   if ($cell->{cellformat}) {
      $str .= ":cf:$cell->{cellformat}";
      }
   if ($cell->{textvalueformat}) {
      $str .= ":tvf:$cell->{textvalueformat}";
      }
   if ($cell->{nontextvalueformat}) {
      $str .= ":ntvf:$cell->{nontextvalueformat}";
      }
   if ($cell->{colspan}) {
      $str .= ":colspan:$cell->{colspan}";
      }
   if ($cell->{rowspan}) {
      $str .= ":rowspan:$cell->{rowspan}";
      }
   if ($cell->{cssc}) {
      $str .= ":cssc:$cell->{cssc}";
      }
   if ($cell->{csss}) {
      $str .= ":csss:" . EncodeForSave($cell->{csss});
      }
   if ($cell->{mod}) {
      $str .= ":mod:$cell->{mod}";
      }
   if ($cell->{comment}) {
      $str .= ":comment:" . EncodeForSave($cell->{comment});
      }

   return $str;

   }

# *************************************
#
# Rendering Code
#
# *************************************

#
# $context = CreateRenderContext($sheet)
#

sub CreateRenderContext {

   my $sheet = shift @_;

   my $context = {

      sheet => $sheet,
      hideRowsCols => 0, # pay attention to row/col hide settings (currently ignored)
      cellIDprefix => "", # if non-null, each cell will render with an ID starting with this
      defaultcolwidth => 80,
      defaultlayout => "padding:2px 2px 1px 2px;vertical-align:top;",

      globaldefaultfontstyle => "normal normal",
      globaldefaultfontsize => "small",
      globaldefaultfontfamily => "Verdana,Arial,Helvetica,sans-serif",

      explicitStyles => {
         skippedcell => "font-size:small;background-color:#CCC",
         comment => "background-repeat:no-repeat;background-position:top right;background-image:url(images/sc-commentbg.gif);"
         },

      classnames => {
         skippedcell => "",
         comment => ""
         },

      # initialize calculated values to be filled in later:

      cellskip => {}, # this-cell => coord of cell covering this cell (only for covered cells)
      colwidth => [], # column widths, taking into account defaults
      totalwidth => 0, # total table width
      maxcol => 0, # max col to display, adding long spans, etc.
      maxrow => 0, # max row to display, adding long spans, etc.
      defaultfontstyle => "",
      defaultfontsize => "",
      defaultfontfamily => "",
      fonts => [], # for each fontnum, {style: fs, weight: fw, size: fs, family: ff}
      layouts => [], # for each layout, "padding:Tpx Rpx Bpx Lpx;vertical-align:va;"

      };

   }

#
# $outstr = RenderSheet($context, $options)
#
# Returns HTML for table rendering the sheet in that context.
#
# The options (passed to cell rendering code) are:
#
#    newwinlinks => t/f (default is true) - whether text-link has target="_blank" by default or not
#

sub RenderSheet {

   my ($context, $options) = @_;
   my $sheet = $context->{sheet};

   CalculateCellSkipData($context, $options);
   PrecomputeSheetFontsAndLayouts($context, $options);
   CalculateColWidthData($context, $options);

   my $outstr = RenderTableTag($context, $options); # start with table tag

   $outstr .= RenderColGroup($context, $options); # then colgroup section

   $outstr .= "<tbody>";

   $outstr .= RenderSizingRow($context, $options); # add tiny row so all cols have something despite spans

   for (my $row=1; $row <= $context->{maxrow}; $row++) {
      $outstr .= "<tr>";
      for (my $col=1; $col <= $context->{maxcol}; $col++) {
         $outstr .= RenderCell($context, $row, $col, $options);
         }
      $outstr .= "</tr>";
      }

   $outstr .= "</tbody>";
   $outstr .= "</table>";

   return $outstr;

   }

#
# CalculateCellSkipData($context, $options)
#
# Figures out which cells are to be skipped because of row and column spans.
#

sub CalculateCellSkipData {

   my ($context, $options) = @_;

   my $sheet = $context->{sheet};
   my $sheetattribs = $sheet->{attribs};

   $context->{maxcol} = $sheetattribs->{lastcol};
   $context->{maxrow} = $sheetattribs->{lastrow};
   $context->{cellskip} = {}; # reset

   # Calculate cellskip data

   for (my $row=1; $row <= $sheetattribs->{lastrow}; $row++) {
      for (my $col=1; $col <= $sheetattribs->{lastcol}; $col++) { # look for spans and set cellskip for skipped cells
         my $coord = CRToCoord($col, $row);
         my $cell = $sheet->{cells}{$coord};
         # don't look at undefined cells (they have no spans) or skipped cells
         if (!$cell || $context->{cellskip}{$coord}) {
            next;
            }
         my $colspan = $cell->{colspan} || 1;
         my $rowspan = $cell->{rowspan} || 1;
         if ($colspan>1 || $rowspan>1) {
            for (my $skiprow=$row; $skiprow<$row+$rowspan; $skiprow++) {
               for (my $skipcol=$col; $skipcol<$col+$colspan; $skipcol++) { # do the setting on individual cells
                  my $skipcoord = CRToCoord($skipcol, $skiprow);
                  if ($skipcoord ne $coord) { # flag other cells to point back here
                     $context->{cellskip}{$skipcoord} = $coord;
                     }
                  if ($skiprow>$context->{maxrow}) {
                     $context->{maxrow} = $skiprow;
                     }
                  if ($skipcol>$context->{maxcol}) {
                     $context->{maxcol} = $skipcol;
                     }
                  }
               }
            }
         }
      }

   return;

   }


#
# PrecomputeSheetFontsAndLayouts($context, $options)
#
# Fills out the fonts and layouts arrays in the context.
#

sub PrecomputeSheetFontsAndLayouts {

   my ($context, $options) = @_;
   my $sheet = $context->{sheet};
   my $attribs = $sheet->{attribs};

   if ($attribs->{defaultfont}) {
      my $defaultfont = $sheet->{fonts}->[$attribs->{defaultfont}];
      $defaultfont =~ s/^\*/$context->{globaldefaultfontstyle}/e;
      $defaultfont =~ s/(.+)\*(.+)/$1.$context->{globaldefaultfontsize}.$2/e;
      $defaultfont =~ s/\*$/$context->{globaldefaultfontfamily}/e;
      $defaultfont =~ m/^(\S+? \S+?) (\S+?) (\S.*)$/;
      $context->{defaultfontstyle} = $1;
      $context->{defaultfontsize} = $2;
      $context->{defaultfontfamily} = $3;
      }
   else {
      $context->{defaultfontstyle} = $context->{globaldefaultfontstyle};
      $context->{defaultfontsize} = $context->{globaldefaultfontsize};
      $context->{defaultfontfamily} = $context->{globaldefaultfontfamily};
      }

   for (my $num=1; $num<@{$sheet->{fonts}}; $num++) { # precompute fonts by filling in the *'s
      my $s = $sheet->{fonts}->[$num];
      $s =~ s/^\*/$context->{defaultfontstyle}/e;
      $s =~ s/(.+)\*(.+)/$1.$context->{defaultfontsize}.$2/e;
      $s =~ s/\*$/$context->{defaultfontfamily}/e;
      $s =~ m/^(\S+?) (\S+?) (\S+?) (\S.*)$/;
      $context->{fonts}->[$num] = {style => $1, weight => $2, size => $3, family => $4};
      }

   my $layoutre = qr/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/;
   my @dparts = ($context->{defaultlayout} =~ m/$layoutre/);
   my @sparts;
   if ($attribs->{defaultlayout}) {
      @sparts = ($sheet->{layouts}->[$attribs->{defaultlayout}] =~ m/$layoutre/); # get sheet defaults, if set
      }
   else {
      @sparts = ("*", "*", "*", "*", "*");
      }

   for (my $num=1; $num<@{$sheet->{layouts}}; $num++) { # precompute layouts by filling in the *'s
      my $s = $sheet->{layouts}->[$num];
      my @parts = ($s =~ m/$layoutre/);
      for (my $i=0; $i<=4; $i++) {
         if ($parts[$i] eq "*") {
            $parts[$i] = ($sparts[$i] ne "*" ? $sparts[$i] : $dparts[$i]); # if *, sheet default or built-in
            }
         }
      $context->{layouts}->[$num] = "padding:$parts[0] $parts[1] $parts[2] $parts[3];vertical-align:$parts[4];";
      }

   return;

   }


#
# CalculateColWidthData($context, $options)
#
# Saves the width of all the columns, taking the default into account,
# and calculates the total width of the table.
#
# Do CalculateCellSkipData first to calculate the max col and row.
#

sub CalculateColWidthData {

   my ($context, $options) = @_;
   my $sheet = $context->{sheet};

   my $totalwidth = 0;

   for (my $col=1; $col <= $context->{maxcol}; $col++) {
      my $colwidth = $sheet->{colattribs}{width}{NumberToCol($col)} ||
         $sheet->{attribs}->{defaultcolwidth} || $context->{defaultcolwidth};
      if ($colwidth eq "blank" || $colwidth eq "auto" || $colwidth eq "") {
         $colwidth = "";
         }
      $context->{colwidth}->[$col] = "$colwidth";
      $totalwidth += ($colwidth && $colwidth>0) ? $colwidth : 10;
      }

   $context->{totalwidth} = $totalwidth;

   return;

   }


#
# $outstr = RenderTableTag($context, $options)
#
# Returns "<table ...>"
#

sub RenderTableTag {

   my ($context, $options) = @_;

   return qq!<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; width: $context->{totalwidth}px">!;

   }

#
# $outstr = RenderColGroup($context, $options)
#
# Returns "<colgroup><col width=...>...</colgroup>"
#

sub RenderColGroup {

   my ($context, $options) = @_;
   my $colwidths = $context->{colwidth};

   my $outstr = "<colgroup>";

   for (my $col=1; $col <= $context->{maxcol}; $col++) {
      if ($colwidths->[$col]) {
         $outstr .= qq!<col width="$colwidths->[$col]"/>!;
         }
      else {
         $outstr .= "<col/>";
         }
      }
   $outstr .= "</colgroup>";

   return $outstr;

   }

#
# $outstr = RenderSizingRow($context, $options)
#
# Returns row with blank cells for all columns just in case spans leave columns
# with no cells -- this has bugs in some cases with some browsers.
#

sub RenderSizingRow {

   my ($context, $options) = @_;
   my $colwidths = $context->{colwidth};

   my $outstr = "<tr>";

   for (my $col=1; $col <= $context->{maxcol}; $col++) {
      if ($colwidths->[$col]) {
         $outstr .= qq!<td height="1" width="$colwidths->[$col]"></td>!;
         }
      else {
         $outstr .= qq!<td height="1"></td>!;
         }
      }
   $outstr .= "</tr>";

   return $outstr;

   }


#
# $outstr = RenderCell($context, $row, $col, $options)
#
# Returns "<td ...>...</td>"
# If skipped cell (because of span) returns "".
#
# If $options->{parts}, returns:
#    {tag => "span=...", style => "padding:...", classes => "commentclass",
#     value => "displayvalue", all => "<td.."}
#

sub RenderCell {

   my ($context, $row, $col, $options) = @_;
   my $sheet = $context->{sheet};
   my $sheetattribs = $sheet->{attribs};

   my $coord = CRToCoord($col, $row);

   my $outstr = "";
   my $tagstr = "";
   my $stylestr = "";
   my $classstr = "";
   my $displayvalue = "";

   if ($context->{cellskip}{$coord}) { # skip if within a span
      return $outstr;
      }

   if ($context->{cellIDprefix}) {
      $tagstr .= " " if $tagstr;
      $tagstr .= qq!id="$context->{cellIDprefix}$coord"!;
      }

   my $cell = $sheet->{cells}{$coord};
   if (!$cell) {
      $cell = {datatype => "b"};
      }

   if ($cell->{colspan} > 1) {
      my $span = 1;
      for (my $num=1; $num<$cell->{colspan}; $num++) {
          if ($sheet->{colattribs}{hide}{NumberToCol($col+$num)} ne "yes") {
             $span++;
             }
          }
      $tagstr .= " " if $tagstr;
      $tagstr .= qq!colspan="$span"!;
      }

   if ($cell->{rowspan} > 1) {
      my $span = 1;
      for (my $num=1; $num<$cell->{rowspan}; $num++) {
          if ($sheet->{colattribs}{hide}{$row+$num} ne "yes") {
             $span++;
             }
          }
      $tagstr .= " " if $tagstr;
      $tagstr .= qq!rowspan="$span"!;
      }

   my $num = $cell->{layout} || $sheetattribs->{defaultlayout};
   if ($num) {
      $stylestr .= $context->{layouts}[$num]; # use precomputed layout with "*"'s filled in
      }
   else {
      $stylestr .= $context->{defaultlayout};
      }

   $num = $cell->{font} || $sheetattribs->{defaultfont};
   if ($num) { # get expanded font strings in context
      my $t = $context->{fonts}->[$num]; # do each - plain "font:" style sets all sorts of other values, too (Safari font-stretch problem on cssText)
      $stylestr .= "font-style:$t->{style};font-weight:$t->{weight};font-size:$t->{size};font-family:$t->{family};";
      }
   else {
      if ($context->{defaultfontsize}) {
         $stylestr .= "font-size:$context->{defaultfontsize};";
         }
      if ($context->{defaultfontfamily}) {
         $stylestr .= "font-family:$context->{defaultfontfamily};";
         }
      }

   $num = $cell->{color} || $sheetattribs->{defaultcolor};
   if ($num) {
      $stylestr .= "color:$sheet->{colors}->[$num];";
      }

   $num = $cell->{bgcolor} || $sheetattribs->{defaultbgcolor};
   if ($num) {
      $stylestr .= "background-color:$sheet->{colors}->[$num];";
      }

   $num = $cell->{cellformat};
   if ($num) {
      $stylestr .= "text-align:$sheet->{cellformats}->[$num];";
      }
   else {
      my $t = substr($cell->{valuetype}, 0, 1);
      if ($t eq "t") {
         $num = $sheetattribs->{defaulttextformat};
         if ($num) {
            $stylestr .= "text-align:$sheet->{cellformats}->[$num];";
            }
         }
      elsif ($t eq "n") {
         $num = $sheetattribs->{defaultnontextformat};
         if ($num) {
            $stylestr .= "text-align:$sheet->{cellformats}->[$num];";
            }
         else {
            $stylestr .= "text-align:right;";
            }
         }
      else {
         $stylestr .= "text-align:left;";
         }
      }

   if ($cell->{bt} &&
      ($cell->{bt}==$cell->{br} && $cell->{bt}==$cell->{bb} && $cell->{bt}==$cell->{bl})) {
      $stylestr .= "border:$sheet->{borderstyles}->[$cell->{bt}];";
      }

   else {
      $num = $cell->{bt};
      if ($num) {
         $stylestr .= "border-top:$sheet->{borderstyles}->[$num];";
         }

      $num = $cell->{br};
      if ($num) {
         $stylestr .= "border-right:$sheet->{borderstyles}->[$num];";
         }

      $num = $cell->{bb};
      if ($num) {
         $stylestr .= "border-bottom:$sheet->{borderstyles}->[$num];";
         }

      $num = $cell->{bl};
      if ($num) {
         $stylestr .= "border-left:$sheet->{borderstyles}->[$num];";
         }
      }

   if ($cell->{comment}) {
      if ($context->{classnames}{comment}) {
         $classstr .= " " if $classstr;
         $classstr .= $context->{classnames}{comment};
         }
      $stylestr .= $context->{explicitStyles}{comment};
#!! put comment text in title tag...?
      }

   $displayvalue = FormatValueForDisplay($sheet, $cell->{datavalue}, $coord, $options);

   # Assemble output

   $outstr .= "<td";

   if ($tagstr) {
      $outstr .= qq! $tagstr"!;
      }

   if ($classstr) {
      $outstr .= qq! class="$classstr"!;
      }

   if ($stylestr) {
      $outstr .= qq! style="$stylestr"!;
      }

   $outstr .= ">";
   $outstr .= "$displayvalue";
   $outstr .= "</td>";

   if ($options->{parts}) {
      return {tag => $tagstr, style => $stylestr, classes => $classstr,
         value => $displayvalue, all => $outstr};
      }
   else {
      return $outstr;
      }

   }


#
# $parts = DecodeSpreadsheetSave($str)
#
# Separates the parts from a spreadsheet save string, returning reference to
# a hash with the sub-strings:
#
#    {type1 => $string, type2 => ...}
#

sub DecodeSpreadsheetSave {

   my $str = shift @_;

   my $parts = {};
   my @partlist;

   # Parse multipart-MIME setup

   my $OK = $str =~ m/^MIME-Version:\s1\.0/mig;
   return $parts unless $OK; # no parts
   $OK = $str =~ /(.*?)\nContent-Type:\s*multipart\/mixed;\s*boundary=(\S+)/mig;
   return $parts unless $OK;
   my $boundary = $2;

   # First part is header which says what comes next

   $OK = $str =~ m/^--$boundary$/mg; # find header top boundary
   return $parts unless $OK;
   $OK = $str =~ m/^$/mg; # skip to after blank line
   return $parts unless $OK;
   $OK = $str =~ m/(.*?)\n--$boundary$/msig; # get header part
   return $parts unless $OK;
   my $header = $1;

   my @lines = split(/\r\n|\n/, $header); # parse header
   for (my $i=0; $i<@lines; $i++) {
      my @line = split(":", $lines[$i]);
      if ($line[0] eq "version") {
         }
      elsif ($line[0] eq "part") {
         push @partlist, $line[1];
         }
      }

   # Get each part

   for (my $pnum=0; $pnum<@partlist; $pnum++) {
      $OK = $str =~ m/^(\r\n|\n)/mg; # skip to after blank line ending mime-part header
      return $parts unless $OK;

      $boundary .= "--" if ($pnum == @partlist-1); # last part has different boundary
      $OK = $str =~ m/(.*?)\n--$boundary$/msig; # get art's string
      return $parts unless $OK;
      $parts->{$partlist[$pnum]} = $1; # save string to return
      }

   return $parts;

   }



#
# $result = CreateCellHTML($sheet, $coord, $options)
#
# Returns the HTML representation of a cell. Blank is "", not "&nbsp;".
#

sub CreateCellHTML {

   my ($sheet, $coord, $options) = @_;

   my $result = "";
   my $cell = $sheet->{cells}{$coord};

   if (!$cell) {
      return "";
      }

   $result = FormatValueForDisplay($sheet, $cell->{datavalue}, $coord, $options);

   if ($result eq "&nbsp;") {
      $result = "";
      }

   return $result;

   }

#
# $result = CreateCellHTMLSave($sheet, $range, $options)
#
# Returns the HTML representation of a range of cells, or the whole sheet if range is null.
# The form is:
#    version:1.0
#    coord:cell-HTML
#    coord:cell-HTML
#    ...
#
# Empty cells are skipped. The cell-HTML is encoded with ":"=>"\c", newline=>"\n", and "\"=>"\b".
#

sub CreateCellHTMLSave {

   my ($sheet, $range, $options) = @_;

   my $prange;

   if ($range) {
      $prange = ParseRange($range);
      }
   else {
      $prange = {cr1 => {row => 1, col => 1},
                cr2 => {row => $sheet->{attribs}{lastrow}, col => $sheet->{attribs}{lastcol}}};
      }
   my $cr1 = $prange->{cr1};
   my $cr2 = $prange->{cr2};

   my $result = "version:1.0\n";

   for (my $row=$cr1->{row}; $row <= $cr2->{row}; $row++) {
      for (my $col=$cr1->{col}; $col <= $cr2->{col}; $col++) {
         my $coord = CRToCoord($col, $row);
         my $cell = $sheet->{cells}{$coord};
         if (!$cell) {
            next;
            }
         my $cellHTML = FormatValueForDisplay($sheet, $cell->{datavalue}, $coord, $options);
         if ($cellHTML eq "&nbsp;") {
            next;
            }
         my $ecellHTML = EncodeForSave($cellHTML);
         $result .= "$coord:$ecellHTML\n";
         }
      }

   return $result;

   }


#
# $result = CreateCSV($sheet, $range, $options)
#
# Returns the CSV representation of a range of cells, or of the whole sheet if range is null.
# It returns the raw values, not formatted as dates, etc.
#

sub CreateCSV {

   my ($sheet, $range, $options) = @_;

   my ($prange, $str);

   if ($range) {
      $prange = ParseRange($range);
      }
   else {
      $prange = {cr1 => {row => 1, col => 1},
                cr2 => {row => $sheet->{attribs}{lastrow}, col => $sheet->{attribs}{lastcol}}};
      }
   my $cr1 = $prange->{cr1};
   my $cr2 = $prange->{cr2};

   my $result = "";

   for (my $row=$cr1->{row}; $row <= $cr2->{row}; $row++) {
      for (my $col=$cr1->{col}; $col <= $cr2->{col}; $col++) {
         my $coord = CRToCoord($col, $row);
         my $cell = $sheet->{cells}{$coord};

         if ($cell->{errors}) {
            $str = $cell->{errors};
            }
         else {
            $str = $cell->{datavalue} || "";
            }

         $str =~ s/"/""/g; # double quotes
         if ($str =~ m/[, \n"]/) {
            $str = qq!"$str"!; # add quotes
            }
         if ($col > $cr1->{col}) {
            $str = ",$str"; # add commas
            }
         $result .= $str;
         }
      $result .= "\n";
      }

   return $result;

   }


# * * * * * * * * * * * * * * * * * * * *
#
# Common helping routines
#
# * * * * * * * * * * * * * * * * * * * *

#
# $rangehash = ParseRange($range)
#
# Takes $range with "A1:B7" and returns hash with:
#    cr1 => {row => rownumber, col => colnumber, coord => CR},
#    cr2 => {row => rownumber, col => colnumber, coord => CR}
#

sub ParseRange {

   my $range = shift @_;

   if (!$range) {
      $range = "A1:A1"; # error return, hopefully benign
      }
   $range = uc $range;

   my ($cr1, $cr2);
   $range =~ m/^(.+):(.+)$/;
   if ($2) { # has two parts
      $cr1 = CoordToCR($1);
      $cr2 = CoordToCR($2);
      }
   else {
      $cr1 = CoordToCR($range);
      $cr2 = CoordToCR($range);
      }
   return {cr1 => $cr1, cr2 => $cr2};

   }


#
# (col => $col, row => $row) = CoordToCR($coord)
#
# Turns B3 into (col => 2, row => 3, coord => "B3"). The default for both is 1.
# If range, only do this to first coord
#

sub CoordToCR {

   my $coord = shift @_;

   $coord = lc($coord);
   $coord =~ s/\$//g;
   $coord =~ m/([a-z])([a-z])?(\d+)/;
   my $col = ord($1) - ord('a') + 1 ;
   $col = 26 * $col + ord($2) - ord('a') + 1 if $2;

   return {col => $col, row => $3, coord => uc $coord};

}


#
# $coord = CRToCoord($col, $row)
#
# Turns (2, 3) into B3. The default for both is 1.
#

sub CRToCoord {

   my ($col, $row) = @_;

   $row = 1 unless $row > 1;
   $col = 1 unless $col > 1;

   my $col_high = int(($col - 1) / 26);
   my $col_low = ($col - 1) % 26;

   my $coord = chr(ord('A') + $col_low);
   $coord = chr(ord('A') + $col_high - 1) . $coord if $col_high;
   $coord .= $row;

   return $coord;

}


#
# $col = ColToNumber($colname)
#
# Turns B into 2. The default is 1.
#

sub ColToNumber {

   my $coord = shift @_;

   $coord = lc($coord);
   $coord =~ m/([a-z])([a-z])?/;
   return 1 unless $1;
   my $col = ord($1) - ord('a') + 1 ;
   $col = 26 * $col + ord($2) - ord('a') + 1 if $2;

   return $col;

}


#
# $coord = NumberToCol($col)
#
# Turns 2 into B. The default is 1.
#

sub NumberToCol {

   my $col = shift @_;

   $col = $col > 1 ? $col : 1;

   my $col_high = int(($col - 1) / 26);
   my $col_low = ($col - 1) % 26;

   my $coord = chr(ord('A') + $col_low);
   $coord = chr(ord('A') + $col_high - 1) . $coord if $col_high;

   return $coord;

}


# # # # # # # # # #
# EncodeForSave($string)
#
# Returns $estring where :, \n, and \ are escaped
# 

sub EncodeForSave {
   my $string = shift @_;

   $string =~ s/\\/\\b/g; # \ to \b
   $string =~ s/:/\\c/g; # : to \c
   $string =~ s/\n/\\n/g; # line end to \n

   return $string;
}


# # # # # # # # # #
# DecodeFromSave($string)
#
# Returns $estring with \c, \n, \b and \\ un-escaped
# 

sub DecodeFromSave {
   my $string = shift @_;

   $string =~ s/\\c/:/g;
   $string =~ s/\\n/\n/g;
   $string =~ s/\\b/\\/g;

   return $string;
}


# # # # # # # # # #
# SpecialChars($string)
#
# Returns $estring where &, <, >, " are HTML escaped
# 

sub SpecialChars {
   my $string = shift @_;

   $string =~ s/&/&amp;/g;
   $string =~ s/</&lt;/g;
   $string =~ s/>/&gt;/g;
   $string =~ s/"/&quot;/g;

   return $string;
}


# # # # # # # # # #
# SpecialCharsNL($string)
#
# Returns $estring where &, <, >, ", and LF are HTML escaped, CR's are removed
# 

sub SpecialCharsNL {
   my $string = shift @_;

   $string =~ s/&/&amp;/g;
   $string =~ s/</&lt;/g;
   $string =~ s/>/&gt;/g;
   $string =~ s/"/&quot;/g;
   $string =~ s/\r//gs;
   $string =~ s/\n/&#10;/gs;

   return $string;
}


# # # # # # # # # #
# URLEncode($string)
#
# Returns $estring with special chars URL encoded
#
# Based on Mastering Regular Expressions, Jeffrey E. F. Friedl, additional legal characters added
# 

sub URLEncode {
   my $string = shift @_;

   $string =~ s!([^a-zA-Z0-9_\-;/?:@=#.])!sprintf('%%%02X', ord($1))!ge;
   $string =~ s/%26/{{amp}}/gs; # let ampersands in URLs through -- convert to {{amp}}

   return $string;
}


# # # # # # # # # #
# URLEncodePlain($string)
#
# Returns $estring with special chars URL encoded for sending to others by HTTP, not publishing
#
# Based on Mastering Regular Expressions, Jeffrey E. F. Friedl, additional legal characters added
# 

sub URLEncodePlain {
   my $string = shift @_;

   $string =~ s!([^a-zA-Z0-9_\-/?:@=#.])!sprintf('%%%02X', ord($1))!ge;

   return $string;
}


# # # # # # # # # #
# estring = ExpandMarkup($string, $sheet, $options)
#
# Returns $estring with expanded simple wikitext - used for error messages
# 

sub ExpandMarkup {

   my ($estring, $sheet, $options) = @_;

   return $estring;

   }


# # # # # # # # # #
# estring = ExpandWikitext($estring, $sheet, $options, $valueformat)
#
# Returns $estring with expanded wikitext - used for cell data
# 

sub ExpandWikitext {

   my ($estring, $sheet, $options, $valueformat) = @_;

   # handle the cases of text-wiki...

   if ($valueformat =~ m/^text-wiki(.+)$/) { # something more than just plain wikitext
      my $subtype = $1;
      my $result = "";

      if ($subtype eq "-pagelink") { # not very secure or robust -- just a demo
         if ($estring =~ m/\[(\S+)\s+(.+)\]$/) {
            my $url = URLEncode("$1");
            my $scestring = SpecialChars($2);
            $result = qq!<a href="?pagename=$url">$scestring</a>!;
            }
         elsif ($estring =~ m/\[(.*)\]$/) {
            my $url = URLEncode("$1");
            my $scestring = SpecialChars($1);
            $result = qq!<a href="?pagename=$url">$scestring</a>!;
            }
         else {
            my $url = URLEncode($estring);
            my $scestring = SpecialChars($estring);
            $result = qq!<a href="?pagename=$url">$scestring</a>!;
            }
         }        
      else {
         $result = SpecialChars("$subtype: $estring");
         }

      return $result || "&nbsp;";

      }

   $estring =~ s!\[(http:.+?)\s+(.+?)\]!'{{lt}}a href={{quot}}' . URLEncode("$1") . "{{quot}}{{gt}}$2\{{lt}}/a{{gt}}"!egs; # Wiki-style links
   $estring =~ s!\[link:(.+?)\s+(.+?)\:link]!'{{lt}}a href={{quot}}' . URLEncode("$1") . "{{quot}}{{gt}}$2\{{lt}}/a{{gt}}"!egs; # [link:url text:link] to link
   $estring =~ s!\[popup:(.+?)\s+(.+?)\:popup]!'{{lt}}a href={{quot}}' . URLEncode("$1") . "{{quot}} target={{quot}}_blank{{quot}}{{gt}}$2\{{lt}}/a{{gt}}"!egs; # [popup:url text:popup] to link with popup result
   $estring =~ s!\[image:(.+?)\s+(.+?)\:image]!'{{lt}}img src={{quot}}' . URLEncode("$1") . '{{quot}} alt={{quot}}' . SpecialCharsNL("$2") . '{{quot}}{{gt}}'!egs; # [image:url alt-text:image] for images
#   $string =~ s!\[page:(.+?)(\s+(.+?))?]!wiki_page_command($1,$3, $linkstyle)!egs; # [page:pagename text] to link to other pages on this site

   # Convert &, <, >, "

   $estring = SpecialChars($estring);

   # Multi-line text has additional formatting options ignored for single line

   if ($estring =~ m/\n/) {
      my ($str, @closingtag);
      foreach my $line (split /\n/, $estring) { # do things on a line-by-line basis
         $line =~ s/\r//g;
         if ($line =~ m/^([\*|#|;]{1,5})\s{0,1}(.+)$/) { # do list items
            my $lnest = length($1);
            my $lchr = substr($1,-1);
            my $ltype;
            if ($lnest > @closingtag) {
               for (my $i=@closingtag; $i<$lnest; $i++) {
                  if ($lchr eq '*') {
                     $ltype = "ul";
                     }
                  elsif ($lchr eq '#') {
                     $ltype = 'ol';
                     }
                  else {
                     $ltype = 'dl';
                     }
                  $str .= "<$ltype>";
                  push @closingtag, "</$ltype>";
                  }
               }
            elsif ($lnest < @closingtag) {
               for (my $i=@closingtag; $i>$lnest; $i--) {
                  $str .= pop @closingtag;
                  }
               }
            if ($lchr eq ';') {
               my $rest = $2;
               if ($rest =~ m/\s*(.*?):(.*)$/) {
                  $str .= "<dt>$1</dt><dd>$2</dd>";
                  }
               else {
                  $str .= "<dt>$rest</dt>";
                  }
               }
            else {
               $str .= "<li>$2</li>";
               }
            next;
            }
         while (@closingtag) {
            $str .= pop @closingtag;
            }
         if ($line =~ m/^(={1,5})\s(.+)\s\1$/) { # = heading =, with equal number of equals on both sides
            my $neq = length($1);
            $str .= "<h$neq>$2</h$neq>";
            next;
            }
         if ($line =~ m/^(:{1,5})\s{0,1}(.+)$/) { # indent 20pts for each :
            my $nindent = length($1) * 20;
            $str .= "<div style=\"padding-left:${nindent}pt;\">$2</div>";
            next;
            }

         $str .= "$line\n";
         }
      while (@closingtag) { # just in case any left at the end
         $str .= pop @closingtag;
         }
      $estring = $str;
      }

   $estring =~ s/\n/<br>/g;  # Line breaks are preserved
   $estring =~ s/('*)'''(.*?)'''/$1<b>$2<\/b>/gs; # Wiki-style bold/italics
   $estring =~ s/''(.*?)''/<i>$1<\/i>/gs;
   $estring =~ s/\[b:(.+?)\:b]/<b>$1<\/b>/gs; # [b:text:b] for bold
   $estring =~ s/\[i:(.+?)\:i]/<i>$1<\/i>/gs; # [i:text:i] for italic
   $estring =~ s/\[quote:(.+?)\:quote]/<blockquote>$1<\/blockquote>/gs; # [quote:text:quote] to indent
   $estring =~ s/\{\{amp}}/&/gs; # {{amp}} for ampersand
   $estring =~ s/\{\{lt}}/</gs; # {{lt}} for less than
   $estring =~ s/\{\{gt}}/>/gs; # {{gt}} for greater than
   $estring =~ s/\{\{quot}}/"/gs; # {{quot}} for quote
   $estring =~ s/\{\{lbracket}}/[/gs; # {{lbracket}} for left bracket
   $estring =~ s/\{\{rbracket}}/]/gs; # {{rbracket}} for right bracket
   $estring =~ s/\{\{lbrace}}/{/gs; # {{lbrace}} for brace

   return $estring;

   }

# * * * * * * * * * * * * * * * * * * * *
#
# Number Formatting code from SocialCalc 1.1.0
#
# * * * * * * * * * * * * * * * * * * * *

   our %WKCStrings = (
"sheetdefaultlayoutstyle" => "padding:2px 2px 1px 2px;vertical-align:top;",
"sheetdefaultfontfamily" => "Verdana,Arial,Helvetica,sans-serif",
"linkformatstring" => '<span style="font-size:smaller;text-decoration:none !important;background-color:#66B;color:#FFF;">Link</span>', # you could make this an img tag if desired:
#"linkformatstring" => '<img border="0" src="http://www.domain.com/link.gif">',
"defaultformatdt" => 'd-mmm-yyyy h:mm:ss',
"defaultformatd" => 'd-mmm-yyyy',
"defaultformatt" => '[h]:mm:ss',
"displaytrue" => 'TRUE', # how TRUE shows when rendered
"displayfalse" => 'FALSE',
      );

# # # # # # # # #
#
# $displayvalue = FormatValueForDisplay($sheet, $value, $coord, $options)
#
# # # # # # # # #

sub FormatValueForDisplay {

   my ($sheet, $value, $coord, $options) = @_;

   my $cell = $sheet->{cells}{$coord};
   my $sheetattribs = $sheet->{attribs};

   my ($valueformat, $has_parens, $has_commas, $valuetype, $valuesubtype);

   # Get references to the parts

   my $displayvalue = $value;

   my $valuetype = $cell->{valuetype}; # get type of value to determine formatting
   my $valuesubtype = substr($valuetype,1);
   $valuetype = substr($valuetype,0,1);

   if ($cell->{errors} || $valuetype eq "e") {
      $displayvalue = ExpandMarkup($cell->{errors}, $sheet, $options) || $valuesubtype || "Error in cell";
      return $displayvalue;
      }

   if ($valuetype eq "t") {
      $valueformat = $sheet->{valueformats}->[($cell->{textvalueformat} || $sheetattribs->{defaulttextvalueformat})] || "";
      if ($valueformat eq "formula") {
         if ($cell->{datatype} eq "f") {
            $displayvalue = SpecialChars("=$cell->{formula}") || "&nbsp;";
            }
         elsif ($cell->{datatype} eq "c") {
            $displayvalue = SpecialChars("'$cell->{formula}") || "&nbsp;";
            }
         else {
            $displayvalue = SpecialChars("'$displayvalue") || "&nbsp;";
            }
         return $displayvalue;
         }
      $displayvalue = format_text_for_display($displayvalue, $cell->{valuetype}, $valueformat, $sheet, $options);
      }

   elsif ($valuetype eq "n") {
      $valueformat = $cell->{nontextvalueformat};
      if (length($valueformat) == 0) { # "0" is a legal value format
         $valueformat = $sheetattribs->{defaultnontextvalueformat};
         }
      $valueformat = $sheet->{valueformats}->[$valueformat];
      if (length($valueformat) == 0) {
         $valueformat = "";
         }
      $valueformat = "" if $valueformat eq "none";
      if ($valueformat eq "formula") {
         if ($cell->{datatype} eq "f") {
            $displayvalue = SpecialChars("=$cell->{formula}") || "&nbsp;";
            }
         elsif ($cell->{datatype} eq "c") {
            $displayvalue = SpecialChars("'$cell->{formula}") || "&nbsp;";
            }
         else {
            $displayvalue = SpecialChars("'$displayvalue") || "&nbsp;";
            }
         return $displayvalue;
         }
      elsif ($valueformat eq "forcetext") {
         if ($cell->{datatype} eq "f") {
            $displayvalue = SpecialChars("=$cell->{formula}") || "&nbsp;";
            }
         elsif ($cell->{datatype} eq "c") {
            $displayvalue = SpecialChars($cell->{formula}) || "&nbsp;";
            }
         else {
            $displayvalue = SpecialChars($displayvalue) || "&nbsp;";
            }
         return $displayvalue;
         }
      $displayvalue = format_number_for_display($displayvalue, $cell->{valuetype}, $valueformat);
      }
   else { # unknown type - probably blank
      $displayvalue = "&nbsp;";
      }

   return $displayvalue;

   }


# # # # # # # # #
#
# $displayvalue = format_text_for_display($rawvalue, $valuetype, $valueformat, $sheetdata, $options)
#
# # # # # # # # #

sub format_text_for_display {

   my ($rawvalue, $valuetype, $valueformat, $sheet, $options) = @_;

   my $valuesubtype = substr($valuetype,1);

   my $displayvalue = $rawvalue;

   $valueformat = "" if $valueformat eq "none";
   $valueformat = "" unless $valueformat =~ m/^(text-|custom|hidden)/;
   if (!$valueformat || $valueformat eq "General") { # determine format from type
      $valueformat = "text-html" if ($valuesubtype eq "h");
      $valueformat = "text-wiki" if ($valuesubtype eq "w");
      $valueformat = "text-link" if ($valuesubtype eq "l");
      $valueformat = "text-plain" unless $valuesubtype;
      }
   if ($valueformat eq "text-html") { # HTML - output as it as is
      ;
      }
   elsif ($valueformat =~ m/^text-wiki/) { # wiki text
      $displayvalue = ExpandWikitext($displayvalue, $sheet, $options, $valueformat); # do wiki markup
      }
   elsif ($valueformat eq "text-url") { # text is a URL for a link with optional description
      my $dvsc = SpecialChars($displayvalue);
      my $dvue = URLEncode($displayvalue);
      $dvue =~ s/\Q{{amp}}/%26/g;
      $displayvalue = qq!<a href="$dvue">$dvsc</a>!;
      }
   elsif ($valueformat eq "text-link") { #  more extensive link capabilities for regular web links
      $displayvalue = ExpandTextLink($displayvalue, $sheet, $options, $valueformat);
      }
   elsif ($valueformat eq "text-image") { # text is a URL for an image
      my $dvue = URLEncode($displayvalue);
      $dvue =~ s/\Q{{amp}}/%26/g;
      $displayvalue = qq!<img src="$dvue">!;
      }
   elsif ($valueformat =~ m/^text-custom\:/) { # construct a custom text format: @r = text raw, @s = special chars, @u = url encoded
      my $dvsc = SpecialChars($displayvalue); # do special chars
      $dvsc =~ s/  /&nbsp; /g; # keep multiple spaces
      $dvsc =~ s/\n/<br>/g;  # keep line breaks
      my $dvue = URLEncode($displayvalue);
      $dvue =~ s/\Q{{amp}}/%26/g;
      my %textval;
      $textval{r} = $displayvalue;
      $textval{s} = $dvsc;
      $textval{u} = $dvue;
      $displayvalue = $valueformat;
      $displayvalue =~ s/^text-custom\://;
      $displayvalue =~ s/@(r|s|u)/$textval{$1}/ge;
      }
   elsif ($valueformat =~ m/^custom/) { # custom
      $displayvalue = SpecialChars($displayvalue); # do special chars
      $displayvalue =~ s/  /&nbsp; /g; # keep multiple spaces
      $displayvalue =~ s/\n/<br>/g;  # keep line breaks
      $displayvalue .= " (custom format)";
      }
   elsif ($valueformat eq "hidden") {
      $displayvalue = "&nbsp;";
      }
   else { # plain text
      $displayvalue = SpecialChars($displayvalue); # do special chars
      $displayvalue =~ s/  /&nbsp; /g; # keep multiple spaces
      $displayvalue =~ s/\n/<br>/g;  # keep line breaks
      }

   return $displayvalue;

   }


# # # # # # # # #
#
# $displayvalue = ExpandTextLink($displayvalue, $sheet, $options, $valueformat)
#
# Given: url = http://www.someurl.com/more, desc = Some descriptive text
#
# Takes the following:
#
#    url => "url" as a link to url or other html (depending on SocialCalc.Constants)
#    <url> => "Link" or image as link (depending on SocialCalc.Constants)
#    desc<url> => "desc" as a link to url
#    "desc"<url> => "desc" as a link to url
#    <<>> instead of <> => target="_blank" (new window) even when not editing
#
# # # # # # # # #


sub ExpandTextLink {

   my ($displayvalue, $sheet, $options, $valueformat) = @_;

   my ($desc, $url);

   my $parts = ParseCellLinkText($displayvalue);

   if ($parts->{desc}) {
      $desc = SpecialChars($parts->{desc});
      }
   else {
      $desc = $WKCStrings{linkformatstring};
      }

   if (length($displayvalue) > 7 && lc(substr($displayvalue,0,7)) eq "http://" 
      && substr($displayvalue,-1,1) ne ">") {
      $desc = substr($desc,7); # remove http:// unless explicit
      }

   my $tb = ($parts->{newwin} || $options->{newwinlinks}) ? ' target="_blank"' : "";

   if ($parts->{pagename}) {
      $url = MakePageLink($parts->{pagename}, $parts->{workspace}, $options, $valueformat);
      }
   else {
      $url = URLEncode($parts->{url});
      }

   my $str = qq!<a href="$url"$tb>$desc</a>!;

   return $str;

   }


# # # # # # # # #
#
# $result = ParseCellLinkText($str)
#
# Given: url = http://www.someurl.com/more, desc = Some descriptive text
#
# Takes the following:
#
#    url
#    <url>
#    desc<url>
#    "desc"<url>
#    <<>> instead of <> => target="_blank" (new window)
#
#    [page name]
#    "desc"[page name]
#    desc[page name]
#    {workspace name [page name]}
#    "desc"{workspace name [page name]}
#    [[]] instead of [] => target="_blank" (new window)
#
# Returns: {url => url, desc => desc, newwin => t/f, pagename => pagename, workspace => workspace}
#
# # # # # # # # #

sub ParseCellLinkText {

   my ($str) = @_;

   my $result = {url => "", desc => "", newwin => 0, pagename => "", workspace => ""};

   if ($str !~ /<.*>$/ && $str !~ /\[.*\]$/ && $str !~ /\{.*\[.*\]\}$/) { # plain url
      $result->{url} = $str;
      $result->{desc} = $str;
      return $result;
      }

   my $desc;

   if ($str =~ /^(.*)\[\[(.*?)\]\]$/) {
      $desc = $1;
      $result->{pagename} = $2;
      $result->{newwin} = 1;
      }
   elsif ($str =~ /^(.*)\[(.*?)\]$/) {
      $desc = $1;
      $result->{pagename} = $2;
      }
   elsif ($str =~ /^(.*)\{(.*?)(\s{0,1})\[\[(.*?)\]\]\}$/) {
      $desc = $1;
      $result->{workspace} = $2;
      $result->{pagename} = $4;
      $result->{newwin} = 1;
      }
   elsif ($str =~ /^(.*)\{(.*?)(\s{0,1})\[(.*?)\]\}$/) {
      $desc = $1;
      $result->{workspace} = $2;
      $result->{pagename} = $4;
      }
   elsif ($str =~ /^(.*)<<(.*?)>>$/) {
      $desc = $1;
      $result->{url} = $2;
      $result->{newwin} = 1;
      }
   else {
      $str =~ /^(.*)<(.*?)>$/;
      $desc = $1;
      $result->{url} = $2;
      }

   if ($desc =~ /^"(.*)"$/) {
      $desc = $1;
      }

   $result->{desc} = $desc;

   return $result;

   }


# # # # # # # # # #
# $estring = MakePageLink($pagename, $workspacename, $options, $valueformat);
#
# Returns $estring with expanded link to page in workspace
# 

sub MakePageLink {

   my ($pagename, $workspacename, $options, $valueformat) = @_;

   my $url = "";

   if ($workspacename) {
      $url = "?&workspace=" . URLEncode($workspacename) . "&pagename=" . URLEncode($pagename);
      }
   else {
      $url = "?&pagename=" . URLEncode($pagename);
      }

   return $url;

   }


# # # # # # # # #
#
# $displayvalue = format_number_for_display($rawvalue, $valuetype, $valueformat)
#
# # # # # # # # #

sub format_number_for_display {

   my ($rawvalue, $valuetype, $valueformat) = @_;

   my ($has_parens, $has_commas);

   my $displayvalue = $rawvalue;
   my $valuesubtype = substr($valuetype,1);

   if ($valueformat eq "Auto" || length($valueformat) == 0) { # cases with default format
      if ($valuesubtype eq "%") { # will display a % character
         $valueformat = "#,##0.0%";
         }
      elsif ($valuesubtype eq '$') {
         $valueformat = '[$]#,##0.00';
         }
      elsif ($valuesubtype eq 'dt') {
         $valueformat = $WKCStrings{"defaultformatdt"};
         }
      elsif ($valuesubtype eq 'd') {
         $valueformat = $WKCStrings{"defaultformatd"};
         }
      elsif ($valuesubtype eq 't') {
         $valueformat = $WKCStrings{"defaultformatt"};
         }
      elsif ($valuesubtype eq 'l') {
         $valueformat = 'logical';
         }
      else {
         $valueformat = "General";
         }
      }

   if ($valueformat eq "logical") { # do logical format
      return $rawvalue ? $WKCStrings{"displaytrue"} : $WKCStrings{"displayfalse"};
      }

   if ($valueformat eq "hidden") { # do hidden format
      return "&nbsp;";
      }

   # Use format

   return format_number_with_format_string($rawvalue, $valueformat);

   }


# # # # # # # # #
#
# $juliandate = convert_date_gregorian_to_julian($year, $month, $day)
#
# From: http://aa.usno.navy.mil/faq/docs/JD_Formula.html
# Uses: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM, Vol. 11, No. 10 (October, 1968).
# Translated from the FORTRAN
#
#      I= YEAR
#      J= MONTH
#      K= DAY
#C
#      JD= K-32075+1461*(I+4800+(J-14)/12)/4+367*(J-2-(J-14)/12*12)
#     2    /12-3*((I+4900+(J-14)/12)/100)/4
#
# # # # # # # # #

sub convert_date_gregorian_to_julian {

   my ($year, $month, $day) = @_;

   my $juliandate= $day-32075+int(1461*($year+4800+int(($month-14)/12))/4);
   $juliandate += int(367*($month-2-int(($month-14)/12)*12)/12);
   $juliandate = $juliandate -int(3*int(($year+4900+int(($month-14)/12))/100)/4);

   return $juliandate;

}


# # # # # # # # #
#
# ($year, $month, $day) = convert_date_julian_to_gregorian($juliandate)
#
# From: http://aa.usno.navy.mil/faq/docs/JD_Formula.html
# Uses: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM, Vol. 11, No. 10 (October, 1968).
# Translated from the FORTRAN
#
# # # # # # # # #

sub convert_date_julian_to_gregorian {

   my $juliandate = shift @_;

   my ($L, $N, $I, $J, $K);

   $L = $juliandate+68569;
   $N = int(4*$L/146097);
   $L = $L-int((146097*$N+3)/4);
   $I = int(4000*($L+1)/1461001);
   $L = $L-int(1461*$I/4)+31;
   $J = int(80*$L/2447);
   $K = $L-int(2447*$J/80);
   $L = int($J/11);
   $J = $J+2-12*$L;
   $I = 100*($N-49)+$I+$L;

   return ($I, $J, $K);

}

