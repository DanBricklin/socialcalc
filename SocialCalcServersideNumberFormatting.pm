#
# SocialCalcServersideNumberFormatting.pm
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

   require Exporter;
   our @ISA = qw(Exporter);
   our @EXPORT = qw(FormatValueForDisplay);
   our $VERSION = '1.0.0';

   #
   # CONSTANTS
   #

   our $julian_offset = 2415019;
   our $seconds_in_a_day = 24 * 60 * 60;
   our $seconds_in_an_hour = 60 * 60;

# * * * * * * * * * * * * * * * * * * * *
#
# Number Formatting code from SocialCalc 1.1.0
#
# * * * * * * * * * * * * * * * * * * * *

   our %NFStrings = (
"decimalchar" => ".",
"separatorchar" => ",",
"currencychar" => '$',
"daynames" => "Sunday Monday Tuesday Wednesday Thursday Friday Saturday",
"daynames3" => "Sun Mon Tue Wed Thu Fri Sat ",
"monthnames3" => "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec",
"monthnames" => "January February March April May June July August September October November December",
      );


# # # # # # # # #
#
# $result = format_number_with_format_string($value, $format_string, $currency_char)
#
# Use a format string to format a numeric value. Returns a string with the result.
# This is a subset of the normal styles accepted by many other spreadsheets, without fractions, E format, and @,
# and with any number of comparison fields and with [style=style-specification] (e.g., [style=color:red])
#
# # # # # # # # #

   my %allowedcolors = (BLACK => "#000000", BLUE => "#0000FF", CYAN => "#00FFFF", GREEN => "#00FF00", MAGENTA => "#FF00FF",
                        RED => "#FF0000", WHITE => "#FFFFFF", YELLOW => "#FFFF00");

   my %alloweddates = (H => "h]", M => "m]", MM => "mm]", "S" => "s]", "SS" => "ss]");

   my %format_definitions;
   my $cmd_copy = 1;
   my $cmd_color = 2;
   my $cmd_integer_placeholder = 3;
   my $cmd_fraction_placeholder = 4;
   my $cmd_decimal = 5;
   my $cmd_currency = 6;
   my $cmd_general = 7;
   my $cmd_separator = 8;
   my $cmd_date = 9;
   my $cmd_comparison = 10;
   my $cmd_section = 11;
   my $cmd_style = 12;

sub format_number_with_format_string {

   my ($rawvalue, $format_string, $currency_char) = @_;

   $currency_char ||= '$';

   my ($op, $operandstr, $fromend, $cval, $operandstrlc);
   my ($yr, $mn, $dy, $hrs, $mins, $secs, $ehrs, $emins, $esecs, $ampmstr);
   my $result;

   my $value = $rawvalue+0; # get a working copy that's numeric

   my $negativevalue = $value < 0 ? 1 : 0; # determine sign, etc.
   $value = -$value if $negativevalue;
   my $zerovalue = $value == 0 ? 1 : 0;

   parse_format_string(\%format_definitions, $format_string); # make sure format is parsed
   my $thisformat = $format_definitions{$format_string}; # Get format structure

   return "Format error!" unless $thisformat;

   my $section = (scalar @{$thisformat->{sectioninfo}}) - 1; # get number of sections - 1

   if ($thisformat->{hascomparison}) { # has comparisons - determine which section
      $section = 0; # set to which section we will use
      my $gotcomparison = 0; # this section has no comparison
      for (my $cpos; ;$cpos++) { # scan for comparisons
         $op = $thisformat->{operators}->[$cpos];
         $operandstr = $thisformat->{operands}->[$cpos]; # get next operator and operand
         if (!$op) { # at end with no match
            if ($gotcomparison) { # if comparison but no match
               $format_string = "General"; # use default of General
               parse_format_string(\%format_definitions, $format_string);
               $thisformat = $format_definitions{$format_string};
               $section = 0;
               }
            last; # if no comparision, matchines on this section
            }
         if ($op == $cmd_section) { # end of section
            if (!$gotcomparison) { # no comparison, so it's a match
               last;
               }
            $gotcomparison = 0;
            $section++; # check out next one
            next;
            }
         if ($op == $cmd_comparison) { # found a comparison - do we meet it?
            my ($compop, $compval) = split(/:/, $operandstr, 2);
            $compval = 0+$compval;
            if (($compop eq "<" && $rawvalue < $compval) ||
                ($compop eq "<=" && $rawvalue <= $compval) ||
                ($compop eq "<>" && $rawvalue != $compval) ||
                ($compop eq ">=" && $rawvalue >= $compval) ||
                ($compop eq ">" && $rawvalue > $compval)) { # a match
               last;
               }
            $gotcomparison = 1;
            }
         }
      }
   elsif ($section > 0) { # more than one section (separated by ";")
      if ($section == 1) { # two sections
         if ($negativevalue) {
            $negativevalue = 0; # sign will provided by section, not automatically
            $section = 1; # use second section for negative values
            }
         else {
            $section = 0; # use first for all others
            }
         }
      elsif ($section == 2) { # three sections
         if ($negativevalue) {
            $negativevalue = 0; # sign will provided by section, not automatically
            $section = 1; # use second section for negative values
            }
         elsif ($zerovalue) {
            $section = 2; # use third section for zero values
            }
         else {
            $section = 0; # use first for positive
            }
         }
      }

   # Get values for our section
   my ($sectionstart, $integerdigits, $fractiondigits, $commas, $percent, $thousandssep) =
      @{%{$thisformat->{sectioninfo}->[$section]}}{qw(sectionstart integerdigits fractiondigits commas percent thousandssep)};

   if ($commas > 0) { # scale by thousands
      for (my $i=0; $i<$commas; $i++) {
         $value /= 1000;
         }
      }
   if ($percent > 0) { # do percent scaling
      for (my $i=0; $i<$percent; $i++) {
         $value *= 100;
         }
      }

   my $decimalscale = 1; # cut down to required number of decimal digits
   for (my $i=0; $i<$fractiondigits; $i++) {
      $decimalscale *= 10;
      }
   my $scaledvalue = int($value * $decimalscale + 0.5);
   $scaledvalue = $scaledvalue / $decimalscale;

   $negativevalue = 0 if ($scaledvalue == 0 && ($fractiondigits || $integerdigits)); # no "-0" unless using multiple sections or General

   my $strvalue = "$scaledvalue"; # convert to string
   if ($strvalue =~ m/e/) { # converted to scientific notation
      return "$rawvalue"; # Just return plain converted raw value
      }
   $strvalue =~ m/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/; # get integer and fraction as character arrays
   my $integervalue = $1;
   $integervalue = "" if ($integervalue == 0);
   my @integervalue = split(//, $integervalue);
   my $fractionvalue = $2;
   $fractionvalue = "" if ($fractionvalue == 0);
   my @fractionvalue = split(//, $fractionvalue);

   if ($thisformat->{sectioninfo}->[$section]->{hasdate}) { # there are date placeholders
      if ($rawvalue < 0) { # bad date
         return "??-???-??&nbsp;??:??:??";
         }
      my $startval = ($rawvalue-int($rawvalue)) * $seconds_in_a_day; # get date/time parts
      my $estartval = $rawvalue * $seconds_in_a_day; # do elapsed time version, too
      $hrs = int($startval / $seconds_in_an_hour);
      $ehrs = int($estartval / $seconds_in_an_hour);
      $startval = $startval - $hrs * $seconds_in_an_hour;
      $mins = int($startval / 60);
      $emins = int($estartval / 60);
      $secs = $startval - $mins * 60;
      $decimalscale = 1; # round appropriately depending if there is ss.0
      for (my $i=0; $i<$fractiondigits; $i++) {
         $decimalscale *= 10;
         }
      $secs = int($secs * $decimalscale + 0.5);
      $secs = $secs / $decimalscale;
      $esecs = int($estartval * $decimalscale + 0.5);
      $esecs = $esecs / $decimalscale;
      if ($secs >= 60) { # handle round up into next second, minute, etc.
         $secs = 0;
         $mins++; $emins++;
         if ($mins >= 60) {
            $mins = 0;
            $hrs++; $ehrs++;
            if ($hrs >= 24) {
               $hrs = 0;
               $rawvalue++;
               }
            }
         }
      @fractionvalue = split(//, $secs-int($secs)); # for "hh:mm:ss.00"
      shift @fractionvalue; shift @fractionvalue;
      ($yr, $mn, $dy) = convert_date_julian_to_gregorian(int($rawvalue+$julian_offset));

      my $minOK; # says "m" can be minutes
      my $mspos = $sectionstart; # m scan position in ops
      for ( ; ; $mspos++) { # scan for "m" and "mm" to see if any minutes fields, and am/pm
         $op = $thisformat->{operators}->[$mspos];
         $operandstr = $thisformat->{operands}->[$mspos]; # get next operator and operand
         last unless $op; # don't go past end
         last if $op == $cmd_section;
         if ($op == $cmd_date) {
            if ((lc($operandstr) eq "am/pm" || lc($operandstr) eq "a/p") && !$ampmstr) {
               if ($hrs >= 12) {
                  $hrs -= 12;
                  $ampmstr = lc($operandstr) eq "a/p" ? "P" : "PM";
                  }
               else {
                  $ampmstr = lc($operandstr) eq "a/p" ? "A" : "AM";
                  }
               $ampmstr = lc $ampmstr if $operandstr !~ m/$ampmstr/;
               }
            if ($minOK && ($operandstr eq "m" || $operandstr eq "mm")) {
               $thisformat->{operands}->[$mspos] .= "in"; # turn into "min" or "mmin"
               }
            if (substr($operandstr,0,1) eq "h") {
               $minOK = 1; # m following h or hh or [h] is minutes not months
               }
            else {
               $minOK = 0;
               }
            }
         elsif ($op != $cmd_copy) { # copying chars can be between h and m
            $minOK = 0;
            }
         }
      $minOK = 0;
      for (--$mspos; ; $mspos--) { # scan other way for s after m
         $op = $thisformat->{operators}->[$mspos];
         $operandstr = $thisformat->{operands}->[$mspos]; # get next operator and operand
         last unless $op; # don't go past end
         last if $op == $cmd_section;
         if ($op == $cmd_date) {
            if ($minOK && ($operandstr eq "m" || $operandstr eq "mm")) {
               $thisformat->{operands}->[$mspos] .= "in"; # turn into "min" or "mmin"
               }
            if ($operandstr eq "ss") {
               $minOK = 1; # m before ss is minutes not months
               }
            else {
               $minOK = 0;
               }
            }
         elsif ($op != $cmd_copy) { # copying chars can be between ss and m
            $minOK = 0;
            }
         }
      }

   my $integerdigits2 = 0; # init counters, etc.
   my $integerpos = 0;
   my $fractionpos = 0;
   my $textcolor = "";
   my $textstyle = "";
   my $separatorchar = $NFStrings{"separatorchar"};
   $separatorchar =~ s/ /&nbsp;/g;
   my $decimalchar = $NFStrings{"decimalchar"};
   $decimalchar =~ s/ /&nbsp;/g;

   my $oppos = $sectionstart;

   while ($op = $thisformat->{operators}->[$oppos]) { # execute format
      $operandstr = $thisformat->{operands}->[$oppos++]; # get next operator and operand
      if ($op == $cmd_copy) { # put char in result
         $result .= $operandstr;
         }

      elsif ($op == $cmd_color) { # set color
         $textcolor = $operandstr;
         }

      elsif ($op == $cmd_style) { # set style
         $textstyle = $operandstr;
         }

      elsif ($op == $cmd_integer_placeholder) { # insert number part
         if ($negativevalue) {
            $result .= "-";
            $negativevalue = 0;
            }
         $integerdigits2++;
         if ($integerdigits2 == 1) { # first one
            if ((scalar @integervalue) > $integerdigits) { # see if integer wider than field
               for (;$integerpos < ((scalar @integervalue) - $integerdigits); $integerpos++) {
                  $result .= $integervalue[$integerpos];
                  if ($thousandssep) { # see if this is a separator position
                     $fromend = (scalar @integervalue) - $integerpos - 1;
                     if ($fromend > 2 && $fromend % 3 == 0) {
                        $result .= $separatorchar;
                        }
                     }
                  }
               }
            }
         if ((scalar @integervalue) < $integerdigits
             && $integerdigits2 <= $integerdigits - (scalar @integervalue)) { # field is wider than value
            if ($operandstr eq "0" || $operandstr eq "?") { # fill with appropriate characters
               $result .= $operandstr eq "0" ? "0" : "&nbsp;";
               if ($thousandssep) { # see if this is a separator position
                  $fromend = $integerdigits - $integerdigits2;
                  if ($fromend > 2 && $fromend % 3 == 0) {
                     $result .= $separatorchar;
                     }
                  }
               }
            }
         else { # normal integer digit - add it
            $result .= $integervalue[$integerpos];
            if ($thousandssep) { # see if this is a separator position
               $fromend = (scalar @integervalue) - $integerpos - 1;
               if ($fromend > 2 && $fromend % 3 == 0) {
                  $result .= $separatorchar;
                  }
               }
            $integerpos++;
            }
         }
      elsif ($op == $cmd_fraction_placeholder) { # add fraction part of number
         if ($fractionpos >= scalar @fractionvalue) {
            if ($operandstr eq "0" || $operandstr eq "?") {
               $result .= $operandstr eq "0" ? "0" : "&nbsp;";
               }
            }
         else {
            $result .= $fractionvalue[$fractionpos];
            }
         $fractionpos++;
         }

      elsif ($op == $cmd_decimal) { # decimal point
         if ($negativevalue) {
            $result .= "-";
            $negativevalue = 0;
            }
         $result .= $decimalchar;
         }

      elsif ($op == $cmd_currency) { # currency symbol
         if ($negativevalue) {
            $result .= "-";
            $negativevalue = 0;
            }
         $result .= $operandstr;
         }

      elsif ($op == $cmd_general) { # insert "General" conversion

         # *** Cut down number of significant digits to avoid floating point artifacts matching JavaScript:

         if ($value != 0) { # only if non-zero
            my $factor = 0.43429448190325181667 * log($value); # get magnitude as a power of 10 the same way as JavaScript
            if ($factor >= 0) { # as an integer (truncating down)
               $factor = int($factor);
               }
            else {
               my $mfactor = -int(-$factor);
               $factor = $factor == $mfactor ? $factor : $mfactor - 1;
               }
            $factor = 10 ** (13-$factor); # turn into scaling factor
            $value = int($factor * $value + 0.5) / $factor; # scale positive value, round, undo scaling
            }

         if ($negativevalue) {
            $result .= "-";
            }
         $strvalue = "$value"; # convert original value to string
         if ($strvalue =~ m/e/) { # converted to scientific notation
            $result .= "$strvalue";
            next;
            }
         $strvalue =~ m/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/;
         $integervalue = $1;
         $integervalue = "" if ($integervalue == 0);
         @integervalue = split(//, $integervalue);
         $fractionvalue = $2;
         $fractionvalue = "" if ($fractionvalue == 0);
         @fractionvalue = split(//, $fractionvalue);
         $integerpos = 0;
         $fractionpos = 0;
         if (scalar @integervalue) {
            for (;$integerpos < scalar @integervalue; $integerpos++) {
               $result .= $integervalue[$integerpos];
               if ($thousandssep) { # see if this is a separator position
                  $fromend = (scalar @integervalue) - $integerpos - 1;
                  if ($fromend > 2 && $fromend % 3 == 0) {
                     $result .= $separatorchar;
                     }
                  }
               }
             }
         else {
            $result .= "0";
            }
         if (scalar @fractionvalue) {
            $result .= $decimalchar;
            for (;$fractionpos < scalar @fractionvalue; $fractionpos++) {
               $result .= $fractionvalue[$fractionpos];
               }
            }
         }

      elsif ($op == $cmd_date) { # date placeholder
         $operandstrlc = lc $operandstr;
         if ($operandstrlc eq "y" || $operandstrlc eq "yy") {
            $result .= substr("$yr",-2);
            }
         elsif ($operandstrlc eq "yyyy") {
            $result .= "$yr";
            }
         elsif ($operandstrlc eq "d") {
            $result .= "$dy";
            }
         elsif ($operandstrlc eq "dd") {
            $cval = 1000 + $dy;
            $result .= substr("$cval", -2);
            }
         elsif ($operandstrlc eq "ddd") {
            $cval = int($rawvalue+6) % 7;
            $result .= (split(/ /, $NFStrings{"daynames3"}))[$cval];
            }
         elsif ($operandstrlc eq "dddd") {
            $cval = int($rawvalue+6) % 7;
            $result .= (split(/ /, $NFStrings{"daynames"}))[$cval];
            }
         elsif ($operandstrlc eq "m") {
            $result .= "$mn";
            }
         elsif ($operandstrlc eq "mm") {
            $cval = 1000 + $mn;
            $result .= substr("$cval", -2);
            }
         elsif ($operandstrlc eq "mmm") {
            $result .= (split(/ /, $NFStrings{"monthnames3"}))[$mn-1];
            }
         elsif ($operandstrlc eq "mmmm") {
            $result .= (split(/ /, $NFStrings{"monthnames"}))[$mn-1];
            }
         elsif ($operandstrlc eq "mmmmm") {
            $result .= substr((split(/ /, $NFStrings{"monthnames"}))[$mn-1], 0, 1);
            }
         elsif ($operandstrlc eq "h") {
            $result .= "$hrs";
            }
         elsif ($operandstrlc eq "h]") {
            $result .= "$ehrs";
            }
         elsif ($operandstrlc eq "mmin") {
            $cval = 1000 + $mins;
            $result .= substr("$cval", -2);
            }
         elsif ($operandstrlc eq "mm]") {
            if ($emins < 100) {
               $cval = 1000 + $emins;
               $result .= substr("$cval", -2);
               }
            else {
               $result .= "$emins";
               }
            }
         elsif ($operandstrlc eq "min") {
            $result .= "$mins";
            }
         elsif ($operandstrlc eq "m]") {
            $result .= "$emins";
            }
         elsif ($operandstrlc eq "hh") {
            $cval = 1000 + $hrs;
            $result .= substr("$cval", -2);
            }
         elsif ($operandstrlc eq "s") {
            $cval = int($secs);
            $result .= "$cval";
            }
         elsif ($operandstrlc eq "ss") {
            $cval = 1000 + int($secs);
            $result .= substr("$cval", -2);
            }
         elsif ($operandstrlc eq "am/pm" || $operandstrlc eq "a/p") {
            $result .= $ampmstr;
            }
         elsif ($operandstrlc eq "ss]") {
            if ($esecs < 100) {
               $cval = 1000 + int($esecs);
               $result .= substr("$cval", -2);
               }
            else {
               $cval = int($esecs);
               $result = "$cval";
               }
            }
         }

      elsif ($op == $cmd_section) { # end of section
         last;
         }

      elsif ($op == $cmd_comparison) { # ignore
         next;
         }

      else {
         $result .= "!! Parse error !!";
         }
      }

   if ($textcolor) {
      $result = qq!<span style="color:$textcolor;">$result</span>!;
      }
   if ($textstyle) {
      $result = qq!<span style="$textstyle;">$result</span>!;
      }

   return $result;
}

# # # # # # # # #
#
# parse_format_string(\%format_defs, $format_string)
#
# Takes a format string (e.g., "#,##0.00_);(#,##0.00)") and fills in %foramt_defs with the parsed info
#
# %format_defs
#    {"#,##0.0"}->{} # elements in the hash are one hash for each format
#       {operators}->[] # array of operators from parsing the format string (each a number)
#       {operands}->[] # array of corresponding operators (each usually a string)
#       {sectioninfo}->[] # one hash for each section of the format
#          {start}
#          {integerdigits}
#          {fractiondigits}
#          {commas}
#          {percent}
#          {thousandssep}
#          {hasdates}
#       {hascomparison} # true if any section has [<100], etc.
#
# # # # # # # # #

sub parse_format_string {

   my ($format_defs, $format_string) = @_;

   return if ($format_defs->{$format_string}); # already defined - nothing to do

   my $thisformat = {operators => [], operands => [], sectioninfo => [{}]}; # create info structure for this format
   $format_defs->{$format_string} = $thisformat; # add to other format definitions

   my $section = 0; # start with section 0
   my $sectioninfo = $thisformat->{sectioninfo}->[$section]; # get reference to info for current section
   $sectioninfo->{sectionstart} = 0; # position in operands that starts this section

   my @formatchars = split //, $format_string; # break into individual characters

   my $integerpart = 1; # start out in integer part
   my $lastwasinteger; # last char was an integer placeholder
   my $lastwasslash; # last char was a backslash - escaping following character
   my $lastwasasterisk; # repeat next char
   my $lastwasunderscore; # last char was _ which picks up following char for width
   my ($inquote, $quotestr); # processing a quoted string
   my ($inbracket, $bracketstr, $cmd); # processing a bracketed string
   my ($ingeneral, $gpos); # checks for characters "General"
   my $ampmstr; # checks for characters "A/P" and "AM/PM"
   my $indate; # keeps track of date/time placeholders

   foreach my $ch (@formatchars) { # parse
      if ($inquote) {
         if ($ch eq '"') {
            $inquote = 0;
            push @{$thisformat->{operators}}, $cmd_copy;
            push @{$thisformat->{operands}}, $quotestr;
            next;
            }
         $quotestr .= $ch;
         next;
         }
      if ($inbracket) {
         if ($ch eq ']') {
            $inbracket = 0;
            ($cmd, $bracketstr) = parse_format_bracket($bracketstr);
            if ($cmd == $cmd_separator) {
               $sectioninfo->{thousandssep} = 1; # explicit [,]
               next;
               }
            if ($cmd == $cmd_date) {
               $sectioninfo->{hasdate} = 1;
               }
            if ($cmd == $cmd_comparison) {
               $thisformat->{hascomparison} = 1;
               }
            push @{$thisformat->{operators}}, $cmd;
            push @{$thisformat->{operands}}, $bracketstr;
            next;
            }
         $bracketstr .= $ch;
         next;
         }
      if ($lastwasslash) {
         push @{$thisformat->{operators}}, $cmd_copy;
         push @{$thisformat->{operands}}, $ch;
         $lastwasslash = 0;
         next;
         }
      if ($lastwasasterisk) {
         push @{$thisformat->{operators}}, $cmd_copy;
         push @{$thisformat->{operands}}, $ch x 5;
         $lastwasasterisk = 0;
         next;
         }
      if ($lastwasunderscore) {
         push @{$thisformat->{operators}}, $cmd_copy;
         push @{$thisformat->{operands}}, "&nbsp;";
         $lastwasunderscore = 0;
         next;
         }
      if ($ingeneral) {
         if (substr("general", $ingeneral, 1) eq lc $ch) {
            $ingeneral++;
            if ($ingeneral == 7) {
               push @{$thisformat->{operators}}, $cmd_general;
               push @{$thisformat->{operands}}, $ch;
               $ingeneral = 0;
               }
            next;
            }
         $ingeneral = 0;
         }
      if ($indate) { # last char was part of a date placeholder
         if (substr($indate,0,1) eq $ch) { # another of the same char
            $indate .= $ch; # accumulate it
            next;
            }
         push @{$thisformat->{operators}}, $cmd_date; # something else, save date info
         push @{$thisformat->{operands}}, $indate;
         $sectioninfo->{hasdate} = 1;
         $indate = "";
         }
      if ($ampmstr) {
         $ampmstr .= $ch;
         if ("am/pm" =~ m/^$ampmstr/i || "a/p" =~ m/^$ampmstr/i) {
            if (("am/pm" eq lc $ampmstr) || ("a/p" eq lc $ampmstr)) {
               push @{$thisformat->{operators}}, $cmd_date;
               push @{$thisformat->{operands}}, $ampmstr;
               $ampmstr = "";
               }
            next;
            }
         $ampmstr = "";
         }
      if ($ch eq "#" || $ch eq "0" || $ch eq "?") { # placeholder
         if ($integerpart) {
            $sectioninfo->{integerdigits}++;
            if ($sectioninfo->{commas}) { # comma inside of integer placeholders
               $sectioninfo->{thousandssep} = 1; # any number is thousands separator
               $sectioninfo->{commas} = 0; # reset count of "thousand" factors
               }
            $lastwasinteger = 1;
            push @{$thisformat->{operators}}, $cmd_integer_placeholder;
            push @{$thisformat->{operands}}, $ch;
            }
         else {
            $sectioninfo->{fractiondigits}++;
            push @{$thisformat->{operators}}, $cmd_fraction_placeholder;
            push @{$thisformat->{operands}}, $ch;
            }
         }
      elsif ($ch eq ".") { # decimal point
         $lastwasinteger = 0;
         push @{$thisformat->{operators}}, $cmd_decimal;
         push @{$thisformat->{operands}}, $ch;
         $integerpart = 0;
         }
      elsif ($ch eq '$') { # currency char
         $lastwasinteger = 0;
         push @{$thisformat->{operators}}, $cmd_currency;
         push @{$thisformat->{operands}}, $ch;
         }
      elsif ($ch eq ",") {
         if ($lastwasinteger) {
            $sectioninfo->{commas}++;
            }
         else {
            push @{$thisformat->{operators}}, $cmd_copy;
            push @{$thisformat->{operands}}, $ch;
            }
         }
      elsif ($ch eq "%") {
         $lastwasinteger = 0;
         $sectioninfo->{percent}++;
         push @{$thisformat->{operators}}, $cmd_copy;
         push @{$thisformat->{operands}}, $ch;
         }
      elsif ($ch eq '"') {
         $lastwasinteger = 0;
         $inquote = 1;
         $quotestr = "";
         }
      elsif ($ch eq '[') {
         $lastwasinteger = 0;
         $inbracket = 1;
         $bracketstr = "";
         }
      elsif ($ch eq '\\') {
         $lastwasslash = 1;
         $lastwasinteger = 0;
         }
      elsif ($ch eq '*') {
         $lastwasasterisk = 1;
         $lastwasinteger = 0;
         }
      elsif ($ch eq '_') {
         $lastwasunderscore = 1;
         $lastwasinteger = 0;
         }
      elsif ($ch eq ";") {
         $section++; # start next section
         $thisformat->{sectioninfo}->[$section] = {}; # create a new section
         $sectioninfo = $thisformat->{sectioninfo}->[$section]; # set to point to the new section
         $sectioninfo->{sectionstart} = 1 + scalar @{$thisformat->{operators}}; # remember where it starts
         $integerpart = 1; # reset for new section
         $lastwasinteger = 0;
         push @{$thisformat->{operators}}, $cmd_section;
         push @{$thisformat->{operands}}, $ch;
         }
      elsif ((lc $ch) eq "g") {
         $ingeneral = 1;
         $lastwasinteger = 0;
         }
      elsif ((lc $ch) eq "a") {
         $ampmstr = $ch;
         $lastwasinteger = 0;
         }
      elsif ($ch =~ m/[dmyhHs]/) {
         $indate = $ch;
         }
      else {
         $lastwasinteger = 0;
         push @{$thisformat->{operators}}, $cmd_copy;
         push @{$thisformat->{operands}}, $ch;
         }
      }

   if ($indate) { # last char was part of unsaved date placeholder
      push @{$thisformat->{operators}}, $cmd_date; # save what we got
      push @{$thisformat->{operands}}, $indate;
      $sectioninfo->{hasdate} = 1;
      }

   return;

   }


# # # # # # # # #
#
# ($operator, $operand) = parse_format_bracket($bracketstr)
#
# # # # # # # # #

sub parse_format_bracket {

   my $bracketstr = shift @_;

   my ($operator, $operand);

   if (substr($bracketstr, 0, 1) eq '$') { # currency
      $operator = $cmd_currency;
      if ($bracketstr =~ m/^\$(.+?)(\-.+?){0,1}$/) {
         $operand = $1 || $NFStrings{"currencychar"} || '$';
         }
      else {
         $operand = substr($bracketstr,1) || $NFStrings{"currencychar"} || '$';
         }
      }
   elsif ($bracketstr eq '?$') {
      $operator = $cmd_currency;
      $operand = '[?$]';
      }
   elsif ($allowedcolors{uc $bracketstr}) {
      $operator = $cmd_color;
      $operand = $allowedcolors{uc $bracketstr};
      }
   elsif ($bracketstr =~ m/^style=([^"]*)$/) { # [style=...]
      $operator = $cmd_style;
      $operand = $1;
      }
   elsif ($bracketstr eq ",") {
      $operator = $cmd_separator;
      $operand = $bracketstr;
      }
   elsif ($alloweddates{uc $bracketstr}) {
      $operator = $cmd_date;
      $operand = $alloweddates{uc $bracketstr};
      }
   elsif ($bracketstr =~ m/^[<>=]/) { # comparison operator
      $bracketstr =~ m/^([<>=]+)(.+)$/; # split operator and value
      $operator = $cmd_comparison;
      $operand = "$1:$2";
      }
   else { # unknown bracket
      $operator = $cmd_copy;
      $operand = "[$bracketstr]";
      }

   return ($operator, $operand);

   }

