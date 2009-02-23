#!/usr/bin/perl

   use strict;
   use CGI qw(:standard);
   use URI;
   use HTTP::Daemon;
   use HTTP::Status;
   use HTTP::Response;
   use Socket;
   use CGI::Cookie;

   use SocialCalcServersideUtilities;


   my $datadir = "se2-data/";
   my $versionstr = "2.01";
   my $jsdir = "/sgi/scjstest/";

#
# This whole first section lets this code run either as a CGI script on a server
# or standalone on the desktop run from the Perl command line.
#
# The main processing starts with process_request.
#

   if ($ENV{REQUEST_METHOD}) { # being run as a CGI on a server
      print "Content-type: text/html\n\n";
      my $q = new CGI;
      print process_request($q);
      exit;
      }

   # running locally - do mini-server

   my $d = HTTP::Daemon->new (
                    LocalPort => 6557,
                    Reuse => 1);

   if (!$d) {
      print "simpleedit could not start on 127.0.0.1:6557\n";
      exit;
      }

   print "simpleedit14\nAccess at: http://127.0.0.1:6557/\n";

   while (my $c = $d->accept) {

      # Make sure the request is from our machine

      if ($c) {
         my ($port, $host) = sockaddr_in(getpeername($c));
         if ($host ne inet_aton("127.0.0.1")) {
            $c->close;  # no - ignore request completely
            undef($c);
            next;
            }
         }

      # Process the request

      while ((defined $c) && (my $r = $c->get_request)) {
         if ($r->method eq 'POST' || $r->method eq 'GET') {
            $c->force_last_request;
            my $uri = $r->uri;
            if ($uri =~ /favicon/) {   # if this is a request for favicon.ico, ignore
               $c->send_error(RC_NOT_FOUND);
               next;
               }
            if ($uri =~ /\/quit$/) {
               $c->send_file_response("quitmessage.html");
               $c->close;
               undef($c);
               exit;
               }
            if ($uri =~ /\/([a-z\-0-9]+)\.(gif|js|css)$/) { # ok request
               $uri = "$1.$2";
               $uri = "images/$uri" if $2 eq "gif";
#               if ($2 eq "js") {
#                  $res->content_type("text/html; charset=UTF-8");
#                  }
               $c->send_file_response($uri);
               next;
               }

            my $resp="";
            if ($r->method eq 'POST') {
               my $q = new CGI($r->content());
               $resp = process_request($q)
               }
            else {
               my $q = new CGI($r->uri->query());
               $resp = process_request($q)
               }
            my $res = new HTTP::Response(200);
            $res->content_type("text/html; charset=UTF-8");
            $res->expires("-1d");
            $res->content($resp);
            $c->send_response($res);
            }

         else {
            $c->send_error(RC_FORBIDDEN);
            }
         }

      $c->close;
      undef($c);
      }

#
# Main routine starts here:
#

sub process_request {

   my ($request) = @_;
   my $q = new CGI($request);

   my $response;

   my ($statusmessage);

   if ($q->param("loadsheet")) { # Ajax request: loadsheet=pagename.sheetname
      my $fullsheetname = $q->param("loadsheet");
      $fullsheetname =~ m/^(.*)\.(.*)$/;
      my $pagename = $1;
      my $sheetname = $2;
      my $page = {};
      load_page($q, $pagename, $page);
      my $sheetstr = $page->{items}{$sheetname}{text};
      my $htmlpos = index($sheetstr, "\nHTML:\n");
      if ($htmlpos >= 0) {
         $sheetstr = substr($sheetstr, 0, $htmlpos);
         }
sleep 1;
      return $sheetstr;
      }

   if ($q->param("wikitext")) { # Ajax request: wikitext=wikitext-url-encoded
      my $wikitext = $q->param("wikitext");
      $wikitext = expand_wikitext($wikitext);
sleep 1;
      return $wikitext;
      }

   my $pagename = $q->param('pagename');

   if ($q->param('newpage')) {
      $pagename = $q->param('newpagename');
      }

   $pagename =~ s/[^a-zA-Z0-9_\-]*//g;
   if (!$pagename) {
      $pagename = "home";
      $statusmessage .= "No pagename given - using '$pagename'<br>";
      }

   if ($q->param('savepageedit')) {
      my $newstr;
      $newstr = $q->param('pagetext');
      if ($q->param('edittype') eq "clean") {
         my $page = {};
         load_page($q, $pagename, $page);
         $newstr =~ s/^\[(spreadsheet|datatable|drawing)\:(.+?)\]/"[$1:$2:\n" . $page->{items}{$2}{text} . "\n:$1]"/gme;
         }
      open (PAGEFILEOUT, ">$datadir$pagename.page.txt");
      print PAGEFILEOUT $newstr;
      close PAGEFILEOUT;
      $statusmessage .= "Saved updated page '$pagename'.<br>";
      }

   if ($q->param("editpage") || $q->param("editrawpage")) { # when one of the "editpage" buttons is pressed
      return do_editpage($q, $pagename, $statusmessage);
      }

   foreach my $p ($q->param) {  # go through all the parameters
      if ($p =~ /^edit(spreadsheet|datatable):(.*)/) { # "editspreadsheet:sheetname" pressed
         return start_editsheet($pagename, $2, $q, $statusmessage);
         }
      if ($p =~ /^editdrawing:(.*)/) { # "editdrawing:sheetname" pressed
         return start_editdrawing($pagename, $1, $q, $statusmessage);
         }
      }

   if ($q->param('savespreadsheet')) { # save the edited spreadsheet
      my $page = {};
      load_page($q, $pagename, $page);
      my $pagestr = $page->{raw};
      my $sheetname = $q->param('sheetname');
      my $sheettype = $page->{items}{$sheetname}{type};

      $pagestr =~ s/^\[$sheettype\:$sheetname:.*?\:$sheettype\]/"[$sheettype:$sheetname:\n" . $q->param('newstr') . "\n:$sheettype]"/sme;

      open (PAGEFILEOUT, ">$datadir$pagename.page.txt");
      print PAGEFILEOUT $pagestr;
      close PAGEFILEOUT;
      $statusmessage =
          "Saved updated $sheettype '$sheetname' on page '$pagename'.<br>";
      }

   $response = do_displaypage($q, $pagename, $statusmessage); # Otherwise, display page

   return $response;

   }

#
# load_page($q, $pagename, \%page)
#
# Loads the specified page and puts the parts into %page:
#
#    %page{raw} - raw page text
#    %page{clean} - page text with embedded items' bodies removed
#    %page{items}{name} - hash with embedded item "name"'s info
#    %page{items}{name}{text} - text of embedded item's body
#    %page{items}{name}{type} - type embedded item (e.g., "spreadsheet")
#

sub load_page {

   my ($q, $pagename, $page) = @_;

   $page->{bodies} = {};

   open (PAGEFILEIN, "$datadir$pagename.page.txt");
   my ($rawstr, $cleanstr);
   while (my $line = <PAGEFILEIN>) {
      $line =~ s/\r//g;
      $rawstr .= $line;
      if ($line =~ m/^\[(spreadsheet|datatable|drawing)\:(.*?)\:/) {
         my $type = $1;
         my $name = $2;
         $page->{items}{$name} = {type => $type, text => ""};
         my $itemstr;
         while (my $sline = <PAGEFILEIN>) {
            $sline =~ s/\r//g;
            $rawstr .= $sline;
            last if substr($sline, 0, length($type)+2) eq ":$type]";
            $itemstr .= $sline;
            }
         $page->{items}{$name}{text} = $itemstr;
         $cleanstr .= "[$type:$name]\n";
         }
      else {
         $cleanstr .= $line;
         }
      }
   close PAGEFILEIN;
   $page->{raw} = $rawstr;
   $page->{clean} = $cleanstr;

   return;
   }

#
# $response = do_displaypage($q, $pagename, $statusmessage) - Display page
#

sub do_displaypage {

   my ($q, $pagename, $statusmessage) = @_;
   my $response;

   my $page = {};
   load_page($q, $pagename, $page);

   open (PAGEFILEIN, "$datadir$pagename.page.txt");
   my $pagestr;
   while (my $line = <PAGEFILEIN>) {
      $line =~ s/\r//g;
      if ($line =~ m/^\[(spreadsheet|datatable|drawing)\:(.*?)\:/) {
         my $sheettype = $1;
         my $sheetname = $2;
         my $sheetstr;
         while (my $sline = <PAGEFILEIN>) {
            $sline =~ s/\r//g;
            last if $sline =~ m/^:(spreadsheet|datatable|drawing)]/;
            $sheetstr .= $sline;
            }
         my $htmlpos = index($sheetstr, "\nHTML:\n");
         my $html="";
         my $localhtml="";
         if ($htmlpos >= 0) {
            $html = substr($sheetstr, $htmlpos+7);
            $sheetstr = substr($sheetstr, 0, $htmlpos);
            my $sheet = CreateSheet();
            my $parts = DecodeSpreadsheetSave($sheetstr);
            ParseSheetSave($sheet, $parts->{sheet});
            my $context = CreateRenderContext($sheet);
            $localhtml = RenderSheet($context);
            }
         else {
            $html = "Sheet goes here";
            $localhtml = "Sheet goes here";
            }
         $sheetstr = special_chars($sheetstr);
         $pagestr .= <<"EOF";
<textarea id="sheetdata-$sheetname" style="display:none;">$sheetstr</textarea>
<table cellspacing="0" cellpadding="0" width="100%">
<tr onmouseout="hide_buttons(!document.getElementById('showb').checked);"
        onmouseover="hide_buttons(false);"><td valign="top" id="sheet-$sheetname">$localhtml</td>
    <td class="hideable" width="5" style="visibility:hidden;border-top:1px solid #80A9F3;border-bottom:1px solid #80A9F3;padding:6px;">&nbsp;</td>
    <td class="hideable" width="100" valign="middle" style="visibility:hidden;border-left:1px solid #80A9F3;padding:6px;">
  <span style="font-size:smaller;">$sheettype:<br>$sheetname<br></span>
  <input type="submit" name="edit$sheettype:$sheetname" value="Edit" style="font-size:smaller;">
</td></tr></table>
<script>
if("$sheettype"!="spreadsheet") show_$sheettype("$sheetname");
</script>
EOF
         }
      else {
         $pagestr .= expand_wikitext($line);
         }
      }

   close PAGEFILEIN;

   $response = <<"EOF";
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"
 "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Simple Page Editor With Spreadsheets $versionstr</title>
<script type="text/javascript" src="${jsdir}socialcalcconstants.js"></script>
<script type="text/javascript" src="${jsdir}socialcalc-3.js"></script>
<script type="text/javascript" src="${jsdir}socialcalctableeditor.js"></script>
<script type="text/javascript" src="${jsdir}formatnumber2.js"></script>
<script type="text/javascript" src="${jsdir}formula1.js"></script>
<script type="text/javascript" src="${jsdir}drawlib3.js"></script>
<script>

function show_datatable(sheetname) {
   var sheet = new SocialCalc.Sheet();
   sheet.ParseSheetSave(document.getElementById("sheetdata-"+sheetname).value);
   var context=new SocialCalc.RenderContext(sheet);
   context.CalculateCellSkipData();
   context.CalculateColWidthData();
   context.showGrid=true;
   context.showRCHeaders=true;
   var teobj=new SocialCalc.TableEditor(context);
   teobj.imageprefix="/sgi/scjstest/sc";
   teobj.recalcFunction = null;
   var sizes = SocialCalc.GetViewportInfo();
   var teelement=teobj.CreateTableEditor(800, 300);
   var teid=document.getElementById("sheet-"+sheetname);
   teid.replaceChild(teelement, teid.firstChild);
   teobj.SchedulePositionCalculations();
   }

function show_drawing(sheetname) {
   var teid=document.getElementById("sheet-"+sheetname);
   var sheet = new SocialDraw.Sheet(teid, false, document.getElementById("sheetdata-"+sheetname).value);
   sheet.Display();
   }

function hide_buttons(hide) {
   var val=hide ? 'hidden':'visible';
   var tds = document.getElementsByTagName("td");
   for (var i=0; i<tds.length; i++) {
      if (tds[i].className=="hideable") tds[i].style.visibility=val;
      }
   }
</script>
<style>
body, td, input, texarea
 {font-family:verdana,helvetica,sans-serif;font-size:small;}
</style>
</head>
<body>
<form action="" method="POST">
<div style="padding:6px;background-color:#80A9F3;">
<div style="font-weight:bold;color:white;">SIMPLE SYSTEM FOR EDITING PAGES WITH SPREADSHEETS AND MORE</div>
<div style="color:#FDD;font-weight:bold;">$statusmessage &nbsp;</div>
<table cellspacing="0" cellpadding="0" width="100%">
<tr>
<td>
 Viewing page: <span style="font-style:italic;font-weight:bold;">$pagename</span>
 &nbsp;<input type="submit" name="editpage" value="Edit This Page" style="font-size:smaller;">
</td>
<td align="right">
 <input type="submit" name="editrawpage" value="Edit Raw Page" style="font-size:smaller;">
</td></tr></table>
</div>
<div style="border:6px solid #80A9F3;padding:0px 10px 10px 10px;">
<div style="padding-top:5px;">
 <table cellpadding="0" cellspacing="0"><tr><td width="100%"></td><td valign="top">
  <input id="showb" type="checkbox" value="1" onclick="hide_buttons(!this.checked);">
 </td><td style="font-size:smaller;padding-left:4px;">Show&nbsp;item<br>edit&nbsp;buttons</td>
 </tr></table>
</div>
$pagestr
</div>
<div style="padding:6px;background-color:#80A9F3;">
<div style="font-weight:bold;font-size:smaller;">Pages:</div>
<div style="font-size:smaller;">
EOF

   my @pagefiles = glob("$datadir*.page.txt"); # Get list of all pages
   for (my $pnum=0; $pnum <= $#pagefiles; $pnum++) {
      $pagefiles[$pnum] =~ m/^(?:.*\/){0,1}(.*).page.txt$/;
      $response .= ", " if $pnum!=0;
      $response .= qq!<a href="?pagename=$1">$1</a>!;
      }

   $response .= <<"EOF";
</div>
<br>
<input type="hidden" name="pagename" value="$pagename">
</form>
<form action="" method="POST">
<input type="text" name="newpagename" value="">
<input type="submit" name="newpage" value="Create New Page">
</form>
</div>
<br>
</body>
</html>
EOF

   return $response;
   }


#
# do_editpage($q, $pagename, $statusmessage) - Show page editing display
#

sub do_editpage {

   my ($q, $pagename, $statusmessage) = @_;
   my $response;

   my $page = {};
   load_page($q, $pagename, $page);

   my $edittype = $q->param("editpage") ? "clean" : "raw";

   my $pagestr = special_chars($page->{$edittype});

   $response = <<"EOF";
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"
 "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Simple Page Editor With Spreadsheets $versionstr</title>
<style>
body, td, input, texarea
 {font-family:verdana,helvetica,sans-serif;font-size:small;}
</style>
</head>
<body>
<form name="f0" action="" method="POST">
<div style="padding:6px;background-color:#80A9F3;">
<div style="font-weight:bold;color:white;">SIMPLE SYSTEM FOR EDITING PAGES WITH SPREADSHEETS AND MORE</div>
<div style="color:#FDD;font-weight:bold;">$statusmessage &nbsp;</div>
<div style="margin-bottom:6px;">Editing page: <span style="font-style:italic;font-weight:bold;">$pagename</span></div>
<script>
function addss(sstype) {
 var now = new Date();
 var name = ""+now.getFullYear()+(now.getMonth()+101).toString().substr(1)+(now.getDate()+100).toString().substr(1)+
        (now.getHours()+100).toString().substr(1)+(now.getMinutes()+100).toString().substr(1)+(now.getSeconds()+100).toString().substr(1);
 var sname =
  prompt("New "+sstype+" name (alphanumeric only, unique on page):", name);
 if (!sname) return false;
 sname = sname.replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
 if (!sname) return false;
 document.f0.pagetext.value =
  document.f0.pagetext.value + "\\n["+sstype+":" + sname +"]\\n";
 return false;
}
</script>
<div style="margin-bottom:4px;">
<table cellspacing="0" cellpadding="0" width="100%">
<tr>
<td>
 <input type="submit" name="savepageedit" value="Save">
 <input type="hidden" name="edittype" value="$edittype">
 <input type="submit" name="canceledit" value="Cancel">
</td>
<td align="right">
EOF

   if ($edittype eq "clean") {
      $response .= <<"EOF";
<input type="submit" value="Add Spreadsheet" onclick="return addss('spreadsheet');" style="font-size:smaller;">
<input type="submit" value="Add Data Table" onclick="return addss('datatable');" style="font-size:smaller;">
<input type="submit" value="Add Character Drawing" onclick="return addss('drawing');" style="font-size:smaller;">
EOF
      }

   $response .= <<"EOF";
</td></tr></table>
</div>
<textarea name="pagetext" rows="20" style="width:100%;">$pagestr</textarea>
<input type="hidden" name="pagename" value="$pagename">
</div>
</form>
</body>
</html>
EOF

   return $response;
   }

#
# start_editsheet($pagename, $sheetname, $q, $statusmessage)
#    - render initial editing display
#

sub start_editsheet {

   my ($pagename, $sheetname, $q, $statusmessage) = @_;

   my ($response, $sheetstr);

   my $page = {};
   load_page($q, $pagename, $page);

   $sheetstr = $page->{items}{$sheetname}{text};

   my $htmlpos = index($sheetstr, "\nHTML:\n");
   if ($htmlpos >= 0) {
      $sheetstr = substr($sheetstr, 0, $htmlpos);
      }

   $response = <<"EOF"; # output page with edit JS code
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Simple Page Editor With Spreadsheets $versionstr</title>
<script type="text/javascript" src="${jsdir}socialcalcconstants.js"></script>
<script type="text/javascript" src="${jsdir}socialcalc-3.js"></script>
<script type="text/javascript" src="${jsdir}socialcalctableeditor.js"></script>
<script type="text/javascript" src="${jsdir}formatnumber2.js"></script>
<script type="text/javascript" src="${jsdir}formula1.js"></script>
<script type="text/javascript" src="${jsdir}socialcalcspreadsheetcontrol.js"></script>
<style>
body, td, input, texarea
 {font-family:verdana,helvetica,sans-serif;font-size:small;}
</style>
</head>
<body>
<form name="f0" action="" method="POST">
<div style="padding:6px;background-color:#80A9F3;">
<div style="font-weight:bold;color:white;">SIMPLE SYSTEM FOR EDITING PAGES WITH SPREADSHEETS AND MORE</div>
<div style="color:#FDD;font-weight:bold;">$statusmessage &nbsp;</div>
<div style="margin-bottom:6px;">Editing page: <span style="font-style:italic;font-weight:bold;">$pagename</span></div>
<input type="submit" name="savespreadsheet" value="Save" onclick="dosave();">
<input type="submit" name="cancelspreadsheet" value="Cancel">
<textarea name="savestr" id="sheetdata" style="display:none;">$sheetstr</textarea>
<input type="hidden" name="newstr" id="newdata" value="">
<input type="hidden" name="pagename" value="$pagename">
<input type="hidden" name="sheetname" value="$sheetname">
</div>
</form>
<div id="msg" style="position:absolute;right:15px;">
<input type="button" style="font-size:x-small;" value="Clear" onclick="addmsg('',true);"><br>
 <input type="button" style="font-size:x-small;" value="ShowCols" onclick="showvalue();"><br>
 <textarea id="msgtext" style="margin-top:10px;width:140px;height:500px;"></textarea>
</div>
<div id="tableeditor" style="margin:8px 170px 10px 0px;">editor goes here</div>
<script>

document.getElementById("msgtext").value = "";

function setmsg(msg) {document.getElementById("msg").innerHTML = msg;}
function addmsg(msg, clear) {
   var msgtextid = document.getElementById("msgtext");
   if (!msgtextid) {
      document.getElementById("msg").innerHTML = '<textarea id="msgtext" cols="60" rows="5"></textarea>';
      msgtextid = document.getElementById("msgtext");
      }
   if (clear) msgtextid.value = "";
   if (msgtextid.value.length>0) msgtextid.value += "\\n";
   msgtextid.value += msg;
   }
var heights=[];
function showvalue() {
   addmsg(spreadsheet.editor.colwidth, true);
   }


// define functions

function loadsheet(sheetname) {
addmsg("loadsheet:"+sheetname);
   return ajaxrequest("", "&loadsheet="+encodeURIComponent(sheetname));
   }

SocialCalc.RecalcInfo.LoadSheet = loadsheet;

// function 

var http_request;

function ajaxrequest(url, contents) {

   http_request = null;

   if (window.XMLHttpRequest) { // Mozilla, Safari,...
      http_request = new XMLHttpRequest();
      }
   else if (window.ActiveXObject) { // IE
      try {
         http_request = new ActiveXObject("Msxml2.XMLHTTP");
         }
      catch (e) {
         try {
            http_request = new ActiveXObject("Microsoft.XMLHTTP");
            }
         catch (e) {}
         }
      }
   if (!http_request) {
      alert('Giving up :( Cannot create an XMLHTTP instance');
      return false;
      }

   // Make the actual request

   http_request.onreadystatechange = alertContents;
   http_request.open('POST', document.URL, true); // async
   http_request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
   http_request.send(contents);

   return true;

   }

function alertContents() {

   var loadedstr = "";

   if (http_request.readyState == 4) {
addmsg("received:"+http_request.responseText.length+" chars");
      try {
         if (http_request.status == 200) {
            loadedstr = http_request.responseText || "";
            http_request = null;
            }
         else {
            ;
            }
         }
      catch (e) {
         }
      parts = spreadsheet.DecodeSpreadsheetSave(loadedstr);
      if (parts) {
         if (parts.sheet) {
            SocialCalc.RecalcLoadedSheet(null, loadedstr.substring(parts.sheet.start, parts.sheet.end), true); // notify recalc loop
            }
         }
      if (loadedstr=="") {
         SocialCalc.RecalcLoadedSheet(null, "", true); // notify recalc loop that it's not available, but that we tried
         }
      }
   }


function dosave() {
   var sheetstr = spreadsheet.CreateSpreadsheetSave();
   var htmlstr = spreadsheet.CreateSheetHTML();
   document.getElementById("newdata").value = sheetstr + "\\nHTML:\\n" + htmlstr;
   }

var http_request2;

function ajaxrequest2(url, contents) {

   http_request2 = null;

   if (window.XMLHttpRequest) { // Mozilla, Safari,...
      http_request2 = new XMLHttpRequest();
      }
   else if (window.ActiveXObject) { // IE
      try {
         http_request2 = new ActiveXObject("Msxml2.XMLHTTP");
         }
      catch (e) {
         try {
            http_request2 = new ActiveXObject("Microsoft.XMLHTTP");
            }
         catch (e) {}
         }
      }
   if (!http_request2) {
      alert('Giving up :( Cannot create an XMLHTTP instance');
      return false;
      }

   // Make the actual request

//   http_request2.onreadystatechange = alertContents2;
   http_request2.open('POST', document.URL, false); // sync
   http_request2.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
   http_request2.send(contents);

   return http_request2.responseText || "";

   }

function alertContents2() {

   var loadedstr = "";
   var parts;

   if (http_request2.readyState == 4) {
      try {
         if (http_request.status == 200) {
            loadedstr = http_request2.responseText || "";
            http_request2 = null;
            }
         else {
            ;
            }
         }
      catch (e) {
         }
      parts = spreadsheet.DecodeSpreadsheetSave(loadedstr);
      if (parts) {
         if (parts.sheet) {
            SocialCalc.RecalcLoadedSheet(null, loadedstr.substring(parts.sheet.start, parts.sheet.end), true); // notify recalc loop
            }
         }

      }
   }

// start everything

   var spreadsheet = new SocialCalc.SpreadsheetControl();
   spreadsheet.imagePrefix="/sgi/scjstest/images/";
   spreadsheet.editor.imageprefix="/sgi/scjstest/images/sc";
   spreadsheet.InitializeSpreadsheetControl("tableeditor", 0, 0, 0);

   var savestr = document.getElementById("sheetdata").value;
   var parts = spreadsheet.DecodeSpreadsheetSave(savestr);
   if (parts) {
      if (parts.sheet) {
         spreadsheet.sheet.ResetSheet();
         spreadsheet.ParseSheetSave(savestr.substring(parts.sheet.start, parts.sheet.end));
         }
      if (parts.edit) {
         spreadsheet.editor.LoadEditorSettings(savestr.substring(parts.edit.start, parts.edit.end));
         }
      }
   if (spreadsheet.editor.context.sheetobj.attribs.recalc=="off") {
      spreadsheet.ExecuteCommand('redisplay', '');
      }
   else {
      spreadsheet.ExecuteCommand('recalc', '');
      }

</script>
</body>
</html>
EOF

   return $response;

   }


#
# start_editdrawing($pagename, $sheetname, $q, $statusmessage)
#    - render initial editing display
#

sub start_editdrawing {

   my ($pagename, $sheetname, $q, $statusmessage) = @_;

   my ($response, $sheetstr);

   my $page = {};
   load_page($q, $pagename, $page);

   $sheetstr = $page->{items}{$sheetname}{text};

   $response = <<"EOF"; # output page with edit JS code
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"
 "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Simple Page Editor With Spreadsheets $versionstr</title>
<script type="text/javascript" src="${jsdir}socialcalc-3.js"></script>
<script type="text/javascript" src="${jsdir}drawlib3.js"></script>
<style>
body, td, input, texarea
 {font-family:verdana,helvetica,sans-serif;font-size:small;}
</style>
</head>
<body>
<form name="f0" action="" method="POST">
<div style="padding:6px;background-color:#80A9F3;">
<div style="font-weight:bold;color:white;">SIMPLE SYSTEM FOR EDITING PAGES WITH SPREADSHEETS AND MORE</div>
<div style="color:#FDD;font-weight:bold;">$statusmessage &nbsp;</div>
<div style="margin-bottom:6px;">Editing page: <span style="font-style:italic;font-weight:bold;">$pagename</span></div>
<input type="submit" name="savespreadsheet" value="Save" onclick="dosave();">
<input type="submit" name="cancelspreadsheet" value="Cancel">
<textarea name="savestr" id="sheetdata" style="display:none;">$sheetstr</textarea>
<input type="hidden" name="newstr" id="newdata" value="">
<input type="hidden" name="pagename" value="$pagename">
<input type="hidden" name="sheetname" value="$sheetname">
</div>
</form>
<div id="drawingeditor" style="margin:8px 0px 10px 0px;"></div>
<script>
   var de = document.getElementById("drawingeditor");
   var sheet = new SocialDraw.Sheet(de, true, document.getElementById("sheetdata").value);
   sheet.Display();

function dosave() {
   document.getElementById("newdata").value = sheet.Save();
   }

</script>
</body>
</html>
EOF

   return $response;

   }


# # # # # # # # # #
# special_chars($string)
#
# Returns $estring where &, <, >, " are HTML escaped
# 

sub special_chars {
   my $string = shift @_;

   $string =~ s/&/&amp;/g;
   $string =~ s/</&lt;/g;
   $string =~ s/>/&gt;/g;
   $string =~ s/"/&quot;/g;

   return $string;
}


#
# decode_from_ajax($string) - Returns a string with 
#       \n, \b, and \c escaped to \n, \, and :
#

sub decode_from_ajax {
   my $string = shift @_;

   $string =~ s/\\n/\n/g;
   $string =~ s/\\c/:/g;
   $string =~ s/\\b/\\/g;

   return $string;
}


#
# encode_for_ajax($string) - Returns a string with 
#       \n, \, :, and ]]> escaped to \n, \b, \c, and \e
#

sub encode_for_ajax {
   my $string = shift @_;

   $string =~ s/\\/\\b/g;
   $string =~ s/\n/\\n/g;
   $string =~ s/\r//g;
   $string =~ s/:/\\c/g;
   $string =~ s/]]>/\\e/g;

   return $string;
}


#
# expand_wikitext($string) - Returns $string doing wiki-style formatting
# 

sub expand_wikitext {
   my $string = shift @_;

   # Process forms that use URL encoding first

   # [page:pagename text] to link to other pages on this site
   $string =~ s!\[page:(.+?)\s+(.+?)?]!'{{lt}}a href={{quot}}?pagename=' . 
            $1 . "{{quot}}{{gt}}$2\{{lt}}/a{{gt}}"!xegs;

   # [url:url text] to link to other pages on other sites
   $string =~ s!\[url:(.+?)\s+(.+?)?]!'{{lt}}a href={{quot}}' . 
            $1 . "{{quot}} target={{quot}}_blank{{quot}}{{gt}}$2\{{lt}}/a{{gt}}"!xegs;

   # Convert &, <, >, "

   $string = special_chars($string);

   $string =~ s/^\= (.*) \=$/<span style="font-size:150%;font-weight:bold;">$1<\/span>/gs;
   $string =~ s/\n/<br>/g;  # Line breaks are preserved
   $string =~ s/('*)'''(.*?)'''/$1<b>$2<\/b>/gs; # Wiki-style bold/italics
   $string =~ s/''(.*?)''/<i>$1<\/i>/gs;
   $string =~ s/\{\{amp}}/&/gs; # {{amp}} for ampersand
   $string =~ s/\{\{lt}}/</gs; # {{lt}} for less than
   $string =~ s/\{\{gt}}/>/gs; # {{gt}} for greater than
   $string =~ s/\{\{quot}}/"/gs; # {{quot}} for quote
   $string =~ s/\{\{lbracket}}/[/gs; # {{lbracket}} for left bracket
   $string =~ s/\{\{rbracket}}/]/gs; # {{rbracket}} for right bracket
   $string =~ s/\{\{lbrace}}/{/gs; # {{lbrace}} for brace

   return $string;
}

