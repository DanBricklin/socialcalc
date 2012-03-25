This is a fork of Dan Bricklin's SocialCalc (http://github.com/DanBricklin/socialcalc), 
with the following modifications made for the Drupal Sheetnode module (http://drupal.org/project/sheetnode):

2012-03-25
* Change positioning strategy to use position:relative parents
* Use DOM element.getBoundingClientRect() to calculate element offset

2011-07-04
* Support multiple sheets per page.

2011-02-18
* Support auto-detecting HTML text in default input mode.

2011-01-02
* Support test formatting using non-text value formats (like Excel and OOoCalc).

2010-11-30
* Support hiding rows and columns.

2010-10-30
* Fix parsing expressions with unary minus.

2010-08-24
* Support max rows and columns. 

2010-07-10
* Merge commit 'upstream/master'. 

2010-06-19 Initial fork with one year's worth of cumulative patches
* Relativize absolute-positioned elements (e.g., scrollbars, input echo, cell handles, etc.) to allow hosting the spreadsheet inside arbitrary DIVs.
* Reposition popups to screen center taking scrolling offset into consideration.
* Display cell comments as HTML "title" attribute.
* Support read-only cells.

