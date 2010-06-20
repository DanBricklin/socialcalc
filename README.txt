This is a fork of Dan Bricklin's SocialCalc (http://github.com/DanBricklin/socialcalc), 
with the following modifications made for the Drupal Sheetnode module (http://drupal.org/project/sheetnode):

2010-06-19 Initial fork with one year's worth of cumulative patches
* Relativize absolute-positioned elements (e.g., scrollbars, input echo, cell handles, etc.) to allow hosting the spreadsheet inside arbitrary DIVs.
* Reposition popups to screen center taking scrolling offset into consideration.
* Display cell comments as HTML "title" attribute.
* Support read-only cells.

