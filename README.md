# SRM-Inventory
Collection of Excel spreadsheets, VBA code, and scripts to help track inventory at my job



The initial purpose of the spreadsheet was to keep track of the SRMs (standard reference materials) that the lab uses.
SRMs are fairly expensive, so we needed a better way to see how much we were using and when we would need to order more.
If we run out, our production in the lab essentially stops.

The spreadsheet was just a list in the beginning, but that was hard to read. Since then, I have made a number of changes in an attempt to make reading a bit easier and make our usage of SRMs more efficient.

### FEATURES

##### ALREADY IMPLEMENTED

* [x] Use excel sorting techniques to make searching easier
* [x] Use excel filter techniques to make serching easier
* [x] Separate SRMs into four categories (good, warning, danger, out of stock)
	* [x] Use conditional formatting to color code lines
	* [x] Use formula to trigger conditional formatting based off stock condition, pressure, and expiration
* [x] Button to sort list when conditions have changed
* [x] Code to automatically populate another sheet with inventory information (sheet is emailed as part of a report)

---
##### GENERAL

* [x] Script to open and update spreadsheet daily
* [ ] Script to run email modules periodically
* [x] Clean up table
	* [x] convert from conditional format to code
	* [x] Delete columns 1 - 3
		* [x] Edit code to reflect correct columns

---
##### EMAIL

* [x] Form to submit emails
* [x] Save emails to a file
* [ ] Send email to list of people
	* [ ] Include certain range
		* [ ] Write lenMessage function 
	* [ ] when SRM or VSL changes state
		* [ ] To Yellow
		* [ ] To red
  
---
##### OTHER FORMS

* [ ] Request GMIS
* [ ] List of SRM/VSL needed
* [ ] Add standard entry

---
##### COA GENERATION

* [ ] Dropdown list to select template
* [ ] Select banner from list

-- or --

* [ ] Form with fields that will load a partially filled template

---
##### LABEL GENERATION
