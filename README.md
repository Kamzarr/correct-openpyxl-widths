# correct-openpyxl-widths

It solves two problems occurring in normal openpyxl column width reading:
    1. When file have column dimensions written in ranges - normal openpyxl width reading returns correct
    width value only for first column in range for next ones in range it returns default width
    so there is no way of knowing if column really have default width or have custom width but is just written in range
    
    2. For some reason reading column width using regular openpyxl function (here using default_width_source = 2)
    it gives value of 13 for default width
	
example of problematic column dimensions are in "dimensions_test.xlsx", columns J-L (10-12):
<col min="10" max="12" width="15.77734375" customWidth="1"/>

openpyxl regular dimensions reading will give correct value only to first (J) column in range
