# Welcome to the ChargeXL!

An XLL & Excel-DNA based add-in for MS Excel which is open for further development and enhancement in the community as well as available to download for MS Excel (2007,2010,2013, and 2016) users. It might cost me and other community developers the huge amount of time but it's free to use for everyone. 

***

### Set-Up (Developers)
1. This XLL add-in depends upon Excel-DNA (64 bit). So, please go through their documentation on setting up the environment (32bit or 64bit). Otherwise, Excel will claim that the file (_ChargeXL.xll_) is in an incorrect format.
2. Clone the git repository and open the file _ChargeXL.sln_ in Visual Studio.
3. In the project properties for the library, configure the debugging start action as follows: (‘Office16’ is the directory for Office 2016, substitute ‘Office15’ for Office 2013, 'Office14' for Office 2010, 'Office12' for Office 2007 or another version as appropriate.)


### Set-Up (Users)
1. Click here to download the latest version (beta) of ChargeXL
2. Open the ChargeXL.xll file which is an Excel Add-in. A new tab will appear in your MS Excel and you would be able to use all the below features and functions in your Excel workbook.

### Features

* Text Utils
    * Insert Text
    * Reverse Text
    * Sort value with in cell
    * Change Case
        * Upper Case
        * Lower Case
        * Proper Case
        * Sentence Case
        * Toggle Case
        * Snake Case
    * Trim Spaces
        * From Left/Right/Both Sides
        * From Both Sides
        * Excessive Spaces
    * Delete Text
        * 'n' number of characters from Left/Right/Both Sides/a position in text
        * Ending CR or LF
        * Numbers
        * Alphabets
    * Super/Sub script
	* Accented characters
	* Extract
        * Valid Email-Address
        * Numbers
        * Alphabets
        * Dates
        * Alpha-Numeric [A-Z][0-9]
   
* Number Utils
	* Convert Text Values to Numbers
	* Convert Numbers to Text Format
	* Change the sign (+ or -)
	* Spell Number
	* Apply Calculation
	* Convert % to numbers and vice-versa
	* Round Numbers (Actual conversion)
	* Fill with leading zeros
	* Convert to displayed value
	* Quick numbering
	* Select Min/Max values
	

* Range Utils
	* Copy multiple selection
	* Advanced sorting
	* Color every n'th row/coloumn
	* Merge Row/Column data
	* Delete
        * Empty Rows/Columns 
        * Odd Rows/Columns
	* Insert
        * Empty Rows/Columns 
        * Odd Rows/Columns
	* Advanced Find/Replace
	* Flip the values 
	* Convert to static formatting
	* Highlight and count duplicates
	* Conditional Hide
	* Range Names
        * Delete in selection
        * Delete in worksheet
        * Delete in workbook
        * Delete with invalid cell reference
        * Create List

* Select Utils
	* Inverse selection
	* Entire Row/Column
	* Active cell  value
	* Based on Type
        * Empty/Non-Empty
        * Protected/Non-protected
        * Numeric/Non-numeric
        * Hidden
        * Formula
            * Precedents/Dependents
        * Error
        * Data Validation	
        * Volatile formulas	
	* Based on Value
        * Odd/Even Numbers
        * Duplicate Values
        * Unique Values
        * Hyperlinks
        * Commented cells
        * 
	* Based on Format
        * Font / background Color
        * Bold/Italic/Underlined cells
        * Text Length
	* Row / Column differences
	* Merge Selection
	* Adjacents cells
        * with same value
        * with same formatting		
	* All Sheets
	* Give count and select all Objects
	* Expand current selection to other worksheets
    * Extend select to last/first row/column

* Formula Utils
    * Change formula to its calculated value
    * Rebuild array formulae
    * Insert/Remove apostrophe 
    * Replace range name with their formula
    * Change reference style (Absolute to Reference)
    * Custom formula error

* Date Utils

* Sheet Utils

* Web Utils

* Information Utils

* System Utils
