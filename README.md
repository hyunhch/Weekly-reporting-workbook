Version 1.2.9

	Fixed bug caused by not unprotecting the Roster Page when creating, loading, or adding students to an activity

Version 1.2.8	
	
	The Add New Activity form can now have multiple lines and scroll vertically in the Decription box
	
	Text on Roster Page now copy into the Add New Activity form, the Label and the Description boxes
	
	The Submission Information form now indicates that the date should be in the format mm/dd/yyyy 
	
Version 1.2.7	
	
	Fixed crash when tabulating the roster on MacOS
	
	Fixed bug with tabulatuing gender and race on MacOS
	
Version 1.2.6	
	
	Corrected Select All button checking hidden students
	
	Fixed bug using Select All button on Report Page
	
	Corrected filtering. Only visible students will be added when creating a new activity or adding to an existing activity
	
Version 1.2.5	
	
	Fixed a bug when students would not pull into a newly created activity
	
	Fixed a bug where an unsaved activity sheet wouldn't delete
	
Version 1.2.4	
	
	Selecting a program is done with a button instead of a popup
	
	Saving a copy or exporting students now includes a roster 
	
	Labels not limited to 31 characters and forbid certain characters
	
	Fixed bug when trying to create a new activity without selecting a practice
	
	Activities with no students marked present or absent are culled
	
	The select all button works when filtering rows and only applies to visible rows
	
	Improved the Import function, though it's still not very robust
	
	Fixed bug when a practice or category could not be found on a reference table
	
	Some flexibility with column order
	
Version 1.2.3	
	
	Fixed a bug when loading activities with either no students marked present or no students marked absent
	
	Fixed bug when tabulating low income or first generation students
	
Version 1.2.2	
	
	Fixed crash for Mac users
	
Version 1.2.1	
	
	Fixed bug not attaching all sheets correctly
	
Version 1.2.1	
	
	Fixed bug when creating the first activity
	
Version 1.2.0	
	
	Overhaul of workbook, most of the code rewritten
	
	Workbooks of the three programs merged into a single file that prompts for which program it will be used for
	
	Dramatic increase in efficiency, especially with deleting students
	
	Most tables are now programmatic, makign altering or expanding categories easier
	
	Added non-binary gender option
	
	Added first-generation and low-income categories for Transfer Prep and MESA University
	
	Added grades down to 1st for College Prep
	
	Custom fields on the Roster Page will now be included in exported attendance reports
	
	The roster and report will be included when saving from the Cover Page
	
	Activities with no students can be retained as placeholders. Students may be added to them and they can be dispalyed on the Report Page
	
	The full activity description is now exported to SharePoint and when saving on the Cover Page
	
	Activity category is also exported
	
	Tabbing through userforms now follows an intuitive order
	
	Real-time filters added to all userforms that display activities
	
	The Report Page now has a table that can be sorted and filtered
	
	Exported attendence now brings up a save prompt and makes a file name. Detailed attendence always exported along with simple
	
	Name, center, and date are prompted by a userform when opening the file for the first time
	
	Many error messages were removed for button clicks when something hasn't been filled out. Now the buttons simply do nothing
	
	Adjusted the size and position of buttons
	
	The worksheet version now generates automatically by referencing the file name
	
	Student attendance is pulled into an activity sheet after students are added or removed
	
	Reference sheets for low-income and first-generation added on the Cover Page for Transer Prep and MESA University
	
	Parsing the roster will display the number of duplicate students removed, if any
	
	Adding students to an activity and removing missing students only displays the number of students. This was to accommodate very large lists of students
	
	Activities can be retabulated without removing them first
	
	When adding students, the list will automatically sort by present and absent on the activity sheet
	
	Fixed errors when starting to save a copy or import a file on the Cover Page but canceling
	
	Saving from the Cover Page no longer opens up the exported file
	
	Trying to load an open activity sheet now activates that sheet
	
	Many small changes to the behavior of operations 