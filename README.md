# Create PPT Presentation From Excel

## This Is An Example of Using a VBA Macro for Excel to Create a PTT
In this repo you will find VB code that is used in Excel to generate a routine PTT
report from a PTT Template file based on data that is refreshed in the Excel Workbook.
Using VB or a Macro to fulfill this function saves hours of formatting time on a 19 slide
presentation that is generated monthly. An example of the first few slides generated is
also included as a .JPG file.

The reason this was needed is because the data behind the charts is much too large to
allow the final file to be emailed (or even opened efficiently on your machine) if you
just link the charts directly. All charts must be inserted into the PTT, resized as charts
to prevent stretching and pixilation, and then pasted as images to reduce the PTT file
size.

In the code you will see the first four slides are created uniquely, and then slides 5-19
are generated through loops, with the exception of slide 15, which is a section break
slide. You will also notice wait time built in to allow the machine used for this report
to fully load the large charts to the clipboard, as well as the "Save As" dialog box hard
coded to open at the end, so the user (most often myself) does not accidentally save over
the PTT template.