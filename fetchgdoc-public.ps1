#! /opt/Microsoft/PowerShell/6/pwsh

# fetchgdoc.ps1
# Learning script for fetching gdoc as odt documents
<#

POWERSHELL LESSONS LEARNED
--------------------------
This was my first project that I attempted with PowerShell. I found that having this project helped to motivate me to learn PowerShell syntax and cmdlets that I could use to complete this project. Below are the lessons I have learned from creating this PowerShell script, and I have included some questions that I still have which are areas where I need to do some more learning. This script was successful in accomplishing all three of the goals I had for this project which were:
1. Find all of the .gdoc files in my Documents directory.
2. Export all of the .gdoc files from my GDrive account so that I had all of their content in local files on my local machine formatted in the LibreOffice format. This would enable me to open any of these files even when I'm offline.
3. Not only export the files, but save the exported files with the LibreOffice extension .odt and save them in the same directory where the original .gdoc file was found.

EXPLANATION OF CODE
-------------------
1. How to get the "doc_id" from a .gdoc file?
.gdoc files are saved to your hard drive when using the Win10 Google Drive syncing program. I have copies of these .gdoc files on my Linux computer from backups that I made of my document tree on my Win10 computer. I'm trying to get the exported files onto my Linux computer. These files are json formatted files, the actual document content is not in the .gdoc file. The following code pulls the "doc_id" from the specified .gdoc json file. I needed the Google doc_id for the PowerShell module PSGSuite which I decided to use to Export the files from my GDrive account.

Get-Content /home/mowest/Documents/somefileonGdrive.gdoc | ConvertFrom-Json | Select-Object doc_id

2. How do I get an array listing all of the .gdoc files in my Documents folder?

Get-ChildItem -Path /home/mowest/Documents/*.gdoc -Recurse

3. How do I get just the information I need from each .gdoc file in a multi column array that I can iterate through to export each document?

- First, I need to get an array with all of the documents and their properties.

$GdocArray = Get-ChildItem -Path /home/mowest/Documents/*.gdoc -Recurse

- Second, I need to extract just the "doc_id" from the .gdoc json files.
    - Step one was to iterate through the file list converting each file from json using the "ConvertFrom-Json" cmdlet.
    - Step two was to use the cmdlet "Select-Object" with the "-ExpandProperty" flag to grab just the "doc_id". These were saved into the "$GdocDoc_idArray"

$GdocDoc_idArray = ForEach-Object {Get-Content $GdocArray | ConvertFrom-Json | Select-Object -ExpandProperty doc_id}

- Third, I needed to extract the full file path so that I could place the exported file into the proper directory where the corresponding .gdoc files were found.
 
$GdocDirectory = ForEach-Object {$GdocArray | Select-Object -ExpandProperty FullName}

- Originally, I did not realize how easy it was to get the full file name path so I used this command to just grab the file names. This works well if you simply want to export the files into a new directory instead of into their origin directory. I later replaced this code below in favor of the one above.

$GdocFileNameArray = Get-ChildItem -Path /home/mowest/Documents/*/ -Name *.gdoc -Recurse

- Forth, finally I needed to combine these two lists of information from the same file set into one multi column array that can be used for the PSGSuite file exporting command. I put this code together from this Stackoverflow question: https://stackoverflow.com/questions/23411202/powershell-combine-single-arrays-into-columns

    - I will be completely honest and say this code did what I needed it to do, but I don't really understand what it is doing. I need to learn more PowerShell syntax in order to fully understand everything that is happening in these 6 lines of code.
    - This is an iterative counter used below in the New-Object cmdlet. My question about it is, do I need this iterative counter? As the New-Object cmdlet is creating my new two column array does this iterative counter tell New-Object to go to the next row of the matrix? As in, "Ok, New-Object you created column 1 called "id" and you created column 2 called "name" and you put the first doc_id from $GdocDoc_idArray into column 1 "id" and you have put the first filename from $GdocDirectory into column 2 "name" now go to the second item in each of those lists for row 2."

$i = 0

    - $GdocIDandName is the new variable name for this two column array.
    - ($GdocDoc_idArray,$GdocDirectory)[1] My question: Why is it necessary to feed these two lists into the factory pipeline at this point since they are called later by the New-Object cmdlet, especially since "[1]" indicates that only $GdocDirectory should be fed into the pipeline because it is the 2nd piece of the collection? Wei-Yen Tan made two points that give me a little more clarity about this bit of code. He said that "foreach" needs something piped into it in order to fuction. He also indicated that [1] directs PowerShell to feed only the second part of the collection in ( ) into the pipeline. From what I understand from the comments in the Stackoverflow site, this bit of code tells how many times the "foreach" loop needs to run. In my case I could probably have [0] or [1] after my collection because both of those lists had 46 items in them, so either way the "foreach" loop needs to run New-Object cmdlet 46 times.
    - "foreach" is a programming loop that runs New-Object cmdlet a specific number of times to create a two column matrix that has the same number of rows as the number of items in $GdocDirectory.
    - I know that New-Object cmdlet is creating the new array as a psobject (PowerShell object) with the properties of column 1 being named "id" and column 2 being named "name". My question: Why are the properties in @{ } with that @ symbol at the front of it? This is an area where I need to do some additional learning of PowerShell syntax. The "@{}" remind me of a json formatted file, and perhaps that is why it is written that way. It might be that it needs to be in {} because you have the iterative variable being evaluated within this section.
    - This section of code also reminds me that I still need to learn when to use {}, [], or () in PowerShell because these lines of code use all three at different times.

$GdocIDandName = ($GdocDoc_idArray,$GdocDirectory)[1] |
  foreach {New-Object psobject -property @{
	       id = $GdocDoc_idArray[$i]
	       name = $GdocDirectory[$i++]
	   }}

4. How do you export the .gdoc files into LibreOffice Writer format?
- First, the multi-column array is piped into "ForEach-Object" cmdlet which creates a loop for each row in the multi-column array.
- Second, using the "doc_id" "Export-GSDriveFile" cmdlet from PSGSuite module exports each file as a LibreOffice Writer file.
- Third, it outputs the exported file into the full file path with just the extension replaced in the file name from .gdoc to .odt.
- This line of code is really magical to me, doing something in seconds which would have taken so many searches and manual clicks in the GDrive web interface.

$GdocIDandName | ForEach-Object {Export-GSDriveFile -FileId $_.id -Type "OpenOfficeDoc" -OutFilePath ($_.name -replace '.gdoc','.odt')}

#>

# ACTUAL PWSH SCRIPT DOING THE WORK
# ---------------------------------
$GdocArray = Get-ChildItem -Path /home/mowest/Documents/*.gdoc -Recurse # Create a master collection . in this a list of files
#$GdocDoc_idArray = ForEach-Object {Get-Content $GdocArray | ConvertFrom-Json | Select-Object -ExpandProperty doc_id}
#$GdocDirectory = ForEach-Object {$GdocArray | Select-Object -ExpandProperty FullName}

$i = 0 #not neccessary We can use this to increment. I can give an example of its use.

$GdocArray| ForEach-Object { # instead of doing a foreach on each collection or property we produce 'one' collection and  then 'take action' oneitem at a time.

       Write-output "This is the $i file of the directory out of $gdocarray.count. Retrieving docid from file $_.fullname"
       $Docid = Get-Content $_.fullname | ConvertFrom-Json | Select-Object -ExpandProperty doc_id
       Export-GSDriveFile -FileId $docid -Type "OpenOfficeDoc" -OutFilePath ($_.name -replace '.gdoc','.odt') # We may need to test this . 
       Write-output " Converting $_.fullname to Officedoc"
       $i++

}
#$GdocIDandName = ($GdocDoc_idArray,$GdocDirectory)[1] |
 # foreach {New-Object psobject -property @{
#	       id = $GdocDoc_idArray[$i]
#	       name = $GdocDirectory[$i++]
#	   }}
#
#$GdocIDandName | ForEach-Object {Export-GSDriveFile -FileId $_.id -Type "OpenOfficeDoc" -OutFilePath ($_.name -replace '.gdoc','.odt')}
