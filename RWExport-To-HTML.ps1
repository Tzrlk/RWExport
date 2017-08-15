﻿<# 
.SYNOPSIS
RWExport-To-HTML.ps1
by EightBitz
Version 1.6
2017-08-14, 9:00 PM CDT

RWExport-To-HTML.ps1 transforms a Realm Works export file into a formatted HTML file. The formatted HTML file relies on a "RWExport_091b_1.6.css" file for formatting information.

LICENSE:
This script is now licensed under the Creative Commons Attribution + Non-Commercial license.
   
Basically, that means:
   -You are free to share and adapt the script.
   -When sharing the script, you must give appropriate credit and indicate if changes were made.

   Summary: https://creativecommons.org/licenses/by/4.0/
   Legal Code: https://creativecommons.org/licenses/by/4.0/legalcode

IMPORTANT:
The export from Realm Works must done with the "Compact Output" option. This script will likely not work with a "Full Export".
You can still export your full realm if you like, but make sure you do so with the "Compact Output" option.

IMPORTANT:
Make sure that the HTML file and the RWExport_091b_1.6.css file are in the same directory, otherwise, the HTML file will have no formatting.
For more information about this, type:

   Get-Help .\RWExport-To-HTML.ps1 -full

And see the NOTES section.

.DESCRIPTION 
RWExport-To-HTML.ps1 loads the XML file exported from Realm Works and transforms it into formatted HTML so it can be printed or imported/pasted into other programs. The export from Realm Works must done with the Compact Output option.
.PARAMETER Source 
Enter the full path and filename for the source file (the Realm Works export file).

IMPORTANT: This must be the first parameter on the command line.
.PARAMETER Destination 
Enter the full path and filename for the destination file (the HTML output).

IMPORTANT: This must be the second parameter on the command line.
(As long as the source and destination parameters are the first two, the remaining parameters are optional and can be used in any order.)
.PARAMETER Sort
Choose your preferred sort order for exported topics.
   1 = Name
   2 = Prefix, Name **Default**
   3 = Category, Name
   4 = Category, Prefix, Name
.PARAMETER Prefix
Include this parameter to display the prefix for each topic. If you've entered prefixes for your topics in Realm Works, this option will add them to the display in the HTML file in the form of "Prefix - Topic Name".
.PARAMETER Suffix
Include this parameter to display the suffix for each topic. If you've entered suffixes for your topics in Realm Works, this option will add them to the display in the HTML file in the form of "Topic Name (Suffix)".
If you include both the Prefix and Suffix parameters, the result will be "Prefix - Topic Name (Suffix)".
.PARAMETER Details
Include this parameter to include topic details (Category, Parent, Linkage, Tags, etc ...)
.PARAMETER Indent
Include this parameter to indent nested topics and section headers.
.PARAMETER SeparateSnippets
Include this parameter to display a line between snippets.
.PARAMETER $InlineStats
Include this parameter to display full stat blocks inline. This won't always look pretty, but if you want the option, it's there.
Whether or not you include this parameter, a given stat block will always be available to view by clicking a hyperlink, and this will display it with its native formatting, so it will look like you would expect it to look.
.PARAMETER SimpleImageScale
Include this parameter to scale the display size, by percentage, of embedded simple pictures.
If you omit this parameter or if you set it to 0, only the thumbnail will display.
.PARAMETER SmartImageScale
Include this parameter to scale the display size, by percentage, of embedded smart images (usually maps).
If you omit this parameter or if you set it to 0, only the thumbnail will display.
.PARAMETER KeepStyles
By default, this script strips some (not all) formatting from the imported data. It does this to allow formatting to be controlled by the the CSS file. Otherwise, the inline formatting will override the settings in the CSS file.
If you want to keep the original formatting, though, you can use this option.
Right now, the format options that are stripped are: font, font size, font color and background color.
Note that the CSS file has different definitions for regular snippet text, bulleted lists, numbered lists and tables.
.PARAMETER ExtractFiles
By default, this script creates one, single HTML file, with any file attachments encoded and embedded in the same, single HTML file.
If you wish, you can override this with the -ExtractFiles parameter. Including this on the command line will tell the script to save attachments as separate files.
If you include this option on the command line, you must specify a folder name for the destination instead of a file name. The HTML file will be stored in the folder you specify, and each file will be stored in a subfolder name "realm_files".
Also note that both the destination folder and the "realm_files" subfolder must already exist. The script will not create them for you.
.PARAMETER CSSFileName
By default, the HTML output file looks for "RWExport_091b_1.6.css" to define its style. You can use this option to specify a different name.
If you have two realms called Realm1 and Realm2, and you want each HTML output to have different fonts, font sizes or colors, you can specify a name or "realma.css" for one file and "realmb.css" for the other.
Not that this option does not create the file. It just tells the HTML file which filename to look for.
For now, at least, you will have to manually copy RWExport_091b_1.6.css or some other css file you like and make whatever changes you wish.
.PARAMETER SplitTopics
Save each topic as a separate HTML file.
If you include this option on the command line, you must specify a folder name for the destination instead of a file name. Each topic will be stored as an individual file within the folder you specify.
Also note that the folder must already exist. The script will not create it for you.
.PARAMETER Format
This option will allow you to choose your output format. Valid options are "HTML" and "Word". The default is "HTML".

If you specify "Word", several command line options will be ignored. These include:
   -Indent
   -SeparateSnippets
   -InlineStats
   -SimpleImageScale
   -SmartImageScale
   -KeepStyles
   -CSSFileName

.PARAMETER Log
Create a log file to record the progress of the conversion as well as any error messages.
.PARAMETER Force
Bypass user input validation.
.INPUTS
The .rwexport file from a Realm Works export that was created with the "Compact Output" option. This can an export of a full realm or a custom or partial export, just so long as it's made with the "Compact Output" option.
.OUTPUTS
An HTML file that uses an external style sheet for formatting. The external style sheet should be named "RWExport_091b_1.6.css" and should be in the same folder as the HTML output. The RWExport_091b_1.6.css file will not be generated by this script, but should already exist.
.EXAMPLE 
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html
This is the most basic example.
This would give you the most basic output. The text would be formatted according to Title, Topic, Section and Snippet, and that's about it. Nothing much more than that.

Topics will be sorted by prefix first, then by name.
Topics will be listed by name only with no prefixes or suffixes.
Nested topics and sections will not be indented. Everything will be left-justified.
Statblocks will not be displayed inline.
Simple Pictures and Smart Images will be displayed as thumbnails.
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -Sort 1
Choose your preferred sort order for exported topics. The example above will sort topics by their names.

Valid options are:
   1 = Sort by topic names
   2 = Sort by topic prefixes first, then by topic names (this is the default value)
   3 = Sort by topic category first, then by topic name
   4 = Sort by topic category first, then by topic prefix, then by topic name
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -Prefix
Prepend topic prefixes to topic names.

If a topic has a prefix of "Dungeon 1" and a name of "Room 3", you would include the Prefix parameter to list the topic in the output as "Dungeon 1 - Room 3"
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -Suffix
Append topic suffixes to topic names.

If a topic has a suffix of "Mess Hall" and a name of "Room 3", you would include the Suffix parameter to list the topic in the output as "Room 3 (Mess Hall)"

If you include the prefix and suffix parameter, the topic would be displayed as "Dungeon 1 - Room 3 (Mess Hall)"
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -Indent
To give your output just a little more style, you can choose to indent subsections and contained topics.
This looks really nice for documents that have content that might be nested 2 or 3 levels in, but for documents that might have 4, 5 or 6 levels of nested topics and subsections, it might not look so great.
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -InlineStats
By default, stat blocks will not display inline, but will instead offer a clickable link where you can then see the full stat block.

If you want everything in one view, though, without having to click, you can use the -InlineStats option.

Inline stat blocks will not always look as nice as you might like, but this parameter gives you the option to include them if you like.

Also, if you include both -Indent and -InlineStats, the stat blocks will NOT indent. That prospect was fraught with too many headaches.

As far as "looking nice", the clickable links are the better option, but if you want everything viewable in one document, this gives you that option.
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -SimpleImageScale 50
Display simple images inline at 50% of the width of the page.
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -SmartImageScale 50
Display smart images inline at 50% of the width of the page.
.NOTES
SORTING:

The specified sort order will occur regardless of whether or not prefixes are included.
If you sort by prefix, but do not include the -Prefix parameter, your topics will still be sorted by prefix, even though the prefix won't be displayed.
Likewise, if you sort by Name, but do include the -Prefix parameter, your topics will still be sorted by name, even though the prefix will be displayed.

It appears that topics will always be sorted under their containers. In other words, topics that are not in containers will be sorted relative to each other. Contained topics will be sorted relative to their peers within that container.



TIPS ON MANAGING COMMANDLINE OPTIONS:

There are a several commandline options for this script, and that can probably be intimidating to some people. Here is my recommended starting point:

   c:\<full path>\RWExport-To-HTML.ps1 -Source C:\<full path>\MyExport.rwoutput -Destination c:\<full path>\MyHTML.html -Indent -Prefix -Suffix -Sort 2

2 is the default sort value, so if that's what you want, you don't have to specify it, but I'm doing so anyway just to be clear as to what's happening in this example.

If you like the output from this combination of options, you don't have to rememeber to type this out all the time. You can save it as it's own PowerShell script. Paste the command into it's own .ps1 file. Name it something like "MyExportOptions.ps1". (Make sure to replace <full path> with the actual path for where the relevant files are.)

But you don't have to stop there. Say you have three different realms, and you want to specify different options for each. You can do something like this:

   c:\<full path>\RWExport-To-HTML.ps1 -Source C:\<full path>\Pathfinder-Realm.rwoutput -Destination c:\<full path>\Pathfinder-Realm.html -Indent -Prefix -Suffix -Sort 2

   c:\<full path>\RWExport-To-HTML.ps1 -Source C:\<full path>\SavageWorlds-Realm.rwoutput -Destination c:\<full path>\SavageWorlds-Realm.html -Indent -InlineStats

   c:\<full path>\RWExport-To-HTML.ps1 -Source C:\<full path>\FATE-Realm.rwoutput -Destination c:\<full path>\FATE-Realm.html -SimpleImageScale 25 -SmartImageScale 50

You can put all three of those command lines in the same PowerShell script and call it "ConvertAllMyExports.ps1".

Or you can put each command line in its own script called "ConvertPathfinderExport.ps1", "ConvertSavageWorldsExport.ps1" and "ConvertFATEExport.ps1"

Now, when you want convert a given export, you don't have to remember all the command line options, because you already have your favorite ones ready to go. You just run your PowerShell script instead of this one, and yours will invoke this one with your favorite options.

As a side note, you can set up a separate folder for each export, and put a different RWExport_091b_1.6.css file in each folder, so you can give each export a completely different look (by editing the RWExport_091b_1.6.css file in each folder).



AND SPEAKING OF RWExport_091b_1.6.css:

This script and its resulting HTML file assume the existence of a file named "RWExport_091b_1.6.css". If this file is not in the same folder as the resulting HTML file, it will not display correctly.

#> 

param (
    # Source file path and name.
    [Parameter(Mandatory,Position=1)] 
    [string]$Source,

    # Destination file path and name.
    [Parameter(Mandatory,Position=2)] 
    [string]$Destination,

    # Sort topics by:
    #    1 = Name
    #    2 = Prefix, Name **Default**
    #    3 = Category, Name
    #    4 = Category, Prefix, Name
    #
    # Choosing options 2 or 4 will sort by prefix, regardless of whether or not
    #    the -Prefix switch is specified. Likewise, choosing options 1 or 3 will
    #    sort by name, regardless of whether or not the -Prefix switch is specified.
    [Parameter()] 
    [int]$Sort = 2,

    # Include prefix in topic name ($true or $false)
    [switch]$Prefix,

    # Include suffix in topic name ($true or $false)
    [switch]$Suffix,

    # Include topic details (Category, Parent, Linkage, Tags, etc ...)
    [switch]$Details,

    # Indent nested topics and sections.
    [switch]$Indent,

    # Include a separator line between snippets.
    [switch]$SeparateSnippets,

    # Display Statblocks inline.
    [switch]$InlineStats,

    # Scale percentage for displaying Simple Pictures inline.
    #    Omitting this parameter, or setting it to 0
    #    will display thumbnails.
    [Parameter()] 
    [int]$SimpleImageScale = 0,

    # Scale percentage for displaying Smart Images inline.
    #    Omitting this parameter, or setting it to 0
    #    will display thumbnails.
    [Parameter()] 
    [int]$SmartImageScale = 0,

    # Preserve the text styles as defined in Realm Works. 
    [Parameter()] 
    [switch]$KeepStyles,

    # Preserve the text styles as defined in Realm Works. 
    [Parameter()] 
    [switch]$ExtractFiles,

    # Specify a different name for the CSS file.
    # This does not create the CSS file. It only changes the name
    #    defined in the header of the HTML file.
    [Parameter()] 
    [string]$CSSFileName = "RWExport_091b_1.6.css",

    # Preserve the text styles as defined in Realm Works. 
    [Parameter()] 
    [switch]$SplitTopics,

    # Specify the path for an optional log file.
    [Parameter()] 
    [string]$Format = "HTML",

    # Specify the path for an optional log file.
    [Parameter()] 
    [string]$Log,

    # Do not validate user input.
    [Parameter()] 
    [switch]$Force

) # param

Function ParseTopic($PassedTopic,$Outputfile,$Sort,$Prefix,$Suffix,$Details,$Indent,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$Parent,$TitleCSS,$TopicCSS,$TopicDetailsCSS,$SectionCSS,$SnippetCSS,$CurrentIndent,$Indcrement,$KeepStyles,$ExtractFiles,$ExportPath,$SplitTopics,$Log,$CSSFileName,$CommandLine,$Format,$wordfile) {
   Try {
      $TopicName = $PassedTopic.public_name
      $TopicID = $PassedTopic.topic_id
      if ($Log) {[System.IO.File]::AppendAllText($Log,"`r`nTopic: $TopicName`r`n")}
      if ($SplitTopics) {
          
         $TopicFile = $TopicName + "_" + $TopicID
         # if ($PassedTopic.Prefix) {$TopicFile = $PassedTopic.Prefix + "-" + $TopicFile}
         # if ($PassedTopic.Suffix) {$TopicFile = $TopicFile + "(" + $PassedTopic.Suffix + ")"}

         $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
         for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
            $TopicFile = $TopicFile.replace($invalidChars[$CheckChar],"_")
         } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)

         Switch ($Format) {
            "Word" {
               $wordfile = InitializeWordFile
               $Topicfile = $OutputFile + $TopicFile + ".doc"
            } # "Word"
            "HTML" {
               $HTMLFile = $OutputFile + $TopicFile + ".html"
               $TopicFile = $HTMLFile
               WriteHTMLHeader $TopicFile $CSSFileName $CommandLine
            } # "HTML"
         } # Switch ($Format)
      } else {
         $TopicFile = $OutputFile
      } # if ($SplitTopics)

      if ($Prefix -and $PassedTopic.Prefix) {$TopicName = $PassedTopic.Prefix + " - " + $TopicName}
      if ($Suffix -and $PassedTopic.Suffix) {$TopicName = $TopicName + " (" + $PassedTopic.Suffix + ")"}

      $TopicDetails = $null
      If ($PassedTopic.alias) {
         $Aliases = ParseAliases $PassedTopic $TopicFile $ExportTopics $Log
         $TopicDetails = $Aliases
      } # If ($PassedTopic.linkage)

      $ParentName = "Parent Topic: $Parent"
      $CategoryName = "Category: " + $PassedTopic.category_name.Trim()

      If ($PassedTopic.tag_assign) {
         $tagline = ParseTags $PassedTopic $TopicFile $Log
      } # If ($PassedTopic.tag_assign)

      If ($PassedTopic.linkage) {
         [array]$linkage = ParseLinkage $PassedTopic $TopicFile $SplitTopics $Log $Format
      } # If ($PassedTopic.linkage)
   } Catch {
      ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
   } # Try

   Try {
      Switch ($Format) {
         "Word" {
            $selection = $wordfile.selection
            $selection.Style = "Heading 1"
            $selection.typeText($TopicName)
            $Selection.typeParagraph()
            if ($Details) {
               $selection.Style = "Topic_Details"
               # $selection.paragraphs().shading.BackgroundPatternColor = 16764057
               if ($Aliases) {$selection.typeText($Aliases);$selection.InsertBreak(6)}
               $selection.typeText($ParentName);$selection.InsertBreak(6)
               $selection.typeText($CategoryName);$selection.InsertBreak(6)
               if ($tagline) {$selection.typeText($tagline);$selection.InsertBreak(6)}
               if ([array]$Linkage.count -gt 1) {
                  for ($link = 0; $link -le $Linkage.count; $link++) {
                     $selection.typeText($linkage[$link])
                     if ($link -lt $Linkage.count-1) {$selection.InsertBreak(6)} else {$selection.typeparagraph()}  
                  } # for ($link = 0; $link -le $Linkage.count; $link++)
               } else {
                  $selection.typeText($linkage)
                  $selection.typeParagraph()
               } # if ([array]$Linkage.count -gt)
            } # if ($Details)

         } # "Word"
         "HTML" {
            if ($Aliases) {$TopicDetails = "$Aliases<BR>"}
            $TopicDetails = $TopicDetails + "$ParentName<BR>$CategoryName"
            if ($tagline) {$TopicDetails = "$TopicDetails<BR>$tagline"}
            if ($linkage) {$TopicDetails = "$TopicDetails<BR>$linkage"}
            $TopicDetails = $TopicDetailsCSS.Replace("*",$TopicDetails)
            $TopicName = $TopicCSS.Replace("*",$TopicName)
            $TopicName = $TopicName.Replace('id=""','id="' + $TopicID + '"')
            [System.IO.File]::AppendAllText($TopicFile,$TopicName)
            If ($Details) {[System.IO.File]::AppendAllText($TopicFile,$TopicDetails)}
         } # "HTML"
      } # Switch ($Format)
   } Catch {
      ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
   } # Try

   [int32]$SectionLevel = 2
   foreach ($Section in $PassedTopic.section) {
      ParseSection $PassedTopic.public_name $Section $TopicFile $SectionCSS $SnippetCSS $CurrentIndent $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $ExtractFiles $ExportPath $Log $Format $wordfile $SectionLevel
   } # foreach ($Section in $PassedTopic.section)

   if ($SplitTopics) {
      Switch ($Format) {
         "Word" {
            CloseWordFile $TopicFile $wordfile
         } # "Word"
         "HTML" {
            WriteHTMLFooter $TopicFile $CommandLine
         } # "HTML"
      } # Switch ($Format)
   } # if ($SplitTopics)
   Switch ($Sort) {
      1 {$TopicList = $PassedTopic.topic | Sort-Object public_name}
      2 {$TopicList = $PassedTopic.topic | Sort-Object prefix,public_name}
      3 {$TopicList = $PassedTopic.topic | Sort-Object category_name,public_name}
      4 {$TopicList = $PassedTopic.topic | Sort-Object category_name,prefix,public_name}
      default {$TopicList = $PassedTopic.topic}
   } # Switch ($Sort)

   $Parent = $Parent + $PassedTopic.public_Name + "/"
   foreach ($Topic in $TopicList) {
      If ($Indent) {
         $SubTopicCSS = AddIndent $TopicCSS $Indcrement $Log
         $SubTopicDetailsCSS = AddIndent $TopicDetailsCSS $Indcrement $Log
         $SubSectionCSS = AddIndent $SectionCSS $Indcrement $Log
         $SubIndent = AddIndent $CurrentIndent $Indcrement $Log
      } else {
         $SubTopicCSS = $TopicCSS
         $SubTopicDetailsCSS = $TopicDetailsCSS
         $SubSectionCSS = $SectionCSS
         $SubIndent = $CurrentIndent
      } # If ($Indent)
      ParseTopic $Topic $OutputFile $Sort $Prefix $Suffix $Details $Indent $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $Parent $TitleCSS $SubTopicCSS $SubTopicDetailsCSS $SubSectionCSS $SnippetCSS $SubIndent $Indcrement $KeepStyles $ExtractFiles $ExportPath $SplitTopics $Log $CSSFileName $CommandLine $Format $wordfile
   } # foreach ($Topic in $TopicList)
} # Function ParseTopic($PassedTopic)

Function ParseSection ($TopicName,$PassedSection,$Outputfile,$SectionCSS,$SnippetCSS,$CurrentIndent,$Indcrement,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$KeepStyles,$ExtractFiles,$ExportPath,$Log,$Format,$wordfile,$SectionLevel) {
   Try {
      Switch ($Format) {
         "Word" {
            $SectionName = $PassedSection.name
            $selection = $wordfile.selection
            $selection.Style = "Heading $SectionLevel"
            $selection.typeText($SectionName)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            if ($Log) {[System.IO.File]::AppendAllText($Log,"Section: $SectionName`r`n")}
            $SectionName = $SectionCSS.Replace("*",$PassedSection.name)
            [System.IO.File]::AppendAllText($outputfile,$sectionname)
            $DivCSS = '<div class="snippet" ' + $CurrentIndent + '>'
            [System.IO.File]::AppendAllText($outputfile,$DivCSS)
         } # "HTML"
      } # Switch
   } Catch {
      ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
   } # Try

   foreach ($Snippet in $PassedSection.snippet) {
      ParseSnippet $TopicName $PassedSection.name $Snippet $Outputfile $SnippetCSS $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $ExtractFiles $ExportPath $Log $Format $wordfile $SectionLevel
   } # foreach ($Snippet in $PassedSection.snippet)

   if ($Format -eq "HTML") {[System.IO.File]::AppendAllText($outputfile,'</div>')}

   foreach ($Section in $PassedSection.section) {
      If ($Indent) {
         $SubSectionCSS = AddIndent $SectionCSS $Indcrement $Log
         $SubIndent = AddIndent $CurrentIndent $Indcrement $Log
         if ($SectionLevel -le 9) {$SubLevel = $SectionLevel + 1}
      } else {
         $SubSectionCSS = $SectionCSS
         $SubIndent = $CurrentIndentCSS
         $SubLevel = $SectionLevel
      } # If ($Indent)
      ParseSection $TopicName $Section $Outputfile $SubSectionCSS $SnippetCSS $SubIndent $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $ExtractFiles $ExportPath $Log $Format $wordfile $SubLevel
   } # foreach ($Section in $PassedSection.section)
} # Function ParseSection ($PassedSection)

Function ParseSnippet ($TopicName,$SectionName,$PassedSnippet,$Outputfile,$SnippetCSS,$Indcrement,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$KeepStyles,$ExtractFiles,$ExportPath,$Log,$Format,$wordfile,$sectionlevel) {
   switch ($PassedSnippet.type) { 
      "Audio" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Audio</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               Switch ($Extension) {
                  "aac" {$audioLink = '<a href="data:audio/x-aac;base64,*">Audio</a>'}
                  "aiff" {$audioLink = '<a href="data:audio/aiff;base64,*">Audio</a>'}
                  "alac" {$audioLink = '<a href="data:audio/m4a;base64,*">Audio</a>'}
                  "flac" {$audioLink = '<a href="data:audio/flac;base64,*">Audio</a>'}
                  "mp3" {$audioLink = '<a href="data:audio/mpeg;base64,*">Audio</a>'}
                  "m4a" {$audioLink = '<a href="data:audio/mp4;base64,*">Audio</a>'}
                  "oga" {$audioLink = '<a href="data:audio/ogg;base64,*">Audio</a>'}
                  "ogg" {$audioLink = '<a href="data:audio/ogg;base64,*">Audio</a>'}
                  "wav" {$audioLink = '<a href="data:audio/wav;base64,*">Audio</a>'}
                  "wma" {$audioLink = '<a href="data:audio/x-ms-wma;base64,*">Audio</a>'}
                  "wv" {$audioLink = '<a href="data:audio/wavpack;base64,*">Audio</a>'}
                  default {$ImageLink = '<a href="data:application/octet-stream;base64,*">Audio</a>'}
               } # switch ($Extension)
               $DataLink = $audiolink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External Audio File)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "")) {
                      $annotation = "Annotation: $annotation"
                      $selection.typeText($annotation)
                   }
                   $selection.TypeParagraph() 
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)

         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Audio"

      "Date_Game" {
         Try {
            $Text = $PassedSnippet.game_date.display
            if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
               if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
               $Text = $LabelPrefix + $Text
            } # if ($PassedSnippet.Label -ne $null)

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            Switch ($Format) {
               "Word" {
                  $selection = $wordfile.selection
                  $selection.Style = "Normal"
                  if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;")) {
                     $selection.typeText($Text)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {

                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $Text = $Text + "<br>Annotation: $annotation"
                  } # if (($annotation -ne $null) -and ($annotation -ne ""))
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)

        } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
        } # Try
      } # Date_Game

      "Date_Range" {
         Try {
            $Text = $PassedSnippet.date_range.display_start + " to " + $PassedSnippet.date_range.display_end
            if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
               if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
               $Text = $LabelPrefix + $Text
            } # if ($PassedSnippet.Label -ne $null)
         
            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log
            Switch ($Format) {
               "Word" {
                  $selection = $wordfile.selection
                  $selection.Style = "Normal"
                  if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;")) {
                     $selection.typeText($Text)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {

                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $Text = $Text + "<br>Annotation: $annotation"
                  } # if (($annotation -ne $null) -and ($annotation -ne ""))
                  
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try

      } # Date_Range

      "Foreign" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Foreign Object</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $ForeignLink = '<a href="data:application/octet-stream;base64,*">Foreign Object</a>'
               $DataLink = $ForeignLink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External Foreign Object)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $annotation = "Annotation: $annotation"
                      $selection.InsertBreak(6) # Line break.
                      $selection.typeText($annotation)
                   } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                   $selection.typeParagraph()
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Foreign"

      "HTML" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">HTML</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $HTMLLink = '<a href="data:text/html;base64,*">HTML</a>'
               $DataLink = $HTMLLink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External HTML File)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "HTML"

      "Labeled_Text" {
         Try {
             $Text = $PassedSnippet.contents

             $TagPreface = '<span class="RWSnippet">'
             if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
                $Text = $Text.insert($InsertPoint,$LabelPrefix)
             } # if ($PassedSnippet.Label -ne $null)
         
             Switch ($Format) {
               "Word" {
                  $Text = ParseSnippetText $Text $Log
                  $selection = $wordfile.selection
                  $selection.Style = "Normal"
                  if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;")) {
                     $selection.typeText($Text.Trim())
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {
                  $TagPreface = '<span class="RWSnippet">'
                  if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                  if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                     $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
                     $Text = $Text.insert($InsertPoint,$LabelPrefix)
                  } # if ($PassedSnippet.Label -ne $null)
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Labeled_Text"

      "Multi_Line" {
         Try {
            if ($PassedSnippet.purpose -eq "directions_only") {
               $Text = $PassedSnippet.gm_directions
               Switch ($Format) {
                  "Word" {
                     $Text = ParseSnippetText $Text
                     $selection = $wordfile.selection
                     $selection.Style = "GM_Directions"

                     # $selection.paragraphs().shading.BackgroundPatternColor = 10079487
                     if ($Text.Count -ge 1) {$selection.typeText("GM Directions: ")}
                     foreach ($paragraph in $Text) {
                        if (($paragraph -ne $null) -and ($paragraph -ne "")) {
                           $selection.typeText($paragraph)
                           $selection.typeParagraph()
                        } # if (($paragraph -ne $null) -and ($paragraph -ne ""))
                     } # foreach ($paragraph in $Text)
                  } # "Word"
                  "HTML" {
                     $TagPreface = '<span class="RWSnippet">'
                     $LabelPrefix = 'GM Directions: '
                     $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
                     $Text = $Text.insert($InsertPoint,$LabelPrefix)
                     If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                     if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                     [System.IO.File]::AppendAllText($outputfile,$Text)
                  } # "HTML"
               } # Switch ($Format)
            } elseif ($PassedSnippet.purpose -eq "Both") {
               Switch ($Format) {
                  "Word" {
                     if ($PassedSnippet.gm_directions -ne $null) {
                        $Text = $PassedSnippet.gm_directions
                        $Text = ParseSnippetText $Text
                        $selection = $wordfile.selection
                        $selection.Style = "GM_Directions"
                        # $selection.paragraphs().shading.BackgroundPatternColor = 10079487 # 26367 # 1028336
                        if ($Text.Count -ge 1) {$selection.typeText("GM Directions: ")}
                        foreach ($paragraph in $Text) {
                           if (($paragraph -ne $null) -and ($paragraph -ne "")) {
                              $selection.typeText($paragraph)
                              $selection.typeParagraph()
                           } # if (($paragraph -ne $null) -and ($paragraph -ne ""))
                        } # foreach ($paragraph in $Text)
                     } # if ($PassedSnippet.gm_directions -ne $null)

                     if ($PassedSnippet.contents -ne $null) {
                        $Text = $PassedSnippet.contents
                        [array]$Text = ParseSnippetText $Text
                        foreach ($paragraph in $Text) {
                           $selection = $wordfile.selection
                           $selection.Style = "Normal"
                           $selection.typeText($paragraph)
                           $Selection.typeParagraph()
                        } # foreach ($paragraph in $Text)
                     } # if ($PassedSnippet.contents -ne $null)
                  } # "Word"
                  "HTML" {
                     if ($PassedSnippet.gm_directions -ne $null) {
                        $Text = $PassedSnippet.gm_directions
                        $TagPreface = '<span class="RWSnippet">'
                        $LabelPrefix = 'GM Directions: '
                        $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
                        $Text = $Text.insert($InsertPoint,$LabelPrefix)
                        If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                        # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                        [System.IO.File]::AppendAllText($outputfile,$Text)
                     } # if ($PassedSnippet.gm_directions -ne $null)

                     if ($PassedSnippet.contents -ne $null) {
                        $Text = $PassedSnippet.contents
                        If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                        if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                        [System.IO.File]::AppendAllText($outputfile,$Text)
                     } # if ($PassedSnippet.contents -ne $null)
                  } # "HTML"
               } # Switch ($Format)
            } elseif (($PassedSnippet.contents -ne $null) -and (($PassedSnippet.contents.contains("<ul")) -or ($PassedSnippet.contents.contains("<ol")) -or ($PassedSnippet.contents.contains("<table")))) {
               $Text = $PassedSnippet.contents
               Switch ($Format) {
                  "Word" {
                     ParseFormattedStuff $Text $Format $wordfile
                  } # "Word"
                  "HTML" {
                     If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                     if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                     [System.IO.File]::AppendAllText($outputfile,$Text)
                  } # "HTML"
               } # Switch ($Format)
               #>
            } else {
               $Text = $PassedSnippet.contents
               Switch ($Format) {
                  "Word" {
                     [array]$Text = ParseSnippetText $Text
                     foreach ($paragraph in $Text) {
                        $selection = $wordfile.selection
                        $selection.Style = "Normal"
                        $selection.typeText($paragraph)
                        $Selection.typeParagraph()
                     } # foreach ($paragraph in $Text)
                  } # "Word"
                  "HTML" {
                     If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                     if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                     [System.IO.File]::AppendAllText($outputfile,$Text)
                  } # "HTML"
               } # Switch ($Format)
            } # if ($PassedSnippet.purpose -eq "directions_only")
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Multi_Line"

      "Numeric" {
         Try {
             $Text = $PassedSnippet.contents
             if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                $Text = $LabelPrefix + $Text
             } # if ($PassedSnippet.Label -ne $null)

             $annotation = $PassedSnippet.annotation
             $annotation = ParseSnippetText $annotation $Log
             Switch ($Format) {
               "Word" {
                  $selection = $wordfile.selection
                  $selection.Style = "Normal"
                  if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;")) {
                     $selection.typeText($Text)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {

                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $Text = $Text + "<br>Annotation: $annotation"
                  } # if (($annotation -ne $null) -and ($annotation -ne ""))
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Numeric"

      "PDF" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">PDF</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $PDFLink = '<a href="data:application/pdf;base64,*">PDF</a>'
               $DataLink = $PDFlink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External PDF Document)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "PDF"

      "Picture" {
         Try {
             $FullImage = $PassedSnippet.ext_object.asset.contents
             $Thumbnail = $PassedSnippet.ext_object.asset.thumbnail
             $ImageName = $PassedSnippet.ext_object.name
             $FileName = $PassedSnippet.ext_object.asset.filename
             if (($ImageName -eq $null) -or ($ImageName -eq "") -or ($ImageName -eq "&nbsp;")) {$ImageName = "Unnamed"}

             if ($ExtractFiles) {
                $FilePath = $ExportPath + $FileName
                $ThumbName = "thumb_" + $PassedSnippet.ext_object.asset.filename
                $ThumbPath = $ExportPath + $ThumbName
                $SplitExportPath = $ExportPath.split("\")
                $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
                $bytes = [Convert]::FromBase64String($FullImage)
                [IO.File]::WriteAllBytes($FilePath,$bytes)
                $bytes = [Convert]::FromBase64String($ThumbNail)
                [IO.File]::WriteAllBytes($ThumbPath, $bytes)
                $ImageLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Picture</a>'
                If ($SimpleImageScale -eq 0) {
                   $EncodedImage = '<img src="' + $ExportSubFolder + $ThumbName + '">'
                } else {
                   $EncodedImage = '<img width="' + $SimpleImageScale + '%" src="' + $ExportSubFolder + $FileName + '">'
                } # If ($SimpleImageScale -eq 0)
             } else {
                $SplitFileName = $FileName.split(".")
                $Extension = $SplitFileName[$SplitFileName.Count-1]
                Switch ($Extension) {
                   "bmp" {$ImageLink = '<a href="data:image/bmp;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/bmp;base64,*">'}
                   "gif" {$ImageLink = '<a href="data:image/gif;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/gif;base64,*">'}
                   "jpeg" {$ImageLink = '<a href="data:image/jpeg;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/jpeg;base64,*">'}
                   "jpg" {$ImageLink = '<a href="data:image/jpeg;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/jpeg;base64,*">'}
                   "png" {$ImageLink = '<a href="data:image/png;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/png;base64,*">'}
                   "tif" {$ImageLink = '<a href="data:image/x-tiff;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/x-tiff;base64,*">'}
                   "tiff" {$ImageLink = '<a href="data:image/x-tiff;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/x-tiff;base64,*">'}
                    default {$ImageLink = '<a href="data:application/octet-stream;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:application/octet-stream;base64,*">'}
                } # switch ($Extension)

                $ImageLink = $ImageLink.replace("*",$FullImage)
                If ($SimpleImageScale -eq 0) {
                   $EncodedImage = $EmbeddedLink.replace("*",$Thumbnail)
                } else {
                   $Scale = 'width="' + $SimpleImageScale + '%"'
                   $EncodedImage = $EmbeddedLink.replace("*",$FullImage)
                   $EncodedImage = $EncodedImage.Replace("<img src=","<img $Scale src=")
                } # If ($SimpleImageScale -eq 0)
             } # if ($ExtractFiles)

             $Text = $ImageName + ": [$ImageLink]"

             $annotation = $PassedSnippet.annotation
             $annotation = ParseSnippetText $annotation $Log

             Switch ($Format) {
                "Word" {
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($ImageName)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $annotation = "Annotation: $annotation"
                      $selection.InsertBreak(6) # Line break.
                      $selection.typeText($annotation)
                   } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                   $selection.InsertBreak(6) # Line break.
                   $ImageStream=[System.IO.MemoryStream][System.Convert]::FromBase64String($PassedSnippet.ext_object.asset.contents)
                   $ImageBmp=[System.Drawing.Bitmap][System.Drawing.Image]::FromStream($ImageStream)
                   [Windows.Forms.Clipboard]::SetImage($ImageBmp)
                   $selection.Paste()
                   $selection.TypeParagraph() 
                   [Windows.Forms.Clipboard]::Clear()
                } #"Word" 
                "HTML" {
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Picture"

      "Portfolio" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Hero Lab Portfolio</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $ForeignLink = '<a href="data:application/octet-stream;base64,*">Hero Lab Portfolio</a>'
               $DataLink = $ForeignLink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (Hero Lab Portfolio)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try

      } # "Portfolio"

      "Rich_Text" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">RTF</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $RTFLink = '<a href="data:text/rtf;base64,*">RTF</a>'
               $DataLink = $RTFLink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External RTF Document)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try

      } # "Rich Text"

      "Smart_Image" {
         Try {
             $FullImage = $PassedSnippet.smart_image.asset.contents
             $Thumbnail = $PassedSnippet.smart_image.asset.thumbnail
             $ImageName = $PassedSnippet.smart_image.name
             $FileName = $PassedSnippet.smart_image.asset.filename
             if (($ImageName -eq $null) -or ($ImageName -eq "") -or ($ImageName -eq "&nbsp;")) {$ImageName = "Unnamed"}

             if (($FileName -eq $null) -or ($FileName -eq "") -or ($FileName -eq "&nbsp;")) {
                $ImageName = "[No Filename]"
                $EncodedImage = "[No Image Data]"
                $annotation = $PassedSnippet.annotation
                $annotation = ParseSnippetText $annotation $Log
             } else {
                 if ($ExtractFiles) {
                    $FilePath = $ExportPath + $FileName
                    $ThumbName = "thumb_" + $PassedSnippet.smart_image.asset.filename
                    $ThumbPath = $ExportPath + $ThumbName
                    $SplitExportPath = $ExportPath.split("\")
                    $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
                    $bytes = [Convert]::FromBase64String($FullImage)
                    [IO.File]::WriteAllBytes($FilePath,$bytes)
                    $bytes = [Convert]::FromBase64String($ThumbNail)
                    [IO.File]::WriteAllBytes($ThumbPath, $bytes)
                    $ImageLink = '<a href="' + $ExportSubFolder + $PassedSnippet.smart_image.asset.filename + '">Picture</a>'
                    If ($SmartImageScale -eq 0) {
                       $EncodedImage = '<img src="' + $ExportSubFolder + $ThumbName + '">'
                    } else {
                       $EncodedImage = '<img width="' + $SmartImageScale + '%" src="' + $ExportSubFolder + $FileName + '">'
                    } # If ($SmartImageScale -eq 0)
                 } else {
                    $SplitFileName = $FileName.split(".")
                    $Extension = $SplitFileName[$SplitFileName.Count-1]
                    Switch ($Extension) {
                       "bmp" {$ImageLink = '<a href="data:image/bmp;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/bmp;base64,*">'}
                       "gif" {$ImageLink = '<a href="data:image/gif;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/gif;base64,*">'}
                       "jpeg" {$ImageLink = '<a href="data:image/jpeg;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/jpeg;base64,*">'}
                       "jpg" {$ImageLink = '<a href="data:image/jpeg;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/jpeg;base64,*">'}
                       "png" {$ImageLink = '<a href="data:image/png;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/png;base64,*">'}
                       "tif" {$ImageLink = '<a href="data:image/x-tiff;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/x-tiff;base64,*">'}
                       "tiff" {$ImageLink = '<a href="data:image/x-tiff;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:image/x-tiff;base64,*">'}
                       default {$ImageLink = '<a href="data:application/octet-stream;base64,*">Picture</a>';$EmbeddedLink = '<img src="data:application/octet-stream;base64,*">'}
                    } # switch ($Extension)

                    $ImageLink = $ImageLink.replace("*",$FullImage)
                    if ($SmartImageScale -eq 0) {
                       $EncodedImage = $EmbeddedLink.replace("*",$Thumbnail)
                    } else {
                       $Scale = 'width="' + $SmartImageScale + '%"'
                       $EncodedImage = $EmbeddedLink.replace("*",$FullImage)
                       $EncodedImage = $EncodedImage.Replace("<img src=","<img $Scale src=")
                    } # If ($SmartImageScale -eq 0)
                 } # if ($ExtractFiles)

                 $Text = $ImageName + ": [$ImageLink]"

                 $annotation = $PassedSnippet.annotation
                 $annotation = ParseSnippetText $annotation $Log
            } # if (($FileName -eq $null) -or ($FileName -eq "") -or ($FileName -eq "&nbsp;"))

             Switch ($Format) {
                "Word" {
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($ImageName)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                     $selection.InsertBreak(6) # Line break.
                   } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                   if (($FileName -eq $null) -or ($FileName -eq "") -or ($FileName -eq "&nbsp;")) {
                       $selection.TypeParagraph()
                       [Windows.Forms.Clipboard]::Clear()
                   } else {
                       $ImageStream=[System.IO.MemoryStream][System.Convert]::FromBase64String($PassedSnippet.smart_image.asset.contents)
                       $ImageBmp=[System.Drawing.Bitmap][System.Drawing.Image]::FromStream($ImageStream)
                       [Windows.Forms.Clipboard]::SetImage($ImageBmp)
                       $selection.Paste()
                       $selection.TypeParagraph() 
                       [Windows.Forms.Clipboard]::Clear()
                   } # if (($FileName -eq $null) -or ($FileName -eq "") -or ($FileName -eq "&nbsp;"))
                } # "Word" 
                "HTML" {
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
                } # "HTML"
             } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
         
      } # "Smart_Image"

      "Statblock" {
         Try {
            $Data = $PassedSnippet.ext_object.asset.contents
            $Name = $PassedSnippet.ext_object.name
            $FileName = $PassedSnippet.ext_object.asset.filename
            if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            if ($ExtractFiles) {
               $FilePath = $ExportPath + $FileName
               $SplitExportPath = $ExportPath.split("\")
               $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
               $bytes = [Convert]::FromBase64String($Data)
               [IO.File]::WriteAllBytes($FilePath,$bytes)
               $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Statblock</a>'
            } else {
               $SplitFileName = $FileName.split(".")
               $Extension = $SplitFileName[$SplitFileName.Count-1]
               $StatLink = '<a href="data:text/html;base64,*">Statblock</a>'
               $DataLink = $Statlink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External Statblock)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph() 
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Statblock"

      "Tag_Multi_Domain" {
         Try {
             $taglist = ParseTags $PassedSnippet $Log

             $annotation = $PassedSnippet.annotation
             $annotation = ParseSnippetText $annotation $Log

             $Text = $null
             Switch ($Format) {
               "Word" {
                  $selection = $wordfile.selection
                  $selection.style = "Normal"
                  if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                     if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                     $selection.typeText($LabelPrefix)
                     $selection.InsertBreak(6) # Line break.
                  } # if ($PassedSnippet.Label -ne $null)

                  for ($tag=0;$tag -le $taglist.count-1;$tag++) {
                     if (($taglist[$tag] -ne $null) -and ($taglist[$tag] -ne "") -and ($taglist[$tag] -ne "&nbsp;")) {
                        $selection.typeText($taglist[$tag])
                        if ($tag -lt $taglist.count-1) {$selection.InsertBreak(6)}
                     } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  }
                  if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {

                  for ($tag=0;$tag -le $taglist.count-1;$tag++) {
                     if (($taglist[$tag] -ne $null) -and ($taglist[$tag] -ne "") -and ($taglist[$tag] -ne "&nbsp;")) {
                        $Text = $Text + $taglist[$tag]
                        if ($tag -lt $taglist.count-1) {$Text = "$Text<BR>"}
                     } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  }
                  
                  if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                     if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                     $Text = $LabelPrefix + "<BR>$Text"
                   } # if ($PassedSnippet.Label -ne $null
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $Text = $Text + "<br>Annotation: $annotation"
                  } # if (($annotation -ne $null) -and ($annotation -ne ""))
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Tag_Multi_Domain"

      "Tag_Standard" {
         Try {
            $Taglist = ParseTags $PassedSnippet $Log

            $annotation = $PassedSnippet.annotation
            $annotation = ParseSnippetText $annotation $Log

            Switch ($Format) {
               "Word" {
                  $selection = $wordfile.selection
                  $selection.Style = "Normal"
                  if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                     if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                     $selection.typeText($LabelPrefix)
                     $selection.InsertBreak(6) # Line break.
                  } # if ($PassedSnippet.Label -ne $null)

                  
                  if (($taglist -ne $null) -and ($taglist -ne "") -and ($taglist -ne "&nbsp;")) {
                     $selection.typeText($taglist)
                     # $selection.InsertBreak(6)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  
                  if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph()
               } # "Word"
               "HTML" {
                  if (($taglist -ne $null) -and ($taglist -ne "") -and ($taglist -ne "&nbsp;")) {
                     $Text = $taglist
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))


                  if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                     if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                     $Text = $LabelPrefix + "<BR>$Text"
                  } # if ($PassedSnippet.Label -ne $null)

                  if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $Text = $Text + "<br>Annotation: $annotation"
                  } # if (($annotation -ne $null) -and ($annotation -ne ""))
                  # if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  $Text = $SnippetCSS.Replace("*",$Text)
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
            ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Tag_Standard"

      "Video" {
         Try {
             $Data = $PassedSnippet.ext_object.asset.contents
             $Name = $PassedSnippet.ext_object.name
             $FileName = $PassedSnippet.ext_object.asset.filename
             if (($Name -eq $null) -or ($Name -eq "") -or ($Name -eq "&nbsp;")) {$Name = "Unnamed"}

             $annotation = $PassedSnippet.annotation
             $annotation = ParseSnippetText $annotation $Log

             if ($ExtractFiles) {
                $FilePath = $ExportPath + $FileName
                $SplitExportPath = $ExportPath.split("\")
                $ExportSubFolder = $SplitExportPath[$SplitExportPath.Count-2] + "/"
                $bytes = [Convert]::FromBase64String($Data)
                [IO.File]::WriteAllBytes($FilePath,$bytes)
                $DataLink = '<a href="' + $ExportSubFolder + $PassedSnippet.ext_object.asset.filename + '">Video</a>'
             } else {
                $SplitFileName = $FileName.split(".")
                $Extension = $SplitFileName[$SplitFileName.Count-1]
                Switch ($Extension) {
                   "avi" {$videoLink = '<a href="data:video/avi;base64,*">Video</a>'}
                   "flv" {$videoLink = '<a href="data:video/x-flv;base64,*">Video</a>'}
                   "hdmov" {$videoLink = '<a href="data:video/quicktime;base64,*">Video</a>'}
                   "m2ts" {$videoLink = '<a href="data:video/MP2T;base64,*">Video</a>'}
                   "m4v" {$videoLink = '<a href="data:video/x-m4v;base64,*">Video</a>'}
                   "mkv" {$videoLink = '<a href="data:video/x-matroska;base64,*">Video</a>'}
			       "mov" {$videoLink = '<a href="data:video/quicktime;base64,*">Video</a>'}
			       "mp4" {$videoLink = '<a href="data:video/mp4;base64,*">Video</a>'}
                   "mpg" {$videoLink = '<a href="data:video/mpeg;base64,*">Video</a>'}
                   "ogv" {$videoLink = '<a href="data:video/ogg;base64,*">Video</a>'}
                   "ts" {$videoLink = '<a href="data:video/MP2T;base64,*">Video</a>'}
                   "wmv" {$videoLink = '<a href="data:video/x-tiff;base64,*">Video</a>'}
                   default {$ImageLink = '<a href="data:application/octet-stream;base64,*">Video</a>'}
                } # switch ($Extension)
                $DataLink = $VideoLink.replace("*",$Data)
             } # if ($ExtractFiles)

             Switch ($Format) {
                "Word" {
                   $Text = "$Name (External Video File)"
                   $selection = $wordfile.selection
                   $selection.Style = "Normal"
                   $selection.typeText($Text)
                   # $selection.InsertBreak(6) # Line break.
                   if (($annotation -ne $null) -and ($annotation -ne "")) {$annotation = $annotation.trim()}
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                     $annotation = "Annotation: $annotation"
                     $selection.InsertBreak(6) # Line break.
                     $selection.typeText($annotation)
                  } # if (($Text -ne $null) -and ($text -ne "") -and ($text -ne "&nbsp;"))
                  $selection.typeParagraph() 
                } #"Word" 
                "HTML" {
                   $Text = $Name + ": [$DataLink]"
                   if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                      $Text = $Text + "<br>Annotation: $annotation"
                   } # if (($annotation -ne $null) -and ($annotation -ne ""))
                   If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                   $Text = $SnippetCSS.replace("*",$Text)
                   [System.IO.File]::AppendAllText($outputfile,$Text)
                   if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
                   [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
               } # "HTML"
            } # Switch ($Format)
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # "Video"

      default {
         Try {
             $Text = $PassedSnippet.contents

             $TagPreface = '<span class="RWSnippet">'
             if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
                if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
                $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
                $Text = $Text.insert($InsertPoint,$LabelPrefix)
             } # if ($PassedSnippet.Label -ne $null)

             if (($Text -ne $null) -and ($Text -ne "") -and ($Text -ne "&nbsp;")) {$Text = "$Text<BR>"}
             $Text = "$Text[Unknown Snippet Type]"
             
             $annotation = $PassedSnippet.annotation
             $annotation = ParseSnippetText $annotation $Log
             if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
                $Text = $Text + "<br>Annotation: $annotation"
             } # if (($annotation -ne $null) -and ($annotation -ne ""))

             If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
             $Text = $SnippetCSS.replace("*",$Text)
             if ($SeparateSnippets) {$Text = $Text + "<hr>"}
             [System.IO.File]::AppendAllText($outputfile,$Text)

             $Warning = "`r`n[WARNING]`r`n"
             $Warning = $Warning + "   Unknown snippet type detected.`r`n"
             $Warning = $Warning + "   Topic: $TopicName`r`n"
             $Warning = $Warning + "   Section: $SectionName`r`n"
             $Warning = $Warning + "[/WARNING]`r`n`r`n"
         } Catch {
           ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
         } # Try
      } # default
   } # switch ($PassedSnippet.type)
} # Function ParseSnippet

Function ParseSnippetText ($PassedText,$Log) {
   [array]$Text = @()
   if ($PassedText -ne $null) {
      $Startpos = $PassedText.IndexOf("<p ")
      While ($Startpos -ge 0) {
         $Nextpos = $PassedText.IndexOf("<p ",$Startpos + 1)
         if ($Nextpos -lt 0) {$Endpos = $PassedText.Length} else {$Endpos = $Nextpos}
         $Length = $Endpos - $Startpos
         $paragraph = $PassedText.substring($Startpos,$Length)
         $paragraph = $paragraph.replace("<br>","`r`n")
         $paragraph = $paragraph -replace "<.*?>"
         $paragraph = $paragraph.replace("&nbsp;"," ")
         $paragraph = $paragraph.replace("&lt;","<")
         $paragraph = $paragraph.replace("&gt;",">")
         [array]$Text = [array]$Text + $paragraph #.substring($TextStart,$TextLength)
         $Startpos = $Nextpos
      } # While ($Startpos -ge 0)
   } # if ($PassedText -ne $null)
   Return $Text
} # Function ParseSnippetText

Function ParseTags ($PassedTags,$Log) {
   # This block of code assumes that tags of the same domain will always be grouped together.
   [array]$taglist = @()
   [array]$domainlist = @()
   $tagline = ""

   foreach ($domain in $PassedTags.tag_assign.domain_name) {
      if (-not [array]$domainlist.Contains($domain)) {
         [array]$domainlist += $domain
      } # if (-not [array]$domainlist.Contains($domain))
   } # foreach ($domain in $PassedTags.tag_assign.domain_name)

   foreach ($domain in $domainlist) {
      foreach ($tag in $PassedTags.tag_assign) {
         if ($tag.domain_name -eq $domain) {
            if ($tagline -eq "") {$tagline = $tag.tag_name} else {$tagline = $tagline + ", " + $tag.tag_name}
         } # if ($tag.domain_name -eq $domain)
      } # foreach ($tag in $PassedTags.tag_assign)
      $tagline = $domain + ": " + $tagline
      [array]$taglist = [array]$taglist + $tagline
      $tagline = ""
   } # foreach ($domain in $domainlist)

   Return $taglist
} # Function ParseTags

Function ParseLinkage ($PassedTopic,$OutputFile,$SplitTopics,$Log,$Format) {
   $Linkage = $PassedTopic.linkage
   $Linkage = $Linkage | Sort-Object target_name
   $InboundList = ""
   $OutboundList = ""

   Switch ($Format) {
      "Word" {
           foreach ($link in $linkage) {
              if ($Link.direction -eq "Inbound") {
                 if ($InboundList -eq "") {$InboundList = 'Linked from: ' + $link.target_name} else {$InboundList = $InboundList + ', ' + $link.target_name}
              } elseif ($Link.direction -eq "Outbound") {
                 if ($OutboundList -eq "") {$OutboundList = 'Links to: ' + $link.target_name} else {$OutboundList = $OutboundList + ', ' + $link.target_name}
              }
           } # foreach ($link in $linkage)

           [array]$Linklist = @()
           if (($OutboundList) -and ($InboundList)) {
              [array]$Linklist += $OutboundList
              [array]$LInklist += $InboundList
           } elseif ($OutboundList) {
              [array]$Linklist += $OutboundList
           } elseif ($InboundList) {
              [array]$Linklist += $InboundList
           } else {
              [array]$Linklist = $null
           } # if (($OutboundList) -and ($InboundList))
      } # "Word"
      "HTML" {
           if ($SplitTopics) {
              foreach ($link in $linkage) {
                 $linkname = $link.target_name
                 $linkid = $link.target_id
                 $linkreference = $linkname + "_" + $linkid + ".html"
                 $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
                 for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
                    $linkrefernce = $linkreference.replace($invalidChars[$CheckChar],"_")
                 } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
                 if ($Link.direction -eq "Inbound") {
                    if ($InboundList -eq "") {$InboundList = 'Linked from: <a href="' + $linkreference + '">' + $link.target_name + '</a>'} else {$InboundList = $InboundList + ', <a href="' + $linkreference + '">' + $link.target_name + '</a>'}
                 } elseif ($Link.direction -eq "Outbound") {
                    if ($OutboundList -eq "") {$OutboundList = 'Links to: <a href="' + $linkreference + '">' + $link.target_name + '</a>'} else {$OutboundList = $OutboundList + ', <a href="' + $linkreference + '">' + $link.target_name + '</a>'}
                 }
              } # foreach ($link in $linkage)
           } else {
              foreach ($link in $linkage) {
                 if ($Link.direction -eq "Inbound") {
                    if ($InboundList -eq "") {$InboundList = 'Linked from: <a href="#' + $link.target_id + '">' + $link.target_name + '</a>'} else {$InboundList = $InboundList + ', <a href="#' + $link.target_id + '">' + $link.target_name + '</a>'}
                 } elseif ($Link.direction -eq "Outbound") {
                    if ($OutboundList -eq "") {$OutboundList = 'Links to: <a href="#' + $link.target_id + '">' + $link.target_name + '</a>'} else {$OutboundList = $OutboundList + ', <a href="#' + $link.target_id + '">' + $link.target_name + '</a>'}
                 }
              } # foreach ($link in $linkage)
           } # if ($SplitTopics)
           if (($OutboundList) -and ($InboundList)) {
              $Linklist = "$OutboundList<BR>$InboundList"
           } elseif ($OutboundList) {
              $Linklist = $OutboundList
           } elseif ($InboundList) {
              $Linklist = $InboundList
           } else {
              $Linklist = $null
           } # if (($OutboundList) -and ($InboundList))
      } # "HTML"
   } # Switch ($Format)
   Return $Linklist
} # Function ParseLinkage

Function ParseAliases ($PassedTopic,$OutputFile,$Log) {
   $Aliases = $PassedTopic.alias
   $AliasList = ""
   foreach ($Alias in $Aliases) {
      if ($AliasList -eq "") {$AliasList = "Aliases: " + $Alias.name} else {$AliasList = $AliasList + ", " + $Alias.name}
   } # foreach ($link in $linkage)
   Return $Aliaslist
} # Function ParseLinkage


Function AddIndent ($CSSCode,$Indcrement,$Log) {
   $FirstString = 'style="margin-left:'
   $LastString = 'px;"'
   $FirstIndex = $CSSCode.IndexOf($FirstString)
   $LastIndex = $CSSCode.IndexOf($LastString,$FirstIndex)

   $FirstMargin = $FirstIndex + $FirstString.Length
   $Length = $LastIndex - $FirstMargin

   $CurrentMargin = $CSSCode.substring($FirstMargin,$Length)
   $NewMargin = $CurrentMargin.ToInt32($null) + $Indcrement
   $NewMargin = $NewMargin.ToString()
   $NewMargin = $FirstString + $NewMargin + $LastString

   $LastIndex = $LastIndex + $LastString.Length
   $Length = $LastIndex-$FirstIndex
   $OldMargin = $CSSCode.substring($FirstIndex,$Length)

   $NewCSS = $CSSCode.replace($OldMargin,$NewMargin)

   Return $NewCSS

} # Function AddIndent


Function GetTagLocations ($PassedSnippet) {
   $taglist = $null
   # Get the locations of all the start tags.
   $ul = $PassedSnippet.IndexOf("<ul")
   $ol = $PassedSnippet.IndexOf("<ol")
   $table = $PassedSnippet.IndexOf("<table")

   While (($ul -ge 0) -or ($ol -ge 0) -or ($table -ge 0)) {
      if ($ul -ge 0) {
         $properties = @{
            'Tag'='ul';
            'Start'=$ul;
            'End'=0
         } # $properties

         $TagLocation = New-Object –TypeName PSObject –Prop $properties
         [array]$Taglist = $Taglist + $TagLocation
      } # if ($ul -ge 0)

      if ($ol -ge 0) {
         $properties = @{
            'Tag'='ol';
            'Start'=$ol;
            'End'=0
         } # $properties

         $TagLocation = New-Object –TypeName PSObject –Prop $properties
         [array]$Taglist = $Taglist + $TagLocation
      } # if ($ol -ge 0)

      if ($table -ge 0) {
         $properties = @{
            'Tag'='table';
            'Start'=$table;
            'End'=0
         } # $properties

         $TagLocation = New-Object –TypeName PSObject –Prop $properties
         [array]$Taglist = $Taglist + $TagLocation
      } # if ($table -ge 0)

      $Taglist = $Taglist | Sort-Object Start
      $HighTag = $Taglist.Count-1
      $StartPosition = $Taglist[$HighTag].Start + 1

      $ul = $PassedSnippet.IndexOf("<ul",$StartPosition)
      $ol = $PassedSnippet.IndexOf("<ol",$StartPosition)
      $table = $PassedSnippet.IndexOf("<table",$StartPosition)
   } # While (($ul -ge 0) -or ($ol -ge 0) -or ($table -ge 0))

   # Get all the end tags.
   foreach ($tag in $Taglist) {
      $endtag = "</" + $tag.tag
      $endtag = $PassedSnippet.IndexOf($endtag,$tag.Start + 1)
      $Tag.end = $endtag+4
   } # foreach ($tag in $Taglist)

   # Fill in the gaps, if there are any.
   $firstitem = 0
   $firstpos = 0
   $lastitem = $Taglist.count -1
   for ($item=$firstitem; $item -le $lastitem; $item++) {
      if (($firstpos -lt $Taglist[$item].start)) {
         $endpos = $Taglist[$item].start - 1

         $properties = @{
            'Tag'='text';
            'Start'=$firstpos;
            'End'=$endpos
         } # $properties

         $TagLocation = New-Object –TypeName PSObject –Prop $properties
         [array]$gaplist = $gaplist + $TagLocation

      } # if (($firstpos -lt $Taglist[$item].start))

      if ($item -eq $lastitem) {
         $firstpos = $Taglist[$item].end + 1
         $endpos = $PassedSnippet.length - 1

         $properties = @{
            'Tag'='text';
            'Start'=$firstpos;
            'End'=$endpos
         } # $properties

         $TagLocation = New-Object –TypeName PSObject –Prop $properties
         [array]$gaplist = $gaplist + $TagLocation
      } # if ($item -eq $lastitem)
      $firstpos = $taglist[$item].end+1
   } # for ($item=$firstitem; $item -le $lastitem; $item++)

   $Taglist = $Taglist + $gaplist
   $Taglist = $Taglist | Sort-Object start
   Return $Taglist
} # GetTagLocations


Function ParseFormattedStuff ($PassedSnippet,$Format,$wordfile) {
   $TagLocations = GetTagLocations $PassedSnippet

   foreach ($tag in $TagLocations) {
      $tagtext = $PassedSnippet.substring($tag.start,$tag.end-$tag.start+1)
      if ($tagtext.contains("<li ")) {$tagtext = $tagtext.replace("<li ","<p ><li ")}
      [array]$Text = ParseSnippetText $tagText
      Switch ($tag.tag) {
         "text" {
            if ($Text) {
               if ($Text.contains("&#xd;")) {$Text = $Text.Replace("&#xd;","")}
               if ($Text.contains("&nbsp;")) {$Text = $Text.Replace("&nbsp;"," ")}
               Switch ($Format) {
               "Word" {
                  foreach ($paragraph in $Text) {
                     $selection = $wordfile.selection
                     $selection.Style = "Normal"
                     $selection.typeText($paragraph)
                     $Selection.typeParagraph()
                  } # foreach ($paragraph in $Text)
               } # "Word"
               "HTML" {
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
            }
         } # "text"
            
         "ul" {
            Switch ($Format) {
               "Word" {
                  foreach ($paragraph in $Text) {
                     $selection = $wordfile.selection
                     $selection.Style = "Normal"
                     $selection.typeText($paragraph)
                     # $selection.Range.ListFormat.ApplyBulletDefault()
                     $selection.Range.ListFormat.ApplyBulletDefault()
                     $Selection.typeParagraph()
                  } # foreach ($paragraph in $Text)
               } # "Word"
               "HTML" {
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } # "ul"

         "ol" {
            Switch ($Format) {
               "Word" {
                  foreach ($paragraph in $Text) {
                     $selection = $wordfile.selection
                     $selection.Style = "Normal"
                     $selection.typeText($paragraph)
                     # $selection.Range.ListFormat.ApplyBulletDefault()
                     $selection.Range.ListFormat.ApplyNumberDefault()
                     $Selection.typeParagraph()
                  } # foreach ($paragraph in $Text)
               } # "Word"
               "HTML" {
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)

         } # "ol"

         "table" {
            Switch ($Format) {
               "Word" {
                  $tabledata = ParseTable ($tagtext)
                  $selection = $wordfile.selection
                  $Table = $Selection.Tables.add(
                     $Selection.Range,($Tabledata.Count),($Tabledata[0].Count),
                     [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdDefaultTableBehavior,
                     [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                  )

                  for ($row = 0;$row -le $tabledata.count-1;$row++) {
                     for ($column = 0;$column -le $tabledata[0].count-1;$column++) {
                        $Table.cell($row+1,$column+1).range.text = $tabledata[$row][$column]
                     } # for ($column = 0;$column++; $column -le $table[0].count)
                  } # for ($row = 0;$row++;$row -le $table.count)
                  $selection.Start= $wordfile.ActiveDocument.Content.End
                  $selection.TypeParagraph()
               } # "Word"
               "HTML" {
                  If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
                  if ($SeparateSnippets) {$Text = $Text + "<hr>"}
                  [System.IO.File]::AppendAllText($outputfile,$Text)
               } # "HTML"
            } # Switch ($Format)
         } # "table"
      } # Switch ($tag.tag)
   } # foreach ($tag in $TagLocations)
   # Add-Content -Path $outputfile -Value ""
} # Function ParseFormattedStuff

Function ParseTable ($PassedText) {
   if ($PassedText -ne $null) {
      $opentr = $PassedText.Indexof('<tr')
   } else {
      $opentr = -1
   }
   [array]$RowList = $null
   While ($opentr -ge 0) {
      $closetr = $PassedText.Indexof("</tr",$opentr)
      $length = $closetr-$opentr
      if ($length -gt 1) {
         $Row = $PassedText.substring($opentr+1,$length-1)
         $Row = $Row.Replace("tr>&#xd;","")
         $Row = $Row.Trim()
         [array]$RowList = [array]$RowList + $Row
      }

      if ($closetr -lt $opentr) {
         $opentr = -1
      } else {
         $PassedText = $PassedText.substring($closetr,$PassedText.length-$closetr)
         $opentr = $PassedText.Indexof('<tr')
      }
   } # While ($opentr -ge 0)

   $Table = new-object system.collections.arraylist
   $NewRow = @()
   foreach ($Row in $RowList) {
      if ($Row -ne $null) {
         $opentd = $Row.Indexof('<td')
      } else {
         $opentd = -1
      }
      $ItemList = new-object psobject
      $ItemNumber = 1
      While ($opentd -ge 0) {
         $closetd = $Row.Indexof("</td",$opentd)
         $length = $closetd-$opentd
         if ($length -gt 1) {
            $Item = $Row.substring($opentd+1,$length-1)
            $Item = ParseSnippetText $Item
            # $ItemList | Add-Member -Name "Item$ItemNumber" -Type noteproperty -value $Item
            [array]$NewRow = [array]$NewRow + $Item
         }

         if ($closetd -lt $opentd) {
            $opentd = -1
         } else {
            $Row = $Row.substring($closetd,$Row.length-$closetd)
            $opentd = $Row.Indexof('<td')
         }
         $ItemNumber++
      } # While ($opentd -ge 0)
      # [array]$Table = [array]$Table + $ItemList

      $AddToArrayList = '$Table.Add(('
      foreach ($cell in $NewRow) {
         $AddToArrayList = $AddToArrayList + '"' + $Cell + '",'
      }
      $NewRow = $null
      $AddToArrayList = $AddToArrayList + ')) > $null'
      $AddToArrayList = $AddToArrayList.replace(",))","))")
      Invoke-Expression $AddToArrayList
      $AddtoArrayList = $null
   } # foreach ($Row in $RowList)
   Return $Table
} # Function ParseTable

Function StripStyles ($Text,$Log) {
   <#
   $Standards = @('"font-family:',
                  ';font-family:',
                  '"background-color:',
                  ';background-color:',
                  '"background:',
                  ';background:',
                  '"color:',
                  ';color:'
                  '"font-size:',
                  ';font-size:')
   
   foreach ($Standard in $Standards) {
      While ($Text.contains($Standard)) {
         $SubStandard = $Standard.substring(1,$Standard.Length-1) 
         $Start = $Text.IndexOf($SubStandard)
         $EndSemi = $Text.IndexOf(";",$Start)
         $EndQuote = $Text.IndexOf('"',$Start)
         If ($EndSemi -lt $EndQuote) {$End = $EndSemi} elseif ($EndQuote -lt $EndSemi) {$End = $EndQuote-1} 
         $Length = $End-$Start+1
         if ($Length -ge 1) {
            $Substring = $Text.Substring($Start,$End-$Start+1)
            if (-not $Substring.contains("wingdIngs")) {$Text = $Text.Replace($Substring,"")}
         } # if ($Length -ge 1)
      } # While ($Text.contains($Standard))
   } # foreach ($Standard in $Standards)
   #>
   Return $Text
} # Function StripStyles

# Just leaving this bit here for later reference as I implement logging
# $ErrorActionPreference = 'Stop'
#
# Try {
#    
# } Catch {
#     $ErrorMessage = $_.Exception.Message
#     $FailedItem = $_.Exception.ItemName
# ------------------------
# $e = $_.Exception
  #       $line = $_.InvocationInfo.ScriptLineNumber
    #     $msg = $e.Message 
# } Finally {
#     
#     $Time=Get-Date
#     "This script made a read attempt at $Time" | out-file c:\logs\ExpensesScript.log -append
# } # Try

Function ReportException ($CommandLine,$ErrorMsg,$ErrorObject,$Log,$TopicName,$SectionName,$SnippetType) {

   $CustomError = "`r`n[ERROR]`r`n"
   $CUstomError = $CustomError + "     Command Line: $CommandLine`r`n"
   $CUstomError = $CustomError + "         Reason: " + $ErrorObject.CategoryInfo.Category + " - " + $ErrorObject.ToString() +"`r`n`r`n"
   $CUstomError = $CustomError + "       Location: Line " + $ErrorObject.InvocationInfo.ScriptLineNumber + ", Character " + $ErrorObject.InvocationInfo.OffsetInLine + "`r`n"
   $CUstomError = $CustomError + "    Script Line: " + $ErrorObject.InvocationInfo.Line.Trim() + "`r`n"
   if (($Topic -ne $null) -and ($Topic -ne "") -and ($Topic -ne "&nbsp;")) {$CustomError = $CustomError + "`r`n          Topic: $TopicName`r`n"}
   if (($Section -ne $null) -and ($Section -ne "") -and ($Section -ne "&nbsp;")) {$CustomError = $CustomError + "        Section: $SectionName`r`n"}
   if (($SnippetType -ne $null) -and ($SnippetType -ne "") -and ($SnippetType -ne "&nbsp;")) {$CustomError = $CustomError + "   Snippet Type: $SnippetType`r`n"}
   $CUstomError = $CustomError + "[/ERROR]"

   Write-Host $CustomError
   if ($Log) {
      [System.IO.File]::AppendAllText($Log,$CustomError)
      $EndTime = Get-Date
      $Date = $EndTime.ToString()
      $RunTime = $EndTime - $StartTime
      $RunTime = $RunTime.ToString()
      [System.IO.File]::AppendAllText($Log,"`r`n`r`nAborted: $Date`r`n")
      [System.IO.File]::AppendAllText($Log,"Run Time: $RunTime")
   }
   Exit
} # Function FileException

Function ComposeDetails ($ExportedDetails,$RealmTitleCSS,$RealmDetailsCSS,$Format,$wordfile) {
   $RealmDetails = $null

   $CoverArt = $ExportedDetails.cover_art
   if (($CoverArt -ne $null) -and ($CoverArt -ne "")) {
         Switch ($Format) {
            "Word" {
               $selection = $wordfile.selection
               $ImageStream=[System.IO.MemoryStream][System.Convert]::FromBase64String($CoverArt)
               $ImageBmp=[System.Drawing.Bitmap][System.Drawing.Image]::FromStream($ImageStream)
               [Windows.Forms.Clipboard]::SetImage($ImageBmp)
               $selection.Paste()
               $Selection.InsertBreak(7) # Page
               [Windows.Forms.Clipboard]::Clear()
            }
            "HTML" {
               $CoverArtScale = 'width="100%"'
               $CoverArt = '<img ' + $CoverArtSCale + ' src="data:image/png;base64,' + $CoverArt + '">'
           }
        }
      $RealmDetails = $CoverArt
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $RealmTitle = $ExportedDetails.name
   if (($RealmTitle -ne $null) -and ($RealmTitle -ne "")) {
      Switch ($Format) {
         "Word" {
            $RealmTitle = $RealmTitle.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.Style = "Title"
            $selection.typeText("$RealmTitle")
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $RealmTitle = $RealmTitleCSS.replace("*",$RealmTitle)
            $RealmDetails = $RealmDetails + $RealmTitle
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $Version = $ExportedDetails.version
   if (($Version -ne $null) -and ($Version -ne "")) {
      $Version = "Version: $Version"
      Switch ($Format) {
         "Word" {
            $Version = $Version.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText($Version)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Version = $RealmDetailsCSS.replace("*",$Version)
            $RealmDetails = $RealmDetails + $Version
         } # "HTML"
      } # Switch ($Format)
   } # (($Version -ne $null) -and ($Version -ne ""))

   $Summary = $ExportedDetails.summary
   if (($Summary -ne $null) -and ($Summary -ne "")) {
      Switch ($Format) {
         "Word" {
            $Summary = $Summary.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Summary:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Summary)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Summary = "Summary:<br>$Summary"
            $Summary = $RealmDetailsCSS.replace("*",$Summary)
            $RealmDetails = $RealmDetails + $Summary
         } # HTML
      } # Switch ($Format)
   } # if (($Summary -ne $null) -and ($Summary -ne ""))

   $Description = $ExportedDetails.description
   if (($Description -ne $null) -and ($Description -ne "")) {
      Switch ($Format) {
         "Word" {
            $Description = $Description.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Description:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Description)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Description = "Description:<br>$Description"
            $Description = $RealmDetailsCSS.replace("*",$Description)
            $RealmDetails = $RealmDetails + $Description
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $Requirements = $ExportedDetails.requirements
   if (($Requirements -ne $null) -and ($Requirements -ne "")) {
      Switch ($Format) {
         "Word" {
            $Requirements = $Requirements.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Requirements:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Requirements)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Requirements = "Requirements:<br>$Requirements"
            $Requirements = $RealmDetailsCSS.replace("*",$Requirements)
            $RealmDetails = $RealmDetails + $Requirements
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $Credits = $ExportedDetails.credits
   if (($Credits -ne $null) -and ($Credits -ne "")) {
      Switch ($Format) {
         "Word" {
            $Credits = $Credits.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Credits:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Credits)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Credits = "Credits:<br>$Credits"
            $Credits = $RealmDetailsCSS.replace("*",$Credits)
            $RealmDetails = $RealmDetails + $Credits
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $Legal = $ExportedDetails.legal
   if (($Legal -ne $null) -and ($Legal -ne "")) {
      Switch ($Format) {
         "Word" {
            $Legal = $Legal.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Legal:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Legal)
            $Selection.typeParagraph()
         } # "Word"
         "HTML" {
            $Legal = "Legal:<br>$Legal"
            $Legal = $RealmDetailsCSS.replace("*",$Legal)
            $RealmDetails = $RealmDetails + $Legal
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))

   $Notes = $ExportedDetails.other_notes
   if (($Notes -ne $null) -and ($Notes -ne "")) {
      Switch ($Format) {
         "Word" {
            $Notes = $Notes.replace("&#xd;","")
            $selection = $wordfile.selection
            $selection.typeText("Notes:")
            $selection.InsertBreak(6) # Line break.
            $selection.typeText($Notes)
         } # "Word"
         "HTML" {
            $Notes = "Additional Notes:<br>$Notes"
            $Notes = $RealmDetailsCSS.replace("*",$Notes)
            $RealmDetails = $RealmDetails + $Notes
         } # "HTML"
      } # Switch ($Format)
   } # if (($RealmTitle -ne $null) -and ($RealmTitle -ne ""))
   if ($RealmDetails) {Return $RealmDetails}
} # Function ComposeDetails

Function InitializeWordFile() {
   [ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
   $wordfile = New-Object -ComObject word.application
   $wordfile.visible = $true
   $doc = $wordfile.documents.add()
   
   $GMDirections = $doc.Styles.Add("GM_Directions")  
   $GMDirections.BaseStyle = $doc.Styles("Normal") 
   $GMDirections.shading.BackgroundPatternColor = 10079487
   $GMDirections.QuickStyle = $true

   $TopicDetails = $doc.Styles.Add("Topic_Details")  
   $TopicDetails.BaseStyle = $doc.Styles("Normal") 
   $TopicDetails.shading.BackgroundPatternColor = 16764057
   $TopicDetails.QuickStyle = $true
   
   Return $wordfile
} # Function InitializeWordDoc()

Function WriteHTMLHeader($Destination,$CSSFileName,$CommandLine) {
   Try {
      [System.IO.File]::WriteAllText($Destination,"<!DOCTYPE html>")
      [System.IO.File]::AppendAllText($Destination,"<html>")
      [System.IO.File]::AppendAllText($Destination,"<head>")
      [System.IO.File]::AppendAllText($Destination,"<title>$Title</title>")
      [System.IO.File]::AppendAllText($Destination,'<link rel="stylesheet" type="text/css" href="' + $CSSFilename + '">')

      # Just in case this file gets posted on a public web site, tell googlebot not to index this page.
      [System.IO.File]::AppendAllText($Destination,'<meta name="googlebot" content="noindex">')

      # Continue priming the HTML output file.
      [System.IO.File]::AppendAllText($Destination,"</head>")
      [System.IO.File]::AppendAllText($Destination,"<body>")
      # [System.IO.File]::AppendAllText($Destination,$TitleCSS.Replace("*",$Title)
   } Catch {
        ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
   } # Try
} # Function WriteHTMLHeader

Function WriteHTMLFooter($Destination,$CommandLine) {
   Try {
      # Close out the HTML file
      [System.IO.File]::AppendAllText($Destination,"</body>")
      [System.IO.File]::AppendAllText($Destination,"</html>")
   } Catch {
        ReportException $CommandLine $Error $_ $Log $Log $TopicName $SectionName
   } # Try
} # Function WriteHTMLFooter

Function CloseWordFile ($Destination,$wordfile) {
   $wordfile.ActiveDocument.saveas([ref] $Destination, [ref]$saveFormat::wdFormatDocument)
   $wordfile.Quit()
   [Runtime.InteropServices.Marshal]::ReleaseComObject($wordfile) > $null
   [GC]::Collect()
   [GC]::WaitForPendingFinalizers()
} # CloseWordFile ($wordfile)

# Main {
   # Set all errors as terminating errors to facilitate error trapping.
   $ErrorActionPreference = "Stop"
   $Validator = "Script"
   $Script = $PSCmdlet.MyInvocation.PSScriptRoot + "\" + $PSCmdlet.MyInvocation.MyCommand

   [void][reflection.assembly]::loadwithpartialname("System.Drawing")
   [void][reflection.assembly]::loadwithpartialname("System.Windows.Forms")
   if ($Log) {$CreateLog = $true} else {$CreateLog = $false}
   
   if ($Force) {
      Write-Host "Running script with no input validation."
   } else {
      if (Test-Path ".\ValidateInput.ps1") {
         Write-Host "Validating input." 
         $ValidInput = &(".\ValidateInput") $Script $Source $Destination $Sort $SimpleImageScale $SmartImageScale $ExtractFiles $CSSFileName $SplitTopics $Format $CreateLog $Log "Script"
         If ($ValidInput) {
            Write-Host "Input validated. Running Script."
         } else {
            Write-Host "Please verify the command line options and try again."
            exit
         } # If ($ValidInput)
      } else {
         $MsgTitle ="Can't Validate"
         $MsgText = "Cannot find input validation script. Proceed anyway?"
         $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Proceeds without validation."
         $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exits."
         $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
         $result = $host.ui.PromptForChoice($MsgTitle, $MsgText, $options, 0) 
         switch ($result) {
            0 {Write-Host "You have chosen to continue with no input validation."}
            1 {exit}
         } # switch ($result)

         Write-Host "Input validation script not found. Proceeding anyway..."
      } # if (Test-Path ".\ValidateInput.ps1")
   } # if ($Force) {

   # Get Date and CommandLine for logging purposes.
   $StartTime = Get-Date
   $CommandLine = $PSCmdlet.MyInvocation.Line

   # If the -Log option was invoked, log the current date and the full command line.
   # If anything goes wrong with this, write the error to the console and exit.
   if ($Log) {
      Try {
        [System.IO.File]::WriteAllText($Log,"RWExport-To-HTML.ps1`r`n")
        [System.IO.File]::AppendAllText($Log,"by EightBitz`r`n")
        [System.IO.File]::AppendAllText($Log,"Version 1.6`r`n")
        [System.IO.File]::AppendAllText($Log,"2017-08-14, 9:00 PM CDT`r`n`r`n")
        $Date = $StartTime.ToString() + "`r`n`r`n"
        $CommandLine = $CommandLine + "`r`n"
        [System.IO.File]::AppendAllText($Log,"Started: $Date")
        [System.IO.File]::AppendAllText($Log,"Command line:`r`n")
        [System.IO.File]::AppendAllText($Log,$CommandLine)
      } Catch {
        ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
     } # Try
   } # if ($Log)


   # Setup CSS tags so all we have to do later is a string.replace of "*"
   #    with whatever text we wish to use. Also, the variable names will
   #    remind us which tags are for which elements.
   $InitialIndent = 'style="margin-left:5px;"'
   $RealmTitleCSS = "<h1>*</h1>"
   $RealmDetailsCSS = "<h2>*</h2>"
   $TopicCSS = '<h3 id="" ' + $InitialIndent + '>*</h3>'
   $TopicDetailsCSS = '<h4 ' + $InitialIndent + '>*</h4>'
   $SectionCSS = '<h5 ' + $InitialIndent + '>*</h5>'
   $SnippetCSS = '<p>*</p>'
   $Indcrement = 50

   # Import data from the specified source file.
   # If anything goes wrong with this, write the error to the console and exit.
   Try {
      # [xml]$RWExportData = Get-Content -Path $Source

      $RWExportData = New-Object Xml
      $RWExportData.Load((Convert-Path $Source))
   } Catch {
        ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
   } # Try

   if (($SplitTopics) -or ($ExtractFiles)) {
      $TopicRoot = Get-Item $Destination
      if ($TopicRoot) {$TestPath = Test-Path $TopicRoot}
      if ($TestPath) {$TestFolder = Test-Path $TopicRoot -PathType Container}

      if ((Get-ChildItem $TopicRoot -force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
         $EmptyPath = $true
      } else {
         $EmptyPath = $false
      } # if((Get-ChildItem $TopicPath -force | Select-Object -First 1 | Measure-Object).Count -eq 0)

      if (-not $TopicRoot.ToString().endswith("\")) {$TopicRoot = $TopicRoot.ToString() + "\"} else {$TopicRoot = $TopicRoot.ToString()}
      $Destination = $TopicRoot
    } # if (($SplitTopics) -or ($ExtractFiles))

    if ($ExtractFiles) {
       $RealmTitle = $RWExportData.output.definition.details.name
       $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
       for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
         Switch ($Format) {
            "Word" {$TopicFile = $RealmTitle.replace($invalidChars[$CheckChar],"_") + ".doc"}
            "HTML" {$TopicFile = $RealmTitle.replace($invalidChars[$CheckChar],"_") + ".html"}
         } # Switch ($Format)
       } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
       if (-not $SplitTopics) {$Destination = $TopicRoot + $TopicFile}
       $ExportPath = $TopicRoot + "Realm_Files\"
    } # if ($ExtractFiles)

    if ($SplitTopics) {
       Switch ($Format) {
          "Word" {
             $RealmDetailsFile = $Destination + "RealmDetails.doc"
             $wordfile = InitializeWordFile
             $RealmDetails = ComposeDetails $RWExportData.Output.definition.details $RealmTitleCSS $RealmDetailsCSS $Format $wordfile
             CloseWordFile $RealmDetailsFile $wordfile
          } # "Word"
          "HTML" {
             $RealmDetailsFile = $Destination + "RealmDetails.html"
             WriteHTMLHeader $RealmDetailsFile $CSSFileName $CommandLine
      
             Try {
                # Get the basic realm info from the specified export file.
                $RealmDetails = ComposeDetails $RWExportData.Output.definition.details $RealmTitleCSS $RealmDetailsCSS
                if (($RealmDetails -ne $null) -and ($RealmDetails -ne "")) {
                   [System.IO.File]::AppendAllText($RealmDetailsFile,$RealmDetails)
                } # if (($RealmDetails -ne $null)
                WriteHTMLFooter $RealmDetailsFile $CommandLine
             } Catch {
                ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
             } # Try
          } # "HTML"
       } # Switch ($Format)
    } else {
      Switch ($Format) {
         "Word" {
            $wordfile = InitializeWordFile
            $RealmDetails = ComposeDetails $RWExportData.Output.definition.details $RealmTitleCSS $RealmDetailsCSS $Format $wordfile
            $selection = $wordfile.selection
            $Selection.InsertBreak(7) # Page
            $toc = $wordfile.ActiveDocument.TablesOfContents.Add($selection.range)
            $Selection.InsertBreak(7) # Page
         } # "Word"
         "HTML" {
            WriteHTMLHeader $Destination $CSSFileName $CommandLine
            Try {
               # Get the basic realm info from the specified export file.
               $RealmDetails = ComposeDetails $RWExportData.Output.definition.details $RealmTitleCSS $RealmDetailsCSS $Format $wordfile
               if (($RealmDetails -ne $null) -and ($RealmDetails -ne "")) {
                  [System.IO.File]::AppendAllText($Destination,$RealmDetails)
               } # if (($RealmDetails -ne $null) -and ($RealmDetails -ne ""))
            } Catch {
              ReportException $CommandLine $Error $_ $Log $TopicName $SectionName $PassedSnippet.type
            } # Try
         } # "HTML"
      } # Switch ($Format)
   } # if (($SplitTopics) -and ($Format -eq "HTML"))

   # Get the main content from the export file.
   $Contents = $RWExportData.output.contents

   # Get the topic list, and sort it according to the specified method.
   Switch ($Sort) {
      1 {$TopicList = $Contents.topic | Sort-Object public_name}
      2 {$TopicList = $Contents.topic | Sort-Object prefix,public_name}
      3 {$TopicList = $Contents.topic | Sort-Object category_name,public_name}
      4 {$TopicList = $Contents.topic | Sort-Object category_name,prefix,public_name}
      default {$TopicList = $Contents.topic}
   } # Switch ($Sort)


   $Parent = "/"
   # $selection = $wordfile.selection
   # $selection.style = "Heading 1"
   foreach ($Topic in $TopicList) {
      ParseTopic $Topic $Destination $Sort $Prefix $Suffix $Details $Indent $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $Parent $TitleCSS $TopicCSS $TopicDetailsCSS $SectionCSS $SnippetCSS $InitialIndent $Indcrement $KeepStyles $ExtractFiles $ExportPath $SplitTopics $Log $CSSFileName $CommandLine $Format $wordfile
   } # foreach ($Topic in $Contents.topic)
   
   if (-not $SplitTopics) {
      Switch ($Format) {
         "Word" {
            $wordfile.ActiveDocument.TablesOfContents.item(1).Update()
            CloseWordFile $Destination $wordfile
         } # "Word"
         "HTML" {
            WriteHTMLFooter $Destination $CommandLine
         } # "HTML"
      } # Switch ($Format)
   } # if (-not $SplitTopics)

   if ($log) {
      $EndTime = Get-Date
      $Date = $EndTime.ToString()
      $RunTime = $EndTime - $StartTime
      $RunTime = $RunTime.ToString()
      [System.IO.File]::AppendAllText($Log,"`r`nFinished: $Date`r`n")
      [System.IO.File]::AppendAllText($Log,"Run Time: $RunTime")
   }
# } Main