<# 
.SYNOPSIS
RWExport-To-HTML.ps1
by EightBitz
Version 0.8a
2017-01-24, 05:00 AM CST

RWExport-To-HTML.ps1 transforms a Realm Works export file into a formatted HTML file. The formatted HTML file relies on a "main.css" file for formatting information.

LICENSE:
This script is now licensed under the Creative Commons Attribution
   
Basically, that means:
   -You are free to share and adapt the script.
   -When sharing the script, you must give appropriate credit and indicate if changes were made.
   -You may not use the material for commercial purposes.

   Summary: https://creativecommons.org/licenses/by/4.0/
   Legal Code: https://creativecommons.org/licenses/by/4.0/legalcode

IMPORTANT:
You will likely have to change your execution policy to run this script. You can do that with the following command:

   Set-ExecutionPolicy RemoteSigned

You may have to be logged in as a local administrator to do this, or run at least use the "Run as Administrator" option when opening PowerShell (you can do this with a right-click).
Please do NOT set the execution policy to "Unrestricted". If you still have issues after setting it to "RemoteSigned", then you can open this script as a text file, select all, copy, and paste it into a new text file. Once you do that, delete the old file, and rename the new one to "RWExport-To-HTML.ps1".

IMPORTANT:
The export from Realm Works must done with the "Compact Output" option. This script will likely not work with a "Full Export".
You can still export your full realm if you like, but make sure you do so with the "Compact Output" option.

IMPORTANT:
Make sure that the HTML file and the main.css file are in the same directory, otherwise, the HTML file will have no formatting.
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
.PARAMETER CSSFileName
By default, the HTML output file looks for "main.css" to define its style. You can use this option to specify a different name.
If you have two realms called Realm1 and Realm2, and you want each HTML output to have different fonts, font sizes or colors, you can specify a name or "realma.css" for one file and "realmb.css" for the other.
Not that this option does not create the file. It just tells the HTML file which filename to look for.
For now, at least, you will have to manually copy main.css or some other css file you like and make whatever changes you wish.
.INPUTS
The .rwexport file from a Realm Works export that was created with the "Compact Output" option. This can an export of a full realm or a custom or partial export, just so long as it's made with the "Compact Output" option.
.OUTPUTS
An HTML file that uses an external style sheet for formatting. The external style sheet should be named "main.css" and should be in the same folder as the HTML output. The main.css file will not be generated by this script, but should already exist.
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
Display simple images inline at 50% of their full size.
.EXAMPLE
RWExport-To-HTML.ps1 -Source MyExport.rwoutput -Destination MyHTML.html -SmartImageScale 50
Display smart images inline at 50% of their full size.
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

As a side note, you can set up a separate folder for each export, and put a different main.css file in each folder, so you can give each export a completely different look (by editing the main.css file in each folder).



AND SPEAKING OF MAIN.CSS:

This script and its resulting HTML file assume the existence of a file named "main.css". If this file is not in the same folder as the resulting HTML file, it will not display correctly.

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

    # Specify a different name for the CSS file.
    # This does not create the CSS file. It only changes the name
    #    defined in the header of the HTML file.
    [Parameter()] 
    [string]$CSSFileName = "main.css",

    # Specify the path for an optional log file.
    #    [NOT WORKING YET]
    [Parameter()] 
    [string]$Log

) # param

Function ParseTopic($PassedTopic,$Outputfile,$Sort,$Prefix,$Suffix,$Details,$Indent,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$Parent,$TitleCSS,$TopicCSS,$TopicDetailsCSS,$SectionCSS,$SnippetCSS,$Indcrement,$KeepStyles,$Log) {
   $TopicName = $PassedTopic.public_name
   if ($Prefix -and $PassedTopic.Prefix) {$TopicName = $PassedTopic.Prefix + " - " + $TopicName}
   if ($Suffix -and $PassedTopic.Suffix) {$TopicName = $TopicName + " (" + $PassedTopic.Suffix + ")"}
  
   $ParentName = "Parent Topic: $Parent"
   $CategoryName = "Category: " + $PassedTopic.category_name.Trim()
   $TopicDetails = "$ParentName<br>$CategoryName"

   If ($PassedTopic.tag_assign) {
      $tagline = ParseTags $PassedTopic $Outputfile $Log
      $TopicDetails = "$TopicDetails<br>$tagline"
   } # If ($PassedTopic.tag_assign)

   If ($PassedTopic.linkage) {
      $linkage = ParseLinkage $PassedTopic $outputfile $Log
      $TopicDetails = "$TopicDetails<br>$linkage"
   } # If ($PassedTopic.linkage)

   $TopicName = $TopicCSS.Replace("*",$TopicName)
   $TopicDetails = $TopicDetailsCSS.Replace("*",$TopicDetails)

   [System.IO.File]::AppendAllText($outputfile,$TopicName)
   If ($Details) {[System.IO.File]::AppendAllText($outputfile,$TopicDetails)}

   foreach ($Section in $PassedTopic.section) {
      ParseSection $Section $Outputfile $SectionCSS $SnippetCSS $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $Log
   } # foreach ($Section in $PassedTopic.section)

   Switch ($Sort) {
      1 {$TopicList = $PassedTopic.topic | Sort-Object public_name}
      2 {$TopicList = $PassedTopic.topic | Sort-Object prefix,public_name}
      3 {$TopicList = $PassedTopic.topic | Sort-Object category_name,public_name}
      4 {$TopicList = $PassedTopic.topic | Sort-Object category_name,prefix,public_name}
      default {$TopicList = $PassedTopic.topic}
   }

   $Parent = $Parent + $PassedTopic.Public_Name + "/"
   foreach ($Topic in $TopicList) {
      If ($Indent) {
         $SubTopicCSS = AddIndent $TopicCSS $Indcrement $Log
         $SubTopicDetailsCSS = AddIndent $TopicDetailsCSS $Indcrement $Log
         $SubSectionCSS = AddIndent $SectionCSS $Indcrement $Log
         $SubSnippetCSS = AddIndent $SnippetCSS $Indcrement $Log
      } else {
         $SubTopicCSS = $TopicCSS
         $SubTopicDetailsCSS = $TopicDetailsCSS
         $SubSectionCSS = $SectionCSS
         $SubSnippetCSS = $SnippetCSS
      } # If ($Indent)
      ParseTopic $Topic $Outputfile $Sort $Prefix $Suffix $Details $Indent $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $Parent $TitleCSS $SubTopicCSS $SubTopicDetailsCSS $SubSectionCSS $SubSnippetCSS $Indcrement $KeepStyles $Log
   } # foreach ($Topic in $TopicList)
} # Function ParseTopic($PassedTopic)

Function ParseSection ($PassedSection,$Outputfile,$SectionCSS,$SnippetCSS,$Indcrement,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$KeepStyles,$Log) {
   $SectionName = $SectionCSS.Replace("*",$PassedSection.name)

   [System.IO.File]::AppendAllText($outputfile,$sectionname)
   # Add-Content -Path $outputfile -Value $SectionName

   foreach ($Snippet in $PassedSection.snippet) {
      ParseSnippet $Snippet $Outputfile $SnippetCSS $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $Log
   } # foreach ($Snippet in $PassedSection.snippet)

   foreach ($Section in $PassedSection.section) {
      If ($Indent) {
         $SubSectionCSS = AddIndent $SectionCSS $Indcrement $Log
         $SubSnippetCSS = AddIndent $SnippetCSS $Indcrement $Log
      } else {
         $SubSectionCSS = $SectionCSS
         $SubSnippetCSS = $SnippetCSS
      } # If ($Indent)
      ParseSection $Section $Outputfile $SubSectionCSS $SubSnippetCSS $Indcrement $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $KeepStyles $Log
   } # foreach ($Section in $PassedSection.section)
} # Function ParseSection ($PassedSection)

Function ParseSnippet ($PassedSnippet,$Outputfile,$SnippetCSS,$Indcrement,$SeparateSnippets,$InlineStats,$SimpleImageScale,$SmartImageScale,$KeepStyles,$Log) {
   switch ($PassedSnippet.type) { 
      "Audio" {
         $Type = "Audio File"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Audio"

      "Date_Game" {
         $Text = $PassedSnippet.game_date.display
         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $Text = $LabelPrefix + $Text
         } # if ($PassedSnippet.Label -ne $null)

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.Replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # Date_Game

      "Date_Range" {
         $Text = $PassedSnippet.date_range.display_start + " to " + $PassedSnippet.date_range.display_end
         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $Text = $LabelPrefix + $Text
         } # if ($PassedSnippet.Label -ne $null)
         
         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         }

         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.Replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # Date_Range

      "Foreign" {
         $Type = "Foreign Object"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Foreign"

      "HTML" {
         $Type = "HTML Page (Complete)"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"
         
         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "HTML"

      "Labeled_Text" {
         $Text = $PassedSnippet.contents

         $TagPreface = '<span class="RWSnippet">'
         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
            $Text = $Text.insert($InsertPoint,$LabelPrefix)
         } # if ($PassedSnippet.Label -ne $null)
         
         $Text = InsertMargin $Text $SnippetCSS $Log
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Labeled_Text"

      "Multi_Line" {
        if ($PassedSnippet.purpose -eq "directions_only") {
           $Text = $PassedSnippet.gm_directions
           $TagPreface = '<span class="RWSnippet">'
           $LabelPrefix = 'GM Directions: '
           $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
           $Text = $Text.insert($InsertPoint,$LabelPrefix)
           $Text = InsertMargin $Text $SnippetCSS $Log
           If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
           if ($SeparateSnippets) {$Text = $Text + "<hr>"}
           [System.IO.File]::AppendAllText($outputfile,$Text)
           # Add-Content -Path $Outputfile -Value $Text
        } elseif ($PassedSnippet.purpose -eq "Both") {
           if ($PassedSnippet.gm_directions -ne $null) {
              $Text = $PassedSnippet.gm_directions
              $TagPreface = '<span class="RWSnippet">'
              $LabelPrefix = 'GM Directions: '
              $InsertPoint = $Text.IndexOf($TagPreface) + $TagPreface.Length
              $Text = $Text.insert($InsertPoint,$LabelPrefix)
              $Text = InsertMargin $Text $SnippetCSS $Log
              If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
              if ($SeparateSnippets) {$Text = $Text + "<hr>"}
              [System.IO.File]::AppendAllText($outputfile,$Text)
              # Add-Content -Path $outputfile -Value $Text
           } # if ($PassedSnippet.gm_directions -ne $null)

           if ($PassedSnippet.contents -ne $null) {
              $Text = InsertMargin $PassedSnippet.contents $SnippetCSS $Log
              If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
              if ($SeparateSnippets) {$Text = $Text + "<hr>"}
              [System.IO.File]::AppendAllText($outputfile,$Text)
              # Add-Content -Path $outputfile -Value $Text
           } # if ($PassedSnippet.contents -ne $null)

        } elseif (($PassedSnippet.contents -ne $null) -and (($PassedSnippet.contents.contains("<ul")) -or ($PassedSnippet.contents.contains("<ol")) -or ($PassedSnippet.contents.contains("<table")))) {
           $TagList = GetTagLocations $PassedSnippet.contents $Log
           $Text = InsertFormattedMargins $PassedSnippet.contents $SnippetCSS $TagList $Log

           If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
           if ($SeparateSnippets) {$Text = $Text + "<hr>"}

           [System.IO.File]::AppendAllText($outputfile,$Text)
           # Add-Content -Path $outputfile -Value $Text
        } else {
           $Text = InsertMargin $PassedSnippet.contents $SnippetCSS $Log
           If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
           if ($SeparateSnippets) {$Text = $Text + "<hr>"}
           [System.IO.File]::AppendAllText($outputfile,$Text)
           # Add-Content -Path $outputfile -Value $Text
        } # if ($PassedSnippet.purpose -eq "directions_only")
      } # "Multi_Line"

      "Numeric" {
         $Text = $PassedSnippet.contents
         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $Text = $LabelPrefix + $Text
         } # if ($PassedSnippet.Label -ne $null)

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.Replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Numeric"

      "PDF" {
         $Type = "PDF Document"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "PDF"

      "Picture" {
         $FullImage = $PassedSnippet.ext_object.asset.contents
         $ImageLink = '<a href="data:image/png;base64,' + $FullImage + '">Picture</a>'
         $ImageName = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $ImageName + ": [$ImageLink]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         $Text = $SnippetCSS.replace("*",$Text)

         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $Outputfile -Value $Text

         If ($SimpleImageScale -eq 0) {
            $Thumbnail = $PassedSnippet.ext_object.asset.thumbnail
            $EncodedImage = '<img src="data:image/png;base64,' + $Thumbnail + '">'
         } else {
            $SimpleImageScale = $SimpleImageScale / 100
            $SimpleImageScale = 'style="transform:scale(' + $SimpleImageScale + ');"'
            $EncodedImage = '<img ' + $SimpleImageSCale + ' src="data:image/png;base64,' + $FullImage + '">'
         }
         $tag = "<img "
         $EncodedImage = InsertMiscMargins $EncodedImage $SnippetCSS $tag $Log
         if ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
         [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
         # Add-Content -Path $outputfile -Value $EncodedImage
      } # "Picture"

      "Portfolio" {
         $Statblock = "Hero Lab Portfolio"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Statblock, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Stats = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Stats)
         # Add-Content -Path $outputfile -Value $Stats
      } # "Portfolio"

      "Rich_Text" {
         $Type = "Rich Text Document"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Rich Text"

      "Smart_Image" {
         $FullImage = $PassedSnippet.smart_image.asset.contents
         $ImageLink = '<a href="data:image/png;base64,' + $FullImage + '">Smart Image</a>'
         $ImageName = $PassedSnippet.smart_image.name
         if (($ImageName -eq $null) -or ($ImageName -eq "") -or ($ImageName -eq "&nbsp;")) {$ImageName = "Unnamed"}
         $Text = $ImageName + ": [$ImageLink]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         $Text = $SnippetCSS.replace("*",$Text)

         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $Outputfile -Value $Text

         If ($SmartImageScale -eq 0) {
            $Thumbnail = $PassedSnippet.smart_image.asset.thumbnail
            $EncodedImage = '<img src="data:image/png;base64,' + $Thumbnail + '">'
         } else {
            $SmartImageScale = $SmartImageScale / 100
            $SmartImageScale = 'style="transform:scale(' + $SmartImageScale + ');"'
            $EncodedImage = '<img ' + $SmartImageSCale + ' src="data:image/png;base64,' + $FullImage + '">'
         }
         $tag = "<img "
         $EncodedImage = InsertMiscMargins $EncodedImage $SnippetCSS $tag $Log
         If ($SeparateSnippets) {$EncodedImage = $EncodedImage + "<hr>"}
         [System.IO.File]::AppendAllText($outputfile,$EncodedImage)
         # Add-Content -Path $outputfile -Value $EncodedImage
      } # "Smart_Image"

      "Statblock" {
         $Statblock = "Stat Block"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $EncodedStats = $PassedSnippet.ext_object.asset.contents
         $Stats = '<a href="data:text/html;base64,' + $EncodedStats + '">' + $Statblock + '</a>'
         $Text = $Name + ": [$Stats]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         If ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Stats = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Stats)
         # Add-Content -Path $outputfile -Value $Stats

         if ($InlineStats) {
            $DecodedStats = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($EncodedStats));
            $PTag = $DecodedStats.IndexOf("<p ")
            $BodyTag = $DecodedStats.IndexOf("</body>")
            $SubStringStart = $PTag
            $SubStringLength = $BodyTag - $PTag

            $Body = $DecodedStats.substring($SubStringStart,$SubStringLength)
            $Body = $Body.replace("RWLink","StatBlockLink")
            $Body = $Body.replace("RWSnippet","StatBlockSnippet")
            $Body = $Body.replace("RWDefault","StatBlockDefault")
            $Body = $Body.replace("RWBullet","StatBlockBullet")
            $Body = $Body.replace("RWEnumerated","StatBlockEnumerated")

            $Body = $Body.replace(".td",".StatBlock-td")
            $Body = $Body.replace(".tr","StatBlock-tr")
            $Body = $Body.replace(".p","StatBlock-p")

            If (-not $KeepStyles) {$Body = StripStyles $Body $Log}

            if ($SeparateSnippets) {$Body = $Body + "<hr>"}
            [System.IO.File]::AppendAllText($outputfile,$Body)
            # Add-Content -Path $outputfile -Value $Body
         } # if ($InlineStats)

      } # "Statblock"

      "Tag_Multi_Domain" {
         $Text = ParseTags $PassedSnippet $Log

         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $Text = $LabelPrefix + "<BR>$Text"
         } # if ($PassedSnippet.Label -ne $null)

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         $Text = $SnippetCSS.Replace("*",$Text)

         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Tag_Multi_Domain"

      "Tag_Standard" {
         $Text = ParseTags $PassedSnippet $Log

         if (($PassedSnippet.Label -ne $null) -and ($PassedSnippet.Label -ne "") -and ($PassedSnippet.Label -ne "&nbsp;")) {
            if ($PassedSnippet.Label.endswith(':')) {$LabelPrefix = $PassedSnippet.Label + ' '} else {$LabelPrefix = $PassedSnippet.Label + ': '}
            $Text = $LabelPrefix + "<BR>$Text"
         } # if ($PassedSnippet.Label -ne $null)

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         $Text = $SnippetCSS.Replace("*",$Text)
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Tag_Standard"

      "Video" {
         $Type = "Video File"
         $Name = $PassedSnippet.ext_object.name
         if (($name -eq $null) -or ($name -eq "") -or ($name -eq "&nbsp;")) {$name = "Unnamed"}
         $Text = $Name + ": [$Type, no preview available]"

         $annotation = $PassedSnippet.annotation
         $annotation = ParseSnippetText $annotation $Log
         if (($annotation -ne $null) -and ($annotation -ne "") -and ($annotation -ne "&nbsp;")) {
            $Text = $Text + "<br>Annotation: $annotation"
         } # if (($annotation -ne $null) -and ($annotation -ne ""))

         $Text = $Text.Trim()
         If (-not $KeepStyles) {$Text = StripStyles $Text $Log}
         if ($SeparateSnippets) {$Text = $Text + "<hr>"}
         $Text = $SnippetCSS.replace("*",$Text)
         [System.IO.File]::AppendAllText($outputfile,$Text)
         # Add-Content -Path $outputfile -Value $Text
      } # "Video"

      default {
         # Do nothing here.
      } # default
   } # switch ($PassedSnippet.type)
} # Function ParseSnippet

Function ParseSnippetText ($PassedText,$Log) {
   if ($PassedText -ne $null) {
      $gt = $PassedText.Indexof('>')
   } else {
      $gt = -1
   } # if ($PassedText -ne $null)
   $Text = $null
   While ($gt -ge 0) {
      # $gt = $PassedText.Indexof(">",$RWSnippet)
      $lt = $PassedText.Indexof("<",$gt)
      $length = $lt-$gt
      if ($length -gt 1) {$Text = $Text + $PassedText.substring($gt+1,$length-1)}
      if ($lt -lt $gt) {
         $gt = -1
      } else {
         $PassedText = $PassedText.substring($lt,$PassedText.length-$lt)
         $gt = $PassedText.Indexof('>')
      } # if ($lt -lt $gt)
   } # While ($gt -ge 0)
   # if ($Text.contains("&nbsp;")) {$Text = $Text.replace("&nbsp;","`r`n`r`n")}
   Return $Text
} # Function ParseSnippetText

Function ParseTags ($PassedTags,$Log) {
   # This block of code assumes that tags of the same domain will always be grouped together.
   $domain = 0
   $tagdomains = $PassedTags.tag_assign.domain_name
   if ($tagdomains.count -gt 1) {
      $tagdomain = $tagdomains[$domain]
   } else {
      $tagdomain = $tagdomains
   }
   $tagline = ""

   foreach ($tag in $PassedTags.tag_assign) {
      if ($tag.domain_name -eq $tagdomain) {
         if ($tagline -eq "") {$tagline = $tagdomain + ": " + $tag.tag_name} else {$tagline = $tagline + ", " + $tag.tag_name}
         $domain++
      } else {
         $tagdomain = $tag.domain_name
         $tagline = $tagline + "<br>" + $tagdomain + ": " + $tag.tag_name
         $domain++
      } # if ($tag.domain_name -eq $tagdomain)
   } # foreach ($tag in $PassedSnippet.tag_assign)

   Return $tagline.trim()
} # Function ParseTags

Function ParseLinkage ($PassedTopic,$OutputFile,$Log) {
   $Linkage = $PassedTopic.linkage
   $LinkList = ""
   foreach ($link in $linkage) {
      if ($linklist -eq "") {$linklist = "Linkage: " + $link.target_name} else {$linklist = $linklist + ", " + $link.target_name}
   } # foreach ($link in $linkage)
   Return $Linklist
} # Function ParseLinkage

Function GetTagLocations ($PassedSnippet,$Log) {
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
      $endtag = "</" + $tag.tag + ">"
      $endtag = $PassedSnippet.IndexOf($endtag,$tag.Start + 1)
      $Tag.end = $endtag + $endtag.length-1
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

Function InsertFormattedMargins ($PassedSnippet,$CSSCode,$TagLocations,$Log) {
   $FirstMargin = $CSSCode.IndexOf('s')
   $LastMargin = $CSSCode.IndexOf('>')
   $Length = $LastMargin-$FirstMargin
   $MarginString = $CSSCode.substring($FirstMargin,$Length)
   $MarginString = $MarginString + " "

   $Text = ""
   foreach ($tag in $TagLocations) {
      $tagtext = $PassedSnippet.substring($tag.start,$tag.end-$tag.start+1)
      Switch ($tag.tag) {
         "text" {
            $Text = $Text + $tagtext.replace('<p ',"<p $MarginString")
         } # "text"
            
         "ul" {
            $Text = $Text + $tagtext.replace('<ul ',"<ul $MarginString")
         } # "ul"

         "ol" {
            $Text = $Text + $tagtext.replace('<ol ',"<ol $MarginString")
         } # "ol"

         "table" {
            $Text = $Text + $tagtext.replace('<table ',"<table $MarginString")
            # $Text = $Text + $Text.replace('<p ',"<p $MarginString")
         } # "table"

      } # Switch ($tag.tag)
   } # foreach ($tag in $TagLocations)
   Return $Text
} # Function InsertFormattedMargins

Function InsertMiscMargins ($PassedSnippet,$CSSCode,$Tag,$Log) {
   $FirstMargin = $CSSCode.IndexOf('s')
   $LastMargin = $CSSCode.IndexOf('>')
   $Length = $LastMargin-$FirstMargin
   $MarginString = $CSSCode.substring($FirstMargin,$Length)
   $MarginString = $MarginString + " "
   $MarginString = $Tag + $MarginString
   $Text = $PassedSnippet.replace($tag,$MarginString)

   Return $Text
} # Function InsertMiscMargins

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

Function InsertMargin ($PassedSnippet,$CSSCode,$Log) {
   $FirstMargin = $CSSCode.IndexOf('s')
   $LastMargin = $CSSCode.IndexOf('>')
   $Length = $LastMargin-$FirstMargin
   $MarginString = $CSSCode.substring($FirstMargin,$Length)
   $MarginString = $MarginString + " "
   $PassedSnippet = $PassedSnippet.replace('<p ',"<p $MarginString")

   Return $PassedSnippet
} # Function InsertMargin


Function StripStyles ($Text,$Log) {
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

Function ReportException ($CommandLine,$ErrorMsg,$ErrorObject) {
   Write-Host "[ERROR]"
   Write-Host
   Write-Host "   $CommandLine"
   Write-Host
   Write-Host "      Error: $ErrorMsg"
   Write-Host "     Reason:" $ErrorObject.CategoryInfo.Category " - " $ErrorObject.ToString()
   Write-Host "   Location: Line" $ErrorObject.InvocationInfo.ScriptLineNumber ", Character" $ErrorObject.InvocationInfo.OffsetInLine
   Write-Host "       Line:" $ErrorObject.InvocationInfo.Line.Trim()
   Write-Host "[/ERROR]"
   # Write-Host
   # Read-Host 'Press Enter to continue…' | Out-Null
   Exit
} # Function FileException

# Main {
   # Set all errors as terminating errors to facilitate error trapping.
   $ErrorActionPreference = "Stop"

   # Get Date and CommandLine for logging purposes.
   $Date = Get-Date
   $CommandLine = $PSCmdlet.MyInvocation.Line

   # If the -Log option was invoked, log the current date and the full command line.
   # If anything goes wrong with this, write the error to the console and exit.
   if ($Log) {
      Try {
        $Date = $Date.ToString() + "`r`n"
        $CommandLine = $CommandLine + "`r`n"
        [System.IO.File]::WriteAllText($Log,$Date)
        [System.IO.File]::AppendAllText($Log,$CommandLine)
      } Catch {
        $ErrorMsg = "Error: Cannot create log file."
        ReportException $CommandLine $Error $_
        Exit
     } # Try
   } # if ($Log)


   # Setup CSS tags so all we have to do later is a string.replace of "*"
   #    with whatever text we wish to use. Also, the variable names will
   #    remind us which tags are for which elements.
   $TitleCSS = "<h1>*</h1>"
   $TopicCSS = '<h2 style="margin-left:5px;">*</h2>'
   $TopicDetailsCSS = '<h3 style="margin-left:5px;">*</h3>'
   $SectionCSS = '<h4 style="margin-left:5px;">*</h4>'
   $SnippetCSS = '<p style="margin-left:5px;">*</p>'
   $Indcrement = 50

   # Make sure the value specified for -Sort is valid.
   # If the value is invalid, write to the console and exit.
   if (($Sort -lt 1) -or ($Sort -gt 4)) {
      Write-Host "[ERROR]"
      Write-Host "   Invalid Sort value."
      Write-Host
      Write-Host "   Valid options are:"
      Write-Host "   1 = Sort by topic names"
      Write-Host "   2 = Sort by topic prefixes first, then by topic names (this is the default value)"
      Write-Host "   3 = Sort by topic category first, then by topic name"
      Write-Host "   4 = Sort by topic category first, then by topic prefix, then by topic name"
      Write-Host "[/ERROR]"
      # Write-Host
      # Read-Host 'Press Enter to continue…' | Out-Null
      Exit
   } # if (($Sort -lt 1) -or ($Sort -gt 4))

   # Make sure the values specified for -SimpleImageScale and -SmartImageScale are valid.
   # If either one is invalid, write to the console and exit.
   if (($SimpleImageScale -lt 0) -or ($SimpleImageScale -gt 100) -or ($SmartImageScale -lt 0) -or ($SmartImageScale -gt 100)) {
      Write-Host "[ERROR]"
      Write-Host "   SmartImageScale or SimpleImageScale out of range."
      Write-Host
      Write-Host "   Value must be between 0 and 100."
      Write-Host "[/ERROR]"
      # Write-Host
      # Read-Host 'Press Enter to continue…' | Out-Null
      Exit
   } # if (($Sort -lt 1) -or ($Sort -gt 4))

   # Import data from the specified source file.
   # If anything goes wrong with this, write the error to the console and exit.
   Try {
      [xml]$RWExportData = Get-Content -Path $Source
   } Catch {
        $ErrorMsg = "Error: Cannot read source file."
        ReportException $CommandLine $Error $_
   } # Try

   # Get the title from the specified output file.
   $Title = $RWExportData.Output.definition.details.name

   # Create and prime the HTML output file.
   # If anything goes wrong with this, write the error to the console and exit.
   Try {
      [System.IO.File]::WriteAllText($Destination,"<html>")
      [System.IO.File]::AppendAllText($Destination,"<head>")
      [System.IO.File]::AppendAllText($Destination,"<title>$Title</title>")
      [System.IO.File]::AppendAllText($Destination,'<link rel="stylesheet" type="text/css" href="' + $CSSFilename + '">')

      # Just in case this file gets posted on a public web site, tell googlebot not to index this page.
      [System.IO.File]::AppendAllText($Destination,'<meta name="googlebot" content="noindex">')

      # Continue priming the HTML output file.
      [System.IO.File]::AppendAllText($Destination,"</head>")
      [System.IO.File]::AppendAllText($Destination,"<body>")
      [System.IO.File]::AppendAllText($Destination,$TitleCSS.Replace("*",$Title))
   } Catch {
        $ErrorMsg = "Error: Cannot write destination file."
        ReportException $CommandLine $Error $_
   } # Try

   # Get the main content from the output file.
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
   foreach ($Topic in $TopicList) {
      ParseTopic $Topic $Destination $Sort $Prefix $Suffix $Details $Indent $SeparateSnippets $InlineStats $SimpleImageScale $SmartImageScale $Parent $TitleCSS $TopicCSS $TopicDetailsCSS $SectionCSS $SnippetCSS $Indcrement $KeepStyles $Log
   } # foreach ($Topic in $Contents.topic)

   # Close out the HTML file
   [System.IO.File]::AppendAllText($Destination,"</body>")
   [System.IO.File]::AppendAllText($Destination,"</html>")
# } Main