<#
Input Validation Script for RWExport-To-HTML.ps1
by EightBitz
Version 1.0
2017-02-01, 01:00 PM CST
#>

param (
    # Source file path and name.
    [Parameter(Mandatory,Position=1)] 
    [string]$Script,

    [Parameter(Mandatory,Position=2)] 
    [string]$Source,

    # Destination file path and name.
    [Parameter(Mandatory,Position=3)] 
    [string]$Destination,

    # Destination file path and name.
    [Parameter(Mandatory,Position=4)] 
    [int]$Sort,

    # Destination file path and name.
    [Parameter(Mandatory,Position=5)] 
    [int]$SimpleImageScale,

    # Destination file path and name.
    [Parameter(Mandatory,Position=6)] 
    [int]$SmartImageScale,

    # Destination file path and name.
    [Parameter(Mandatory,Position=7)] 
    [bool]$ExtractFiles,

    # Destination file path and name.
    [Parameter(Mandatory,Position=8)] 
    [string]$CSSFileName,

    # Destination file path and name.
    [Parameter(Mandatory,Position=9)] 
    [bool]$SplitTopics,

    # Destination file path and name.
    [Parameter(Mandatory,Position=10)] 
    [bool]$CreateLog,

   # Destination file path and name.
    [Parameter(Mandatory,Position=11)][AllowNull()][AllowEmptyString()]
    [string]$Log,

    # Destination file path and name.
    [Parameter(Mandatory,Position=12)] 
    [string]$Validator
) # param

# Function ValidateInput($Script,$Source,$Destination,$Sort,$SimpleImageScale,$SmartImageScale,$ExtractFiles,$CSSFileName,$SplitTopics,$CreateLog,$Log,$Validator) {
   $ValidInput = $true
   
   if (($Script.StartsWith("'")) -and ($Script.EndsWith("'"))) {$Script = $Script.TrimStart("'");$Script = $Script.TrimEnd("'")}
   if (($Script.StartsWith('"')) -and ($Script.EndsWith('"'))) {$Script = $Script.TrimStart('"');$Script = $Script.TrimEnd('"')}
   if (($Script.StartsWith('&("')) -and ($Script.EndsWith('")'))) {$Script = $Script.TrimStart('&("');$Script = $Script.TrimEnd('")')}
   if (-not (Test-Path $Script)) {
      $MsgTitle = "Bad Script"
      $MsgText = "Cannot find PowerShell script."
      
      Switch ($Validator) {
         "Script" {
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
   } # if (-not (Test-Path $Script))

   if (($Source.StartsWith("'")) -and ($Source.EndsWith("'"))) {$Source = $Source.TrimStart("'");$Source = $Source.TrimEnd("'")}
   if (($Source.StartsWith('"')) -and ($Source.EndsWith('"'))) {$Source = $Source.TrimStart('"');$Source = $Source.TrimEnd('"')}
   $TestPath = Test-Path $Source
   if ($TestPath) {
      if (-not $Source.EndsWith(".rwoutput")) {
         $MsgTitle = "Bad Source File"
         $MsgText = "Source file should have an extension of .rwoutput."

         Switch ($Validator) {
            "Script" {
               $MsgText = $MsgText + "`r`nPress any key to exit..."
               Write-Host ""
               Write-Host $MsgTitle
               Write-Host $MsgText
               $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
               $ValidInput = $false
               Return $ValidInput
            } # "Script"
            "GUI" {
               $Msg = New-Object -ComObject Wscript.Shell
               $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
               $ValidInput = $false
               Return $ValidInput
            } # "GUI"
         } # Switch ($Validator)
      } # if (-not $Source.EndsWith(".rwoutput"))
   } else {
      $MsgTitle = "Bad Source File"
      $MsgText = "Source file does not exist."
      Switch ($Validator) {
         "Script" {
            $MsgText = $MsgText + "`r`nPress any key to exit..."
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
      Return $ValidInput
   }

   if (($Destination.StartsWith("'")) -and ($Destination.EndsWith("'"))) {$Destination = $Destination.TrimStart("'");$Destination = $Destination.TrimEnd("'")}
   if (($Destination.StartsWith('"')) -and ($Destination.EndsWith('"'))) {$Destination = $Destination.TrimStart('"');$Destination = $Destination.TrimEnd('"')}
   if ($SplitTopics -or $ExtractFiles) {
      $TestPath = Test-Path $Destination
      if ($TestPath) {
          $DestinationObject = Get-Item $Destination
          if ($DestinationObject.PSIsContainer) {
             $files = Get-ChildItem $Destination | Where-Object {!$_.PSIsContainer} | Measure-Object
             if ($files.Count -eq 0) {
                $EmptyPath = $true
             } else {
                $MsgTitle = "Destination Not Empty"
                $MsgText = "Destination folder contains files that may be overwritten.`r`nDo you wish to continue?"
                Switch ($Validator) {
                   "Script" {
                      $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disregards existing files."
                      $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exits to preserve existing files."
                      $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                      $result = $host.ui.PromptForChoice($MsgTitle, $MsgText, $options, 0) 
                      switch ($result) {
                         0 {Write-Host "You have chosen to overwrite the specified file(s)."}
                         1 {$ValidInput = $false; Return $ValidInput}
                      } # switch ($result)
                   } # "Script"
                   "GUI" {
                      $Msg = New-Object -ComObject Wscript.Shell
                      $YesOrNo = $Msg.Popup($MsgText,0,$MsgTitle,4)

                      Switch ($YesOrNo) {
                         6 {}
                         7 {$ValidInput = $false; Return $ValidInput}
                      } # Switch
                   } # "GUI"
                } # Switch ($Validator)
             } # if ((Get-ChildItem $Destination -force | Select-Object -First 1 | Measure-Object).Count -eq 0)

             if ($ExtractFiles) {
                if ($Destination.EndsWith("\")) {$ExtractPath = $Destination + "realm_files"} else {$ExtractPath = $Destination + "\realm_files"}

                if (Test-Path $ExtractPath) {
                   if ((Get-ChildItem $ExtractPath -force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
                      $EmptyPath = $true
                   } else {
                      $MSgTitle = "Extraction Folder Not Empty"
                      $MsgText = "The folder for your extracted files contains existing files that may be overwritten.`r`nDo you wish to continue?"
                      Switch ($Validator) {
                         "Script" {
                            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disregards existing files."
                            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exits to preserve existing files."
                            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                            $result = $host.ui.PromptForChoice($MsgTitle, $MsgText, $options, 0) 
                            switch ($result) {
                               0 {Write-Host "You have chosen to overwrite the specified file(s)."}
                               1 {$ValidInput = $false; Return $ValidInput}
                            } # switch ($result)
                         } # "Script"
                         "GUI" {
                            $Msg = New-Object -ComObject Wscript.Shell
                            $YesOrNo = $Msg.Popup($MsgText,0,$MsgTitle,4)

                            Switch ($YesOrNo) {
                               6 {}
                               7 {$ValidInput = $false; Return $ValidInput}
                            } # Switch
                         } # "GUI"
                      } # Switch ($Validator) 
                   } # if ((Get-ChildItem $ExtractPath -force | Select-Object -First 1 | Measure-Object).Count -eq 0)
                
                } else {
                   $MsgTitle = "Bad Extraction Folder"
                   $MsgText = "File extract folder does not exist.`r`nMake sure there is a 'realm_files' folder in your destination path."
                   Switch ($Validator) {
                      "Script" {
                         $MsgText = $MsgText + "`r`nPress any key to exit..."
                         Write-Host ""
                         Write-Host $MsgTitle
                         Write-Host $MsgText
                         $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                         $ValidInput = $false
                         Return $ValidInput
                      } # "Script"
                      "GUI" {
                         $Msg = New-Object -ComObject Wscript.Shell
                         $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                         $ValidInput = $false
                         Return $ValidInput
                      } # "GUI"
                   } # Switch ($Validator)
               } # if (Test-Path $ExtractPath)
             } # if ($ExtractFiles)
          } else {
             $MsgTitle = "Bad Destination Folder"
             $MsgText = "Destination folder does not exist.`r`n"
             $MsgText = $MsgText + "Note that when you specify -SplitTopics or -ExtractFiles,`r`n"
             $MsgText = $MsgText + "You must specify a folder as the destination instead of a file."
             Switch ($Validator) {
                "Script" {
                   $MsgText = $MsgText + "`r`nPress any key to exit..."
                   Write-Host ""
                   Write-Host $MsgTitle
                   Write-Host $MsgText
                   $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                   $ValidInput = $false
                   Return $ValidInput
                } # "Script"
                   "GUI" {
                   $Msg = New-Object -ComObject Wscript.Shell
                   $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                   $ValidInput = $false
                   Return $ValidInput
                } # "GUI"
             } # Switch ($Validator)
          } # if ($Destination.PSIsContainer)
      } else {
         $MsgTitle = "Bad Destination Folder"
         $MsgText = "Destination folder does not exist.`r`n"
         $Msg = New-Object -ComObject Wscript.Shell
         Switch ($Validator) {
            "Script" {
               $MsgText = $MsgText + "`r`nPress any key to exit..."
               Write-Host ""
               Write-Host $MsgTitle
               Write-Host $MsgText
               $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
               $ValidInput = $false
               Return $ValidInput
            } # "Script"
            "GUI" {
               $Msg = New-Object -ComObject Wscript.Shell
               $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
               $ValidInput = $false
               Return $ValidInput
            } # "GUI"
         } # Switch ($Validator)
      } # if ($TestPath)
   } else {
      if (($Destination.EndsWith(".html")) -or ($Destination.EndsWith(".htm"))) {
          if (Test-Path $Destination) {
             $MsgTitle = "Overwrite Destination?"
             $MsgText = "Destination file exists.`r`nDo you wish to overwrite?"
             Switch ($Validator) {
                "Script" {
                   $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disregards existing files."
                   $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exits to preserve existing files."
                   $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                   $result = $host.ui.PromptForChoice($MsgTitle, $MsgText, $options, 0) 
                   switch ($result) {
                      0 {Write-Host "You have chosen to overwrite the specified file(s)."}
                      1 {$ValidInput = $false; Return $ValidInput}
                   } # switch ($result)
                } # "Script"
                "GUI" {
                   $Msg = New-Object -ComObject Wscript.Shell
                   $YesOrNo = $Msg.Popup($MsgText,0,$MsgTitle,4)

                   Switch ($YesOrNo) {
                      6 {}
                      7 {$ValidInput = $false; Return $ValidInput}
                   } # Switch
                } # "GUI"
             } # Switch ($Validator) 
          } elseif ($Destination.Contains("\")) {
             $SplitDestination = $Destination.Split("\")
             $SplitCount = $SplitDestination.Count
             for ($Count = 0; $Count -le $SplitCount-2; $Count++) {
                $PathOnly = $PathOnly + $SplitDestination[$Count] + "\"
             } # if ($Destination.Contains("\"))
             $NameOnly = $SplitDestination[$SplitCount-1]

             if (Test-Path $PathOnly) {
                $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''   
                $ValidName = $true
                for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
                   if ($NameOnly.Contains($invalidChars[$CheckChar])) {$ValidName=$false}
                   if (-not $ValidName) {
                      $MsgTitle = "Bad Destination"
                      $MsgText = "Invalid Filename for Destination."
                      Switch ($Validator) {
                         "Script" {
                            $MsgText = $MsgText + "`r`nPress any key to exit..."
                            Write-Host ""
                            Write-Host $MsgTitle
                            Write-Host $MsgText
                            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            $ValidInput = $false
                            Return $ValidInput
                         } # "Script"
                         "GUI" {
                            $Msg = New-Object -ComObject Wscript.Shell
                            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                            $ValidInput = $false
                            Return $ValidInput
                         } # "GUI"
                      } # Switch ($Validator)
                   } # if (-not $ValidName)
                } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
            
             } else {
                $MsgTitle = "Bad Destination"
                $MsgText = "Destination folder does not exist."
                Switch ($Validator) {
                   "Script" {
                      $MsgText = $MsgText + "`r`nPress any key to exit..."
                      Write-Host ""
                      Write-Host $MsgTitle
                      Write-Host $MsgText
                      $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                      $ValidInput = $false
                      Return $ValidInput
                   } # "Script"
                   "GUI" {
                      $Msg = New-Object -ComObject Wscript.Shell
                      $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                      $ValidInput = $false
                      Return $ValidInput
                   } # "GUI"
                } # Switch ($Validator)
             } # if (Test-Path $PathOnly)

          } else {
             $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''   
             $ValidName = $true
             for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
                if ($Destination.Contains($invalidChars[$CheckChar])) {$ValidName=$false}
                if (-not $ValidName) {
                   $MsgTitle = "Bad Destination"
                   $MsgText = "Invalid filename for destination file."
                   Switch ($Validator) {
                      "Script" {
                         $MsgText = $MsgText + "`r`nPress any key to exit..."
                         Write-Host ""
                         Write-Host $MsgTitle
                         Write-Host $MsgText
                         $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                         $ValidInput = $false
                         Return $ValidInput
                      } # "Script"
                      "GUI" {
                         $Msg = New-Object -ComObject Wscript.Shell
                         $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                         $ValidInput = $false
                         Return $ValidInput
                      } # "GUI"
                   } # Switch ($Validator)
                } # if (-not $ValidName)
             } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
          } # if (Test-Path $Destination)
      } else {
         $MsgTitle = "Bad Destination"
         $MsgText = "Destination filename shouold end with .html or .htm."
         Switch ($Validator) {
            "Script" {
               $MsgText = $MsgText + "`r`nPress any key to exit..."
               Write-Host ""
               Write-Host $MsgTitle
               Write-Host $MsgText
               $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
               $ValidInput = $false
               Return $ValidInput
            } # "Script"
               "GUI" {
               $Msg = New-Object -ComObject Wscript.Shell
               $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
               $ValidInput = $false
               Return $ValidInput
            } # "GUI"
         } # Switch ($Validator)
      } # if (($Destination.EndsWith(".html")) -or ($Destination.EndsWith(".htm")))
   } # if ($SplitTopics -or $ExtractFiles)

   if (($Sort -lt 1) -or ($Sort -gt 4)) {
      $MsgTitle = "Bad Sort Value"
      $MsgText = "Sort value must be between 1 and 4, inclusively."
      Switch ($Validator) {
         "Script" {
            $MsgText = $MsgText + "`r`nPress any key to exit..."
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
   } # if (($Sort -lt 1) -or ($Sort -gt 4))

   $SimpleImageScale = $SimpleImageScale.ToInt32($null)
   if (($SimpleImageScale -lt 0) -or ($SimpleImageScale -gt 100)) {
      $MsgTitle = "Bad Scale Value"
      $MsgText = "Picture scale value must be between 0 and 100, inclusively."
      Switch ($Validator) {
         "Script" {
            $MsgText = $MsgText + "`r`nPress any key to exit..."
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
   } # if (($Sort -lt 1) -or ($Sort -gt 4))

   $SmartImageScale = $SmartImageScale.ToInt32($null)
   if (($SmartImageScale -lt 0) -or ($SmartImageScale -gt 100)) {
      $MsgTitle = "Bad Scale Value"
      $MsgText = "Smart Image scale value must be between 0 and 100, inclusively."
      Switch ($Validator) {
         "Script" {
            $MsgText = $MsgText + "`r`nPress any key to exit..."
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
   } # if (($Sort -lt 1) -or ($Sort -gt 4))

   if (($CSSFileName.StartsWith("'")) -and ($CSSFileName.EndsWith("'"))) {$CSSFileName = $CSSFileName.TrimStart("'");$CSSFileName = $CSSFileName.TrimEnd("'")}
   if (($CSSFileName.StartsWith('"')) -and ($CSSFileName.EndsWith('"'))) {$CSSFileName = $CSSFileName.TrimStart('"');$CSSFileName = $CSSFileName.TrimEnd('"')}
   $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''   
   $ValidName = $true
   if ($CSSFileName.EndsWith(".css")) {
      for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
         if ($CSSFileName.Contains($invalidChars[$CheckChar])) {$ValidName=$false}
         if (-not $ValidName) {
            $MsgTitle = "Bad Filename"
            $MsgText = "Invalid filename for CSS file."
            Switch ($Validator) {
               "Script" {
                  $MsgText = $MsgText + "`r`nPress any key to exit..."
                  Write-Host ""
                  Write-Host $MsgTitle
                  Write-Host $MsgText
                  $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                  $ValidInput = $false
                  Return $ValidInput
               } # "Script"
               "GUI" {
                  $Msg = New-Object -ComObject Wscript.Shell
                  $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                  $ValidInput = $false
                  Return $ValidInput
               } # "GUI"
            } # Switch ($Validator)
         } # if (-not $ValidName)
      } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
   } else {
      $MsgTitle = "Bad Destination"
      $MsgText = "CSS filename should end with .css."
      Switch ($Validator) {
         "Script" {
            $MsgText = $MsgText + "`r`nPress any key to exit..."
            Write-Host ""
            Write-Host $MsgTitle
            Write-Host $MsgText
            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $ValidInput = $false
            Return $ValidInput
         } # "Script"
         "GUI" {
            $Msg = New-Object -ComObject Wscript.Shell
            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
            $ValidInput = $false
            Return $ValidInput
         } # "GUI"
      } # Switch ($Validator)
   }

   if ($CreateLog) {
      if (($Log.StartsWith("'")) -and ($Log.EndsWith("'"))) {$Log = $Log.TrimStart("'");$Log = $Log.TrimEnd("'")}
      if (($Log.StartsWith('"')) -and ($Log.EndsWith('"'))) {$Log = $Log.TrimStart('"');$Log = $Log.TrimEnd('"')}
      
      If ($Log.EndsWith(".log")) {

          if (($Log -ne "") -and (Test-Path $Log)) {
         
             $MsgTitle = "Overwrite Existing Log?"
             $MsgText = "Log file exists.`r`nDo you wish to overwrite?"
             Switch ($Validator) {
                "Script" {
                   $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disregards existing files."
                   $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exits to preserve existing files."
                   $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                   $result = $host.ui.PromptForChoice($MsgTitle, $MsgText, $options, 0) 
                   switch ($result) {
                      0 {Write-Host "You have chosen to overwrite the specified file(s)."}
                      1 {$ValidInput = $false; Return $ValidInput}
                   } # switch ($result)
                } # "Script"
                "GUI" {
                   $Msg = New-Object -ComObject Wscript.Shell
                   $YesOrNo = $Msg.Popup($MsgText,0,$MsgTitle,4)
                   Switch ($YesOrNo) {
                      6 {}
                      7 {$ValidInput = $false; Return $ValidInput}
                   } # Switch
                } # "GUI"
             } # Switch ($Validator) 
          } elseif ($Log.Contains("\")) {
             $SplitLog = $Log.Split("\")
             $SplitCount = $SplitLog.Count
             for ($Count = 0; $Count -le $SplitCount-2; $Count++) {
                $PathOnly = $PathOnly + $SplitLog[$Count] + "\"
             } # if ($Log.Contains("\"))
             $NameOnly = $SplitLog[$SplitCount-1]

             if (Test-Path $PathOnly) {
                $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''   
                $ValidName = $true
                for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
                   if ($NameOnly.Contains($invalidChars[$CheckChar])) {$ValidName=$false}
                   if (-not $ValidName) {
                      $MsgTitle = "Bad Filename"
                      $MsgText = "Invalid filename for log file."
                      Switch ($Validator) {
                         "Script" {
                            $MsgText = $MsgText + "`r`nPress any key to exit..."
                            Write-Host ""
                            Write-Host $MsgTitle
                            Write-Host $MsgText
                            $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            $ValidInput = $false
                            Return $ValidInput
                         } # "Script"
                         "GUI" {
                            $Msg = New-Object -ComObject Wscript.Shell
                            $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                            $ValidInput = $false
                            Return $ValidInput
                         } # "GUI"
                      } # Switch ($Validator)
                   } # if (-not $ValidName)
                } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)
            
             } else {
                $MsgTitle = "Bad Folder"
                $MsgText = "Log folder does not exist."
                Switch ($Validator) {
                   "Script" {
                      $MsgText = $MsgText + "`r`nPress any key to exit..."
                      Write-Host ""
                      Write-Host $MsgTitle
                      Write-Host $MsgText
                      $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                      $ValidInput = $false
                      Return $ValidInput
                   } # "Script"
                   "GUI" {
                      $Msg = New-Object -ComObject Wscript.Shell
                      $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                      $ValidInput = $false
                      Return $ValidInput
                   } # "GUI"
                } # Switch ($Validator)
             } # if (Test-Path $PathOnly)
          } else {
             $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''   
             $ValidName = $true
             for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++) {
                if ($Log.Contains($invalidChars[$CheckChar])) {$ValidName=$false}
                if (($Log -eq "") -or ($Log -eq $null) -or (-not $ValidName)) {
                   $MsgTitle = "Bad Filename"
                   $MsgText = "Invalid filename for log file."
                   Switch ($Validator) {
                      "Script" {
                         $MsgText = $MsgText + "`r`nPress any key to exit..."
                         Write-Host ""
                         Write-Host $MsgTitle
                         Write-Host $MsgText
                         $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                         $ValidInput = $false
                         Return $ValidInput
                      } # "Script"
                      "GUI" {
                         $Msg = New-Object -ComObject Wscript.Shell
                         $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
                         $ValidInput = $false
                         Return $ValidInput
                      } # "GUI"
                   } # Switch ($Validator)
                } # if (-not $ValidName)
             } # for ($CheckChar = 0 ; $CheckChar -le 40 ; $Checkchar++)         
          } # if (($Log -ne "") -and (Test-Path $Log))
      } else {
         $MsgTitle = "Bad Destination"
         $MsgText = "Log filename shouold end with .log."
         Switch ($Validator) {
            "Script" {
               $MsgText = $MsgText + "`r`nPress any key to exit..."
               Write-Host ""
               Write-Host $MsgTitle
               Write-Host $MsgText
               $anykey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
               $ValidInput = $false
               Return $ValidInput
            } # "Script"
            "GUI" {
               $Msg = New-Object -ComObject Wscript.Shell
               $OKMsg = $Msg.Popup($MsgText,0,$MsgTitle,0)
               $ValidInput = $false
               Return $ValidInput
            } # "GUI"
         } # Switch ($Validator)
      } # If ($Log.EndsWith(".log"))
   } # if ($Log)
   Return $ValidInput
# } # Function ValidateInput