<#
.SYNOPSIS
This script generates a HIDMap.psd1 from a collection of public domain USB HID
keyboard scan codes collected by Github user [MightyPork](https://github.com/MightyPork).
I really appreciate how nice and clean this source is!
.LINK
https://gist.githubusercontent.com/MightyPork/6da26e382a7ad91b5496ee55fdc73db2/raw/e91b2eca00fdf3d8b51a4dddc658913d2baa40e0/usb_hid_keys.h
#>

# get the source from MightyPork's repo
$SourceCode = Invoke-RestMethod -Uri 'https://gist.githubusercontent.com/MightyPork/6da26e382a7ad91b5496ee55fdc73db2/raw/e91b2eca00fdf3d8b51a4dddc658913d2baa40e0/usb_hid_keys.h' -UseBasicParsing -ErrorAction Stop

# this string holds all characters that are the result of holding SHIFT
$ShiftedChars = ''

# holds all defined keys, there are some duplicate characters in the spec, we only want one of each
$DefinedKeys = @{}

# pattern to parse the source file
$Pattern = '^#define KEY_(?!MOD|NONE|ERR|(?:LEFT|RIGHT)(?:ALT|SHIFT|CTRL|META))(?<KeyName>\S+) (?<Code>0x[a-f0-9]{2})(?:\s+\/\/\s+(?<Comment>.+ (?<PrimaryKey>.) and (?<SecondaryKey>.)|.*))?$'

$CustomMappings = @{
    'SPACE'     = ' '
    'BACKSPACE' = 'BS'
    'DELETE'    = 'DEL'
    'ESC'       = 'ESCAPE'
    'INSERT'    = 'INS'
    'TAB'       = "`t"
}

# variable to hold the output
[System.Collections.Generic.List[string]] $HIDMap = @()

# get the formatted output for the PSD1
$SourceCode -split "`n" -replace "`r" | ForEach-Object {

    # anti-pattern, skip lines that don't match
    if ( $_ -notmatch $Pattern ) { return }

    'KeyName', 'PrimaryKey', 'SecondaryKey'  | ForEach-Object {

        # skip the key name if there is a primary key
        if ( $_ -eq 'KeyName' -and $Matches.PrimaryKey ) {
            Write-Verbose ( 'Skipping {0} because primary key {1} and secondary key {2} are defined' -f $Matches.KeyName, $Matches.PrimaryKey, $Matches.SecondaryKey )
            return
        }
        
        # output HID map
        if ( $Matches.$_ -and -not $DefinedKeys[$Matches.$_] ) {
            $Line = '    {0,-20} = {1}{2}{3}' -f "'$($Matches.$_.Replace("'","''"))'", $Matches.Code, ('',' # ')[[bool]$Matches.Comment], $Matches.Comment
            $HIDMap.Add($Line)
            $DefinedKeys[$Matches.$_] = $true
            # process custom mappings
            if ( $CustomMappings.ContainsKey($Matches.KeyName) ) {
                $Line = '    {0,-20} = {1}{2}{3}' -f "'$($CustomMappings[$Matches.KeyName])'", $Matches.Code, ('',' # ')[[bool]$Matches.Comment], $Matches.Comment
                $HIDMap.Add($Line)
                $DefinedKeys[$CustomMappings[$Matches.KeyName]] = $true
            }
            if ( $_ -eq 'SecondaryKey' ) {
                $ShiftedChars += $Matches.SecondaryKey
            }
        }

    }

}

$HIDMap.Insert(0,'@{')
$HIDMap.Insert($HIDMap.Count,'}')

$HIDMap | Out-File "$PSScriptRoot\..\data\HIDMap.psd1" -Encoding ascii
$ShiftedChars | Out-File "$PSScriptRoot\..\data\ShiftedChars.txt" -Encoding ascii -NoNewline