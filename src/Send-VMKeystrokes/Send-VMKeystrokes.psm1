using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace VMware.VimAutomation.ViCore.Types.V1.Inventory

#Requires -Module VMware.VimAutomation.Core

$Script:HIDMap = Import-PowerShellDataFile -Path "$PSScriptRoot\data\HIDMap.psd1"
$Script:ShiftedChars = [string[]][char[]]( Get-Content -Path "$PSScriptRoot\data\ShiftedChars.txt" )


class VirtualMachineTransformAttribute : ArgumentTransformationAttribute {
    
    # NOTE: output type MUST be [object]
    [object] Transform( [EngineIntrinsics]$EngineIntrinsics, [object]$InputObject ) {

        if ( $InputObject -is [VirtualMachine] ) {
            return $InputObject
        }

        if ( $InputObject -is [string] ) {

            $VirtualMachine = Get-VM -Name $InputObject -ErrorAction Stop

            if ( $VirtualMachine.Count -gt 1 ) {
                throw 'Function requires that a single virtual machine matches.'
            }

            return $VirtualMachine
        }
        
        throw 'Failed to translate to Virtual Machine. Input must be a Virtual Machine or string.'

    }

}


function Split-VMKeystrokeString {
    <#
    .SYNOPSIS
    Split a command string into chunks for processing
    .PARAMETER String
    The string to split
    .EXAMPLE
    'Test{BACKSPACE}String^!xExample' | Split-VMKeystrokeString
    T
    e
    s
    t
    {BACKSPACE}
    S
    t
    r
    i
    n
    g
    ^!x
    E
    x
    a
    m
    p
    l
    e
    #>
    [CmdletBinding()]
    [OutputType( [string[]] )]
    param(

        [Parameter( Mandatory, ValueFromPipeline, Position = 0 )]
        [AllowEmptyString()]
        [AllowNull()]
        [string[]]
        $String,

        [switch]
        $Verbatim

    )

    process {

        # always process input strings as a single string
        $JoinedString = $String -join ''

        # if the input string is null or empty we don't want to put an empty array on the pipeline
        if ( [string]::IsNullOrEmpty($JoinedString) ) { return }

        # if -Verbatim then we split the string up into single characters and return it
        if ( $Verbatim ) {
            return [string[]][char[]]$JoinedString
        }

        # otherwise process with our magic regex
        $JoinedString -split '(?=(?<=[^!+#^])[!+#^]+(?:[^!+#^{}]|\{\w+\}|\{[!+#^{}]\}))|(?<=(?<=^|[^!+#^])[!+#^]+(?:[^!+#^{}]|\{\w+\}|\{[!+#^{}]\}))|(?=(?<![!+#^])(?:\{\w+\}|\{[!+#^{}]\}))|(?<=\{\w+\}|\{[!+#^{}]\})' | ForEach-Object {

            # matches any special character, single character with a modifier, or single character
            # ex: {ENTER} or ^!{DELETE} or +a or a
            if ( $_ -match '^[!+#^]*(?:.|\{(?:[!+#^{}]|\w+)\})$' ) {
                return $_
            }
            
            # everything else should be single characters or escaped characters
            # the pattern below un-escapes any escaped characters, i.e. {{} should be {
            return [string[]][char[]]$_
            
        }

    }

}


function Convert-VMKeystrokeStringToHIDEvent {
    <#
    .SYNOPSIS
    Converts a string into a UsbScanCodeSpecKeyEvent object
    .PARAMETER Strings
    String(s) to convert into events
    #>
    [CmdletBinding()]
    [OutputType([VMware.Vim.UsbScanCodeSpecKeyEvent])]
    param(

        [Parameter( Mandatory, ValueFromPipeline, Position = 0 )]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $String

    )

    process {

        $String | ForEach-Object {
        
            Write-Debug ( 'Pre-Processed Character: {0}' -f $_ )

            if ( $_.Length -eq 1 ) {

                [string]$Character = $_
                [string[]]$Modifiers = @()

            } else {

                if ( $_ -notmatch '^(?:.|[!+#^]*(?:\{(?:[!+#^{}]|\w+)\}|[^!+#^{}]))$' ) {
                    throw ( 'Invalid character definition detected: {0}' -f $_ )
                }

                # This regex will always put at least a character into $Character since it
                # matches any modifiers and the beginning of the string. This means it can
                # handle: +x, ^!{DELETE}, and {ENTER}
                [string[]][char[]] $Modifiers, [string]$Character = $_ -split '(?<=^[!+#^]*)(?=[^!+#^])'

                # remove the curly braces around the special characters
                if ( $Character -match '^\{(?<Character>\w+|[!+#^{}])\}$') {
                    $Character = $Matches.Character
                }

            }

            Write-Debug ( 'Post-Processed Character: {0}' -f $Character )
            Write-Debug ( 'Post-Porcessed Modifiers: {0}' -f [string]$Modifiers )

            # Check to see if we've mapped the character to HID code
            if ( -not $Script:HIDMap.ContainsKey($Character) ) {
                throw ( 'The character ''{0}'' has not been mapped, you will need to remove or manually process this character.' -f $Character )
            }
            
            $UsbScanCodeSpecKeyEvent = New-Object VMware.Vim.UsbScanCodeSpecKeyEvent
            $UsbScanCodeSpecKeyEvent.UsbHidCode = ( [Int64]$Script:HIDMap[$Character] -shl 16 ) -bor 7
            $UsbScanCodeSpecKeyEvent.Modifiers = New-Object Vmware.Vim.UsbScanCodeSpecModifierType

            # use a shift modifier if idicated or
            # the character is in the ShiftedChars array or
            # the character is an upper case letter
            if ( $Modifiers -contains '+' ) {
                Write-Debug 'Character has the + prefix, adding SHIFT modifier'
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftShift = $true
            }
            elseif ( $Character -cmatch '^[A-Z]$' ) {
                Write-Debug 'Character is a capital letter, adding SHIFT modifier'
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftShift = $true
            }
            elseif ( $Script:ShiftedChars -contains $Character ) {
                Write-Debug ( 'Characters is one of {0}, adding SHIFT modifier' -f $Script:ShiftedChars )
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftShift = $true
            }

            # use an alt modifier if indicated
            if ( $Modifiers -contains '!' ) {
                Write-Debug 'Character has the ! prefix, adding ALT modifier'
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftAlt = $true
            }

            # use a ctrl modifier if indicated
            if ( $Modifiers -contains '^' ) {
                Write-Debug 'Character has the ^ prefix, adding CTRL modifier'
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftControl = $true
            }

            # use a OS key modifier if indicated
            if ( $Modifiers -contains '#' ) {
                Write-Debug 'Character has the # prefix, adding META modifier'
                $UsbScanCodeSpecKeyEvent.Modifiers.LeftGui = $true
            }

            Write-Debug ( 'Character: {0} -> HIDCode: 0x{1:x2} -> HIDCodeValue: 0x{2:x8} -> Modifiers: {3}' -f $Character, $Script:HIDMap[$Character], $UsbScanCodeSpecKeyEvent.UsbHidCode, ( $UsbScanCodeSpecKeyEvent.Modifiers.PSObject.Properties.Where({ $_.Value -eq $true }).Name -join '+' ) )

            $UsbScanCodeSpecKeyEvent

        }

    }

}


function Send-VMKeystrokeHIDEvent {
    <#
    .SYNOPSIS
    Sends USB keyboard HID events to a VMware virtual machine
    .DESCRIPTION
    Sends USB keyboard HID events to a VMware virtual machine
    .PARAMETER HIDEvents
    A list of HID events to send
    .PARAMETER VM
    The VMware virtual machine to send events to
    .PARAMETER KeyPressDelay
    The delay between key presses in milliseconds, the default is
    send all events together
    .PARAMETER PassThru
    PassThru the VMware virtual machine object on the pipeline
    #>
    [CmdletBinding( DefaultParameterSetName='Default' )]
    [OutputType( [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine], ParameterSetName='PassThru' )]
    param(

        [Parameter( Mandatory, Position = 0 )]
        [AllowEmptyString()]
        [VMware.Vim.UsbScanCodeSpecKeyEvent[]]
        $HIDEvents,

        [Parameter( Mandatory, ValueFromPipeline )]
        [Alias( 'VirtualMachine', 'Name' )]
        [VirtualMachineTransformAttribute()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]
        $VM,

        [ValidateRange(100,[uint]::MaxValue)]
        [System.Nullable[uint32]]
        $KeyPressDelay,

        [Parameter( Mandatory, ParameterSetName='PassThru' )]
        [switch]
        $PassThru

    )

    process {

        $VMView = Get-View -ViewType VirtualMachine -Filter @{ Name = "^${VM}$" } -ErrorAction Stop
        
        $UsbScanCodeSpec = New-Object Vmware.Vim.UsbScanCodeSpec

        Write-Debug ( 'Sending {0} HID events to {1}' -f $HIDEvents.Count, $VM )
        
        if ( $KeyPressDelay ) {
            for ( $i = 0; $i -lt $HIDEvents.Count; $i ++ ) {
                $UsbScanCodeSpec.KeyEvents = $HIDEvents[$i]
                [void] $VMView.PutUsbScanCodes($UsbScanCodeSpec)
                if ( ( $i + 1 ) -lt $HIDEvents.Count ) {
                    Write-Debug ( 'Sending USB HID code 0x{0:x2} with {1} ms delay' -f ( ( $HIDEvents[$i].UsbHidCode -bor 7 ) -shr 16 ), $KeyPressDelay )
                    Start-Sleep -Milliseconds $KeyPressDelay
                } else {
                    Write-Debug ( 'Sending USB HID code 0x{0:x2}' -f ( ( $HIDEvents[$i].UsbHidCode -bor 7 ) -shr 16 ) )
                }
            }
        } else {
            $UsbScanCodeSpec.KeyEvents = $HIDEvents
            [void] $VMView.PutUsbScanCodes($UsbScanCodeSpec)
        }

        if ( $PassThru ) {
            $VM
        }

    }
}


function Send-VMKeystrokes {
    <#
    .SYNOPSIS
    Sends a formatted string of keystrokes to a VMware virtual machine
    .DESCRIPTION
    Sends a formatted string of keystrokes to a VMware virtual machine. This
    function uses mostly the same format as AutoIt's Send function. See the
    HIDMap.psd1 file for a full list of supported keys.
    .PARAMETER String
    Formatted string(s) to send to the VMware virtual machine
    .PARAMETER VM
    The VMware virtual machine to send events to
    .PARAMETER KeyPressDelay
    The delay between key presses in milliseconds, the default is
    send all events together
    .PARAMETER PauseAfterSeconds
    Seconds to pause after sending keypresses, usefull to pipe multiple
    commands together for automation.
    .PARAMETER PauseAfterMilliseconds
    Milliseconds to pause after sending keypresses, usefull to pipe multiple
    commands together for automation.
    .PARAMETER PauseAfterDuration
    Duration specified as a timespan to pause after sending keypresses,
    usefull to pipe multiple commands together for automation.
    .PARAMETER SendCarriageReturn
    Send a trailing {ENTER} key
    .PARAMETER Verbatim
    Send the string without parsing form command keys, useful for sending
    passwords
    .PARAMETER PassThru
    PassThru the VMware virtual machine object on the pipeline
    #>
    [CmdletBinding( DefaultParameterSetName='Default' )]
    [OutputType( [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine], ParameterSetName='PassThru' )]
    [OutputType( [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine], ParameterSetName='PauseAfterSeconds_PassThru' )]
    [OutputType( [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine], ParameterSetName='PauseAfterMilliseconds_PassThru' )]
    [OutputType( [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine], ParameterSetName='PauseAfterDuration_PassThru' )]
    param(

        [Parameter( Mandatory, Position = 0 )]
        [AllowEmptyString()]
        [string[]]
        $String,

        [Parameter( Mandatory, ValueFromPipeline )]
        [Alias( 'VirtualMachine', 'Name' )]
        [VirtualMachineTransformAttribute()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]
        $VM,

        [ValidateRange(100,[uint]::MaxValue)]
        [System.Nullable[uint32]]
        $KeyPressDelay,

        [Parameter( Mandatory, ParameterSetName='PauseAfterSeconds' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterSeconds_PassThru' )]
        [System.Nullable[uint]]
        $PauseAfterSeconds,

        [Parameter( Mandatory, ParameterSetName='PauseAfterMilliseconds' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterMilliseconds_PassThru' )]
        [System.Nullable[uint]]
        $PauseAfterMilliseconds,

        [Parameter( Mandatory, ParameterSetName='PauseAfterDuration' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterDuration_PassThru' )]
        [timespan]
        $PauseAfterDuration,

        [switch]
        $SendCarriageReturn,

        [switch]
        $Verbatim,

        [Parameter( Mandatory, ParameterSetName='PassThru' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterSeconds_PassThru' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterMilliseconds_PassThru' )]
        [Parameter( Mandatory, ParameterSetName='PauseAfterDuration_PassThru' )]
        [switch]
        $PassThru

    )

    # convert input string in to list of HID events
    [List[object]]$HIDEvents = $String | Split-VMKeystrokeString -Verbatim:$Verbatim.IsPresent | Convert-VMKeystrokeStringToHIDEvent

    # Add return carriage to the end of the string input (useful for logins or executing commands)
    if ( $SendCarriageReturn ) {
        $HIDEvents.Add((Convert-VMKeystrokeStringToHIDEvent '{ENTER}'))
    }

    Write-Verbose ( 'Sending keystrokes to {0}' -f $VM )
    $KeyPressDelaySplat = @{}
    if ( $KeyPressDelay ) {
        $KeyPressDelaySplat.KeyPressDelay = $KeyPressDelay
    }
    $VM | Send-VMKeystrokeHIDEvent -HIDEvents $HIDEvents @KeyPressDelaySplat

    if ( $PauseAfterSeconds ) {
        $PauseAfterDuration = [timespan]::FromSeconds($PauseAfterSeconds)
    }

    if ( $PauseAfterMilliseconds ) {
        $PauseAfterDuration = [timespan]::FromMilliseconds($PauseAfterMilliseconds)
    }

    if ( $PauseAfterDuration ) {
        Start-Sleep -Duration $PauseAfterDuration
    }

    if ( $PassThru ) {
        $VM
    }

}


function ConvertTo-VMKeystrokeEscapedString {
    <#
    .SYNOPSIS
    Escapes control characters in a string, useful for sending
    passwords
    .PARAMETER String
    The string to escape
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(

    [Parameter( Mandatory, Position = 0, ValueFromPipeline )]
    [AllowEmptyString()]
    [string]
    $String
    
    )

    $String -replace '([!+#^{}])', '{$1}'

}
