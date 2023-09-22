# Send-VMKeystrokes
Module to facilitate sending keystrokes to a VMware virtual machine. Note that this module depends on PowerCLI being available.

## Formatting Strings
This module uses similar formatting for control characters and special characters to the [AutoIt Send](https://www.autoitscript.com/autoit3/docs/functions/Send.htm) function.

A string of just plain characters will be split up into a sequence of USB Keyboard HID Events and sent to the Virtual Machine.

## Control Characters
The characters "!", "+", "^", and "#" have special meanings. By prefixing any of these characters you will modify the next key in the sequence.

| Character | Modifier                 |
| --------- | ------------------------ |
|     !     | ALT                      |
|     +     | SHIFT                    |
|     ^     | CTRL                     |
|     #     | META (Windows/Apple Key) |

Some examples:
* ^c sends Ctrl + C
* ^!{DELETE} sends Ctrl + Alt + Delete
* !{F4} sends Alt + F4

## Special Characters
There are a number of keys on your keyboard that don't have a direct ASCII representation, to send those keys we use a HID map file.
You can view that file, and see a list of available special keys [here](src\Send-VMKeystrokes\data\HIDMap.psd1). Special characters
should be enclosed in curly braces (i.e. DELETE should be {DELETE} in your string).

## Escaping Characters
If you need to type any of the control characters or a literal curly brace you can wrap the character in curly braces. For example {+} is "+", {{} is "{", etc...

## Usage Examples

Sending a string of keypresses
```powershell
$VM | Send-VMKeystrokes 'hello world'
```

Sending Ctrl+Alt+Del
```powershell
$VM | Send-VMKeystrokes '^!{DELETE}'
```

You can also chain together sends, for example to login you might...
```powershell
$VM |
    Send-VMKeystrokes '^!{DELETE}' -PauseAfterMilliseconds 500 -PassThru |
    Send-VMKeystrokes $UserName -Verbatim -PassThru |
    Send-VMKeystrokes '{TAB}' -PassThru |
    Send-VMKeystrokes $PlainTextPassword -Verbatim -SendCarriageReturn
```

## Credit
Significant inspiration for this module came from code created by [William Lam](https://www.virtuallyghetto.com) and [David Rodriguez](https://www.sysadmintutorials.com). The primary code source I pulled from is [VMKeystrokes.ps1](https://github.com/lamw/vmware-scripts/blob/master/powershell/VMKeystrokes.ps1). Significant changes in both how this module handles input as well as how it's processed has been made.
