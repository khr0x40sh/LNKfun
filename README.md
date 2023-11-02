# LNKfun
A simple LNK file parser/editor written in PS1. The goal is to provide security assessors and defenders means to generate potentially malicious LNKs files for testing of defenses and/or user awareness.

This is currently a Work In Progress as eventually the script will generate a LNK from scratch. See Usage instructions on how to get the most out of this tool.

WARNING: Use of this code any portion of it is up to the user's own risk! Please practice good cyber hygiene and review the code (or any fork) before execution!!! 


## Usage:
__./Gen-LNK.ps1 <options>__

In the current release, you will need to provide an existing LNK. An example would be a shortcut that points to *%System32%\cmd.exe*.

Right-click on the shortcut and go to "Properties". Here there are some items that you will need to configure prior to execution of the *Gen-LNK.ps1*.

1. To change the icon, click on Change Icon, and then browse and select the ICO or EXE you wish to pull the icon from. An eexample would be using the Adobe Acrobat Reader Icon inside of Acrord32.exe
 ![alt text](https://github.com/khr0x40sh/LNKfun/blob/main/images/Properties.PNG "Selecting Adobe PDF Icon as LNK Icon")  
2. To ensure the Has Arguments flag is set, enter a space and a character of your choosing in the box AFTER cmd.exe. I like to use `/C ping 127.0.0.1`.
3. Save these changes.

To change the cursor-over text, create a text document with the data you wish to display. This will overwrite the "Comments" section (HasName field in the LNK file format). I have left *Adobe_desc.txt* as an example.
![alt_text](https://github.com/khr0x40sh/LNKfun/blob/main/images/cmd_HasName.png "CMD EXE default hover-over ")

![alt text](https://github.com/khr0x40sh/LNKfun/blob/main/images/OverwrittenDesc.png "Overwritten Comments to resemble a PDF")

To have the LNK execute a command other than just `cmd.exe`, you will need to create a text document with your desired command. I have left *command.txt* as an example (just executes `pwsh -command ping 127.0.0.1; Start-Process calc`)

### Options:
| Name | Description | Default Value
| ---- | ----| ----| 
| sourcepath | Source LNK file to parse (and subsequently edit) | $pwd\cmd.exe.lnk |
| bamboozle | Pad command with whitespace | *unset* |
| fakepath | Used with bamboozle, command to exec before whitespace | /C start msedge K:\ticket.pdf & |
| machineName | Value to stomp out machineName for OpSec | ABBY-PC |
| destination | Destination to save edited LNK file to | $pwd\ticket.pdf    (encrypted)            .lnk |
| argsFile | File containing Command Line Args to pass to the LNK | $pwd\commands.txt |  
| verbose | Verbose output | *unset* |
| appendFile | Append a File to the tail end of the LNK. This uses one of the reserved fields in the header to specify file start byte | *unset* |
| appendFilePath | File to append (used with above flag) | $pwd\ticket.pdf |
| useextra | Use the Extra Field instead of the tail of the LNK. Extremely Expiramental at the moment | *unset* |
| overwriteComments | Switch to Overwrite the Comment (HasName) field to make the LNK more "legitimate" | *unset* |
| commentPath | File containing overwriteComments (above flag) data | $pwd\Adobe_desc.txt |

### Example execution:
```
./Gen-LNK.ps1 -sourcepath $pwd\cmd.exe.lnk -bamboozle -fakepath "/C start calc &" -machineName ENTERPRISE-1701 -dest "$pwd\SecretPlans.pdf         .lnk" -argsFile $pwd\command.txt -overwriteComments -commentPath $pwd\Adobe_desc.txt
```


