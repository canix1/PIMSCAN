 <#
.Synopsis
    PIMSCAN.ps1
     
    AUTHOR: Robin Granberg (robin.granberg@protonmail.com)
    
    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.
    
.DESCRIPTION
    Tool to create reports on Entra ID Role Assignments.

.EXAMPLE
    .\PIMSCAN.ps1 -TenantID "2e5097a7-4ead-42ae-82ef-c33d910626f6" 

.OUTPUTS
    HTML Report

.LINK
    

.NOTES
    **Version: 1.0**

    **11 April, 2024**


#>
Param
(
       # Tenant ID
    [Alias("tenant")]
    [Parameter(Mandatory=$true, 
                Position=1,
                ParameterSetName='Default')]
    [validatescript({$_ -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'})]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [String] 
    $TenantId = "",
    [string]
    $HTMLReport = "",
    [Switch]
    $InterActive,
    [Switch]
    $NoPIM,
    [Switch]
    $ClearCache,
    [Switch]
    $Show    
)
#$VerbosePreference = "Continue"
$ToolName = "PIMSCAN"
$Author = "Robin Granberg @canix1"
$Version = "1.0"
$ThemeBackGrounColor = "#131313"


Function Write-BlockFont
{
    # AUTHOR: Robin Granberg (robin.granberg@protonmail.com)
    # 9 November, 2023
    # Version: 1.0.0
    [CmdletBinding(DefaultParameterSetName = "Letter")]
    param (
        [Parameter(Mandatory=$True)]
        [string]
        $Phrase,

        # Color 1.
        [Parameter(Mandatory=$false)]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $Color1 = "White",

        # Color 2.
        [Parameter(Mandatory=$false)]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $Color2 = "White", 
           
        #ShadowColor  
        [Parameter(Mandatory=$false)]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $ShadowColor = "White",    

        #FrameColor  
        [Parameter(Mandatory=$false)]
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $FrameColor = "White",      

        [Parameter(Mandatory=$false,
        ParameterSetName='Letter')]
        [ValidateSet(888,9617,9618,9619)]
        [int]
        $LetterChar=9619,      

        [Parameter(Mandatory=$false,
        ParameterSetName='Custom')]
        [int]
        $CustomChar,

        [Parameter(Mandatory=$false)]
        [switch]
        $Frame 

    )


if($CustomChar)
{
    $WrittenLetterChar =$CustomChar
}
else
{
    $WrittenLetterChar=$LetterChar
}


[String[]]$Lframe = "A  ","*","A  ","*","A  ","*","A  ","*","A  ","*","A  "
[String[]]$Rframe = "BY","*","AS","*","AS", "*","AS","*","AS","*","AS"
[String[]]$SPACE = "      ","*","      ","*","      ", "*","      ","*","      ","*","      "
[String[]]$DOT = "      ","*","      ","*","      ", "*","      ","*","PPY   ","*","OIC   "
[String[]]$DASH = "           ","*","           ","*","PPPPPPY    ","*","OIIIIIC    ","*","           ","*","           "
[String[]]$A = "  PPPPPPY     ","*","PPLIIIIIPPY   ","*","PPPPPPPPPPS   ","*","PPLIIIIIPPS   ","*","PPS     PPS   ","*","OIC     OIC   "
[String[]]$B = "PPPPPPPY     ","*","PPLIIIIPPY   ", "*","PPPPPPPLIC   ", "*","PPLIIIIPPY   ", "*","PPPPPPPLIC   ","*","OIIIIIIC     " 
[String[]]$C = "  PPPPPPY   ","*","PPLIIIIIC   ","*","PPS         ","*","PPS         ","*","OIPPPPPPY   ", "*","  OIIIIIC   "
[String[]]$D = "PPPPPPPY     ","*","PPLIIIIPPY   ","*","PPS    PPS   ","*","PPS    PPS   ","*","PPPPPPPLC    ","*","OIIIIIIC     "
[String[]]$E = "PPPPPPPY   ","*","PPLIIIIC   ","*","PPPPPPY    ","*","PPLIIIC    ","*","PPPPPPPY   ","*","OIIIIIIC   "
[String[]]$F = "PPPPPPPY   ","*","PPLIIIIC   ","*","PPPPPPY    ","*","PPLIIIC    ","*","PPS        ","*","OIC        "
[String[]]$G = "  PPPPPPPY   ","*","PPLIIIIIIC   ","*","PPS  PPPPY   ","*","PPS    PPS   ","*","OIPPPPPPPS   ","*","  OIIIIIIC   "
[String[]]$H = "PPY    PPY   ","*","PPS    PPS   ","*","PPPPPPPPPS   ","*","PPLIIIIPPS   ","*","PPS    PPS   ","*","OIC    OIC   "
[String[]]$I = "PPY   ","*","PPS   ","*","PPS   ", "*","PPS   ","*","PPS   ","*","OIC   "
[String[]]$J = "     PPY   ","*","     PPS   ","*","     PPS   ","*","     PPS   ","*","PPPPPPLC   ","*","OIIIIIC    "
[String[]]$K = "PPY   PPY   ","*","PPS PPLIC   ","*","PPPPPLC     ","*","PPLIPPY     ","*","PPS  OPPY   ", "*","OIC   OIC   "
[String[]]$L = "PPY        ","*","PPS        ","*","PPS        ","*","PPS        ","*","PPPPPPPY   ","*","OIIIIIIC   "
[String[]]$M = "PPY     PPY   ","*","PPPPY PPPPS   ","*","PPLIPPLIPPS   ","*","PPS OIC PPS   ","*","PPS     PPS   ","*","OIC     OIC   "
[String[]]$N = "PPY     PPY   ","*","PPPPY   PPS   ","*","PPLIPPY PPS   ","*","PPS OIPPPPS   ","*","PPS   OIPPS   ","*","OIC     OIC   " 
[String[]]$O = "  PPPPPPY     ","*","PPLIIIIIPPY   ","*","PPS     PPS   ","*","PPS     PPS   ","*","OIPPPPPPLIC   ","*","  OIIIIIC     "
[String[]]$P = "PPPPPPY     ","*","PPLIIIPPY   ","*","PPPPPPLIC   ","*","PPLIIIC     ","*","PPS         ","*","OIC         "
[String[]]$Q = "  PPPPPPY       ","*","PPLIIIIIPPY     ","*","PPS     PPS     ","*","PPS   PPPPS     ","*","OIPPPPPPLIPPY   ","*","  OIIIIIC OIC   "
[String[]]$R = "PPPPPPY     ","*","PPLIIIPPY   ","*","PPPPPPLIC   ","*","PPLIIPPY    ","*","PPS   PPY   ","*","OIC   OIC   "
[String[]]$S = "  PPPPPPPPY   ","*","PPLIIIIIIIC   ","*","OIPPPPPPY     ","*","  OIIIIIPPY   ","*","PPPPPPPPLIC   ","*","OIIIIIIIC     "
[String[]]$T = "PPPPPPPPPPY   ","*"," OIIPPLIIIC   ","*","    PPS       ","*","    PPS       ","*","    PPS       ","*","    OIC       "
[String[]]$U = "PPY     PPY   ","*","PPS     PPS   ","*","PPS     PPS   ","*","PPS     PPS   ","*","OIPPPPPPLIC   ","*","  OIIIIIC     "
[String[]]$V = "PPY     PPY   ","*","PPS     PPS   ","*","OIPPY PPLIC   ","*","  PPS PPS     ","*","  OIPPLIC     ","*","    OIC       "
[String[]]$W = "PPY     PPY   ","*","PPS PPY PPS   ","*","PPS PPS PPS   ","*","PPS PPS PPS   ","*","OIPPPPPPLIC   ","*","  OIIIIIC     "
[String[]]$X = "PPY     PPY   ","*","OIPPY PPLIC   ","*","  OIPPLIC     ","*","  PPLCPPY     ","*","PPLIC OIPPY   ","*","OIC     OIC   " 
[String[]]$Y = "PP      PPY   ","*","PP      PPS   ","*","OIPPPPPPLIC   ","*","  OIPPLIC     ","*","    PPS       ","*","    OIC       "
[String[]]$Z = "PPPPPPPPPPY   ","*","OIIIIIPPLIC   ","*","    PPLIC     ","*","  PPLIC       ","*","PPPPPPPPPPY   ","*","OIIIIIIIIIC   "
[String[]]$1 = "PPPPY     ","*","OIPPS     ","*","  PPS     ", "*","  PPS     ","*","PPPPPPY   ","*","OIIIIIC   "
[String[]]$2 = "PPPPPPPPPPY   ","*","OIIIIIIIPPS   ","*","  PPPPPPLIC   ","*","PPLIIIIIC     ","*","PPPPPPPPPPY   ","*","OIIIIIIIIIC   "
[String[]]$3 = "PPPPPPPY     ","*","OIIIIIIPPY   ", "*","  PPPPPLIC   ", "*","  OIIIIPPY   ", "*","PPPPPPPLIC   ","*","OIIIIIIC     " 
[String[]]$4 = "PPY    PPY   ","*","PPS    PPS   ","*","PPPPPPPPPS   ","*","OIIIIIIPPS   ","*","       PPS   ","*","       OIC   "
[String[]]$5 = "PPPPPPPPPPY   ","*","PPLIIIIIIIC   ","*","OIPPPPPPY     ","*","  OIIIIIPPY   ","*","PPPPPPPPLIC   ","*","OIIIIIIIC     "
[String[]]$6 = "  PPPPPPPPY   ","*","PPLIIIIIIIC   ","*","PPPPPPPPY     ","*","PPOIIIIIPPY   ","*","OIPPPPPPLIC   ","*","  OIIIIIC     "
[String[]]$7 = "PPPPPPPPPPY   ","*","OIIIIIIIPPS   ","*","      PPLIC   ","*","    PPLIC     ","*","    PPS       ","*","    OIC       "
[String[]]$8 = "  PPPPPPY     ","*","PPLIIIIIPPY   ","*","  PPPPPPLIC   ","*","PPLIIIIIPPY   ","*","  PPPPPPPLC   ","*","  OIIIIIIC    "
[String[]]$9 = "PPPPPPPPPPY   ","*","PPLIIIIIPPS   ","*","PPPPPPPPPPS   ","*","OIIIIIIIPPS   ","*","PPPPPPPPPPS   ","*","OIIIIIIIIIC   "
[String[]]$0 = "PPPPPPPPPPY   ","*","PPLIIIIIPPS   ","*","PPS     PPS   ","*","PPS     PPS   ","*","PPPPPPPPPPS   ","*","OIIIIIIIIIC   "

$intCharBorderHorizontal1 = 9552
$intCharBorderLowRight1 = 9565
$intCharBorderHorizontal2 = 9472
$intCharBorderVertical1 = 9553
$intCharBorderLowLeft1 = 9562
$intCharBorderLowRight2 = 9496
$intCharBorderLowLeft2 = 9492
$intCharBorderVertical2 = 9474
$intCharBorderUpperRight1 = 9556
$intCharBorderUpperLeft1 = 9559
$intCharBorderUpperRight2 = 9488
$intCharBorderUpperLeft2 =9484
$intCharBorderLowLeft2Connect = 9573
$intCharBorderVerticalConnect2 = 9566

Function Test-ShadowStrings
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Testchar
    )

    # if the char is not the same as the main letter char return True
    switch ($Testchar)
    {
        $([char]$WrittenLetterChar){$false}
        default{$true}
    }


}
Function Create-BlockText
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Phrase,
        [Parameter()]
        [string]
        $ForegroundColor1,
        [Parameter()]
        [string]
        $ForegroundColor2       
    )
    
    if($Phrase)
    {
        $PHRASEOBJECT = New-Object System.Collections.ArrayList
        if($Frame)
        {
            [VOID]$PHRASEOBJECT.add($Lframe)
        }
        $arrString = $Phrase.ToCharArray() 
        foreach($Char in $arrString)
        {
            switch ($Char) {
                "A" {[VOID]$PHRASEOBJECT.add($A)}
                "B" {[VOID]$PHRASEOBJECT.add($B)}
                "C" {[VOID]$PHRASEOBJECT.add($C)}
                "D" {[VOID]$PHRASEOBJECT.add($D)}
                "E" {[VOID]$PHRASEOBJECT.add($E)}
                "F" {[VOID]$PHRASEOBJECT.add($F)}
                "G" {[VOID]$PHRASEOBJECT.add($G)}                
                "H" {[VOID]$PHRASEOBJECT.add($H)}                  
                "I" {[VOID]$PHRASEOBJECT.add($I)}
                "J" {[VOID]$PHRASEOBJECT.add($J)} 
                "K" {[VOID]$PHRASEOBJECT.add($K)}
                "L" {[VOID]$PHRASEOBJECT.add($L)}
                "M" {[VOID]$PHRASEOBJECT.add($M)}                
                "N" {[VOID]$PHRASEOBJECT.add($N)}
                "O" {[VOID]$PHRASEOBJECT.add($O)}
                "Q" {[VOID]$PHRASEOBJECT.add($Q)}                
                "P" {[VOID]$PHRASEOBJECT.add($P)}                
                "R" {[VOID]$PHRASEOBJECT.add($R)}
                "S" {[VOID]$PHRASEOBJECT.add($S)}
                "T" {[VOID]$PHRASEOBJECT.add($T)}                
                "U" {[VOID]$PHRASEOBJECT.add($U)}                   
                "V" {[VOID]$PHRASEOBJECT.add($V)}
                "W" {[VOID]$PHRASEOBJECT.add($W)}
                "X" {[VOID]$PHRASEOBJECT.add($X)}
                "Y" {[VOID]$PHRASEOBJECT.add($Y)}                
                "Z" {[VOID]$PHRASEOBJECT.add($Z)}
                "1" {[VOID]$PHRASEOBJECT.add($1)}
                "2" {[VOID]$PHRASEOBJECT.add($2)}
                "3" {[VOID]$PHRASEOBJECT.add($3)}
                "4" {[VOID]$PHRASEOBJECT.add($4)}                
                "5" {[VOID]$PHRASEOBJECT.add($5)}                     
                "6" {[VOID]$PHRASEOBJECT.add($6)}                      
                "7" {[VOID]$PHRASEOBJECT.add($7)}                
                "8" {[VOID]$PHRASEOBJECT.add($8)}                                
                "9" {[VOID]$PHRASEOBJECT.add($9)}
                "0" {[VOID]$PHRASEOBJECT.add($0)}
                " " {[VOID]$PHRASEOBJECT.add($SPACE)}
                "." {[VOID]$PHRASEOBJECT.add($DOT)}
                "-" {[VOID]$PHRASEOBJECT.add($DASH)}
                Default {[VOID]$PHRASEOBJECT.add($DOT)}
            }

        }
    }
    if($Frame)
    {
        [VOID]$PHRASEOBJECT.add($Rframe)
    }
    [int]$LenOfAll = 0
    foreach($LETTER in $PHRASEOBJECT)
    {
        $LenOfAll = $LenOfAll + $LETTER[0].ToString().Length
        
    }    
    $LenOfAll = $LenOfAll -2
    if($Frame)
    {
        for ($num = 0 ; $num -le $LenOfAll ; $num++)
        {    
            if($num -eq 0)
            {
                Write-Host -NoNewline $([char]$intCharBorderUpperLeft2) -ForegroundColor $FrameColor
            }
            else {            
                if($num -eq $LenOfAll)
                {
                    Write-Host "$([char]$intCharBorderUpperRight2)" -ForegroundColor $FrameColor
                }
                else {
                    Write-Host -NoNewline "$([char]$intCharBorderHorizontal2)"  -ForegroundColor $FrameColor
                }
            }
        }    
    }
    
    for ($num = 0 ; $num -le 12 ; $num++)
    {    
        $LetterCount = 0
        $FrameLetter = $false
        foreach($LETTER in $PHRASEOBJECT)
        {
            $LetterCount++
            if($Frame)
            {
                if(($LetterCount -eq 1) -or ($LetterCount -eq $PHRASEOBJECT.Count))
                {
                    $FrameLetter = $True
                }
                else {
                    $FrameLetter = $false
                }
            }
            
            foreach($LINE in $LETTER[$num])
            {
                $LINE = $LINE.ToString().Replace("P",[char]$WrittenLetterChar)
                $LINE = $LINE.ToString().Replace("Y","$([char]$intCharBorderUpperLeft1)")
                $LINE = $LINE.ToString().Replace("S","$([char]$intCharBorderVertical1)")  
                $LINE = $LINE.ToString().Replace("C","$([char]$intCharBorderLowRight1)") 
                $LINE = $LINE.ToString().Replace("O","$([char]$intCharBorderLowLeft1)") 
                $LINE = $LINE.ToString().Replace("I","$([char]$intCharBorderHorizontal1)") 
                $LINE = $LINE.ToString().Replace("L","$([char]$intCharBorderUpperRight1)")
                $LINE = $LINE.ToString().Replace("A","$([char]$intCharBorderVertical2)")
                $LINE = $LINE.ToString().Replace("B","$([char]$intCharBorderVerticalConnect2)")
                 
                 
                if($LINE -eq "*")
                {
                    if($PREVLINE -ne "*")
                    {
                        Write-Host ""
                    }
                    $PREVLINE = $LINE
                }
                else {
                    if($FrameLetter)
                    {
                        Write-Host -NoNewline $LINE -ForegroundColor $FrameColor
                    }
                    else {
                        
                    
                    if($ForegroundColor1)
                    {
                        if($num -le 6)
                        {
                            if($num -ge 2)
                            {
                                #Write-Host -NoNewline $LINE -ForegroundColor $ForegroundColor1
                                $arrChars = $LINE.ToCharArray() 

                                foreach($Char in $arrChars)
                                {
                                    if(Test-ShadowStrings $Char)
                                    {
                                        Write-Host -NoNewline $Char -ForegroundColor $ShadowColor
                                    } 
                                    else {
                                        Write-Host -NoNewline $Char -ForegroundColor $ForegroundColor1
                                    }
                                }                                
                            }
                            else {
                                #Write-Host -NoNewline $LINE -ForegroundColor $ForegroundColor2
                                $arrChars = $LINE.ToCharArray() 

                                foreach($Char in $arrChars)
                                {
                                    if(Test-ShadowStrings $Char)
                                    {
                                        Write-Host -NoNewline $Char -ForegroundColor $ShadowColor
                                    } 
                                    else {
                                        Write-Host -NoNewline $Char -ForegroundColor $ForegroundColor2
                                    }
                                }                                      
                            }
                        }
                        else {
                            #Write-Host -NoNewline $LINE -ForegroundColor $ForegroundColor2
                            $arrChars = $LINE.ToCharArray() 

                            foreach($Char in $arrChars)
                            {
                                if(Test-ShadowStrings $Char)
                                {
                                    Write-Host -NoNewline $Char -ForegroundColor $ShadowColor
                                } 
                                else {
                                    Write-Host -NoNewline $Char -ForegroundColor $ForegroundColor2
                                }
                            }                              
                        }
                    }
                    else {
                        #Write-Host -NoNewline  $LINE  
                        $arrChars = $LINE.ToCharArray() 

                        foreach($Char in $arrChars)
                        {
                            if(Test-ShadowStrings $Char)
                            {
                                Write-Host -NoNewline $Char -ForegroundColor $ShadowColor
                            } 
                            else {
                                Write-Host -NoNewline $Char 
                            }
                        }                  
                    }
                }
                    $PREVLINE = $LINE
                }
            }
        }

    }
    Write-Host "" 
    if($Frame)
    {
        for ($num = 0 ; $num -le $LenOfAll ; $num++)
        {    
            if($num -eq 0)
            {
                Write-Host -NoNewline "$([char]$intCharBorderLowLeft2)" -ForegroundColor $FrameColor
            }
            elseif($num -eq 1)
            {
                Write-Host -NoNewline "$([char]$intCharBorderLowLeft2Connect)" -ForegroundColor $FrameColor
            }
            else {            
                if($num -eq $LenOfAll)
                {
                    Write-Host "$([char]$intCharBorderLowRight2)$([char]$intCharBorderVertical1)" -ForegroundColor $FrameColor
                }
                else {
                    Write-Host -NoNewline "$([char]$intCharBorderHorizontal2)"  -ForegroundColor $FrameColor
                }
            }
        }    
    }    
    if($Frame)
    {
        for ($num = 1 ; $num -le $LenOfAll ; $num++)
        {    
            if($num -eq 1)
            {
                Write-Host -NoNewline " $([char]$intCharBorderLowLeft1)" -ForegroundColor $FrameColor
            }
            else {            
                if($num -eq $LenOfAll)
                {
                    Write-Host "$([char]$intCharBorderHorizontal1)$([char]$intCharBorderLowRight1)" -ForegroundColor $FrameColor
                }
                else {
                    Write-Host -NoNewline "$([char]$intCharBorderHorizontal1)"  -ForegroundColor $FrameColor 
                }
            }
        }    
    }        
    Write-Host "" 
}

if ($Phrase) {
    
    if($Color1)
    {
        Create-BlockText -Phrase $Phrase -ForegroundColor1 $Color1 -ForegroundColor2 $Color2
    }
    else {
        Create-BlockText -Phrase $Phrase 
    }

}    
else {
    break
}

}




Write-BlockFont -Phrase $ToolName -Frame -Color1 Yellow -Color2 Cyan -ShadowColor Magenta -FrameColor Red 
Write-Host "Author: $($Author) ; Version: $($Version) " -ForegroundColor Yellow 
Write-Host "$(("$([char]9876)-" * 30).TrimEnd("-"))" -ForegroundColor Yellow

$strHTMLTextCurrent = @"
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body {background-color:$ThemeBackGrounColor;font-family: Arial;color: #ffffff  }

#TopTable table {
width: 100%;
border-style: none;
}
#TopTable th {
font-family: Arial;
color: #00BFFF;
text-align: center;
border-style: none;
}
#TopTable td {
border-style: none;
text-align: center;
vertical-align: top
}
#TopTable tr:hover {
background-color: transparent;
}
#TopTable tr:nth-child(even) {
background-color: #000000;
}

#RoleInfoTbl table {
width: 25%;
border-style: none;
}
#RoleInfoTbl th {
font-family: Arial;
color: #00BFFF;
text-align: center;
border-style: none;
}
#RoleInfoTbl td {
border-style: none;
text-align: center;
vertical-align: top
}
#RoleInfoTbl tr:hover {
background-color: transparent;
}
#RoleInfoTbl tr:nth-child(even) {
background-color: #000000;
}

#RoleMemberTblStyle table {
border-style: none;
}
#RoleMemberTblStyle th {
font-family: Arial;
color: #00BFFF;
text-align: center;
border-style: none;
}
#RoleMemberTblStyle td {
border-style: none;
text-align: center;
vertical-align: top
}
#RoleMemberTblStyle tr:hover {
background-color: transparent;
}
#RoleMemberTblStyle tr:nth-child(even) {
background-color: #000000;
}

table {
    width: 100%;
}

th {
    font-family: Arial;
    color: #00BFFF;
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
    border-right: 1px solid #ffffff;
    border-bottom: 1px solid #ffffff;
}
td {
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
    border-right: 1px solid #ffffff;
    border-bottom: 1px solid #ffffff;
}

tr:hover {background-color: coral;}

/* Style the tab */
.tab {
overflow: hidden;
border: 1px solid #ccc;
background-color: #f1f1f1;
}

/* Style the tab */
.tab_sum {
overflow: hidden;
border: 1px solid #ccc;
background-color: #abb0b3;
}

/* Style the buttons inside the tab */
.tab button {
background-color: inherit;
float: left;
border: none;
outline: none;
cursor: pointer;
padding: 14px 16px;
transition: 0.3s;
font-size: 17px;
}

/* Style the buttons inside the tab */
.tab_sum button {
background-color: inherit;
float: left;
border: none;
outline: none;
cursor: pointer;
padding: 14px 16px;
transition: 0.3s;
font-size: 17px;
}

.priv-box {
    background-color: #e6aa06;
    float: right;
    border: 1px solid #ccc;
    outline: none;
    cursor: pointer;
    padding: 3px 3px;
    transition: 0.3s;
    font-size: 9px;
}

.warn-box {
    background-color: #DC143C;
    float: right;
    border: 1px solid #ccc;
    outline: none;
    cursor: pointer;
    padding: 3px 3px;
    transition: 0.3s;
    font-size: 9px;
}

.WARN { color: #DC143C; }


/* Change background color of buttons on hover */
.tab button:hover {
background-color: #ddd;
}

/* Create an active/current tablink class */
.tab button.active {
background-color: #ccc;
}


/* Change background color of buttons on hover */
.tab_sum button:hover {
background-color: #ddd;
}

/* Create an active/current tablink class */
.tab_sum button.active {
background-color: #ccc;
}

/* Style the tab content */
.tabcontent {
display: none;
padding: 6px 12px;
border: 1px solid #ccc;
border-top: none;
}

/* Style the tab content */
.tabcontent_sum {
display: none;
padding: 6px 12px;
border: 1px solid #ccc;
border-top: none;
}
</style>
</head>
<body>
<script>

function openSummary(evt, Summary) {
var i, tabcontent, tablinks;
tabcontent = document.getElementsByClassName("tabcontent_sum");
for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
}
tablinks = document.getElementsByClassName("tablinks_sum");
for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
}
document.getElementById(Summary).style.display = "block";
evt.currentTarget.className += " active";
}

function openRole(evt, RoleName) {
var i, tabcontent, tablinks;
tabcontent = document.getElementsByClassName("tabcontent");
for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
}
tablinks = document.getElementsByClassName("tablinks");
for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
}
document.getElementById(RoleName).style.display = "block";
evt.currentTarget.className += " active";
}

/* W3.JS 1.04 April 2019 by w3schools.com */
"use strict";
var w3 = {};
w3.hide = function (sel) {
w3.hideElements(w3.getElements(sel));
};
w3.hideElements = function (elements) {
var i, l = elements.length;
for (i = 0; i < l; i++) {
    w3.hideElement(elements[i]);
}
};
w3.hideElement = function (element) {
w3.styleElement(element, "display", "none");
};
w3.show = function (sel, a) {
var elements = w3.getElements(sel);
if (a) {w3.hideElements(elements);}
w3.showElements(elements);
};
w3.showElements = function (elements) {
var i, l = elements.length;
for (i = 0; i < l; i++) {
    w3.showElement(elements[i]);
}
};
w3.showElement = function (element) {
w3.styleElement(element, "display", "block");
};
w3.addStyle = function (sel, prop, val) {
w3.styleElements(w3.getElements(sel), prop, val);
};
w3.styleElements = function (elements, prop, val) {
var i, l = elements.length;
for (i = 0; i < l; i++) {    
    w3.styleElement(elements[i], prop, val);
}
};
w3.styleElement = function (element, prop, val) {
element.style.setProperty(prop, val);
};
w3.toggleShow = function (sel) {
var i, x = w3.getElements(sel), l = x.length;
for (i = 0; i < l; i++) {    
    if (x[i].style.display == "none") {
    w3.styleElement(x[i], "display", "block");
    } else {
    w3.styleElement(x[i], "display", "none");
    }
}
};
w3.addClass = function (sel, name) {
w3.addClassElements(w3.getElements(sel), name);
};
w3.addClassElements = function (elements, name) {
var i, l = elements.length;
for (i = 0; i < l; i++) {
    w3.addClassElement(elements[i], name);
}
};
w3.addClassElement = function (element, name) {
var i, arr1, arr2;
arr1 = element.className.split(" ");
arr2 = name.split(" ");
for (i = 0; i < arr2.length; i++) {
    if (arr1.indexOf(arr2[i]) == -1) {element.className += " " + arr2[i];}
}
};
w3.removeClass = function (sel, name) {
w3.removeClassElements(w3.getElements(sel), name);
};
w3.removeClassElements = function (elements, name) {
var i, l = elements.length, arr1, arr2, j;
for (i = 0; i < l; i++) {
    w3.removeClassElement(elements[i], name);
}
};
w3.removeClassElement = function (element, name) {
var i, arr1, arr2;
arr1 = element.className.split(" ");
arr2 = name.split(" ");
for (i = 0; i < arr2.length; i++) {
    while (arr1.indexOf(arr2[i]) > -1) {
    arr1.splice(arr1.indexOf(arr2[i]), 1);     
    }
}
element.className = arr1.join(" ");
};
w3.toggleClass = function (sel, c1, c2) {
w3.toggleClassElements(w3.getElements(sel), c1, c2);
};
w3.toggleClassElements = function (elements, c1, c2) {
var i, l = elements.length;
for (i = 0; i < l; i++) {    
    w3.toggleClassElement(elements[i], c1, c2);
}
};
w3.toggleClassElement = function (element, c1, c2) {
var t1, t2, t1Arr, t2Arr, j, arr, allPresent;
t1 = (c1 || "");
t2 = (c2 || "");
t1Arr = t1.split(" ");
t2Arr = t2.split(" ");
arr = element.className.split(" ");
if (t2Arr.length == 0) {
    allPresent = true;
    for (j = 0; j < t1Arr.length; j++) {
    if (arr.indexOf(t1Arr[j]) == -1) {allPresent = false;}
    }
    if (allPresent) {
    w3.removeClassElement(element, t1);
    } else {
    w3.addClassElement(element, t1);
    }
} else {
    allPresent = true;
    for (j = 0; j < t1Arr.length; j++) {
    if (arr.indexOf(t1Arr[j]) == -1) {allPresent = false;}
    }
    if (allPresent) {
    w3.removeClassElement(element, t1);
    w3.addClassElement(element, t2);          
    } else {
    w3.removeClassElement(element, t2);        
    w3.addClassElement(element, t1);
    }
}
};
w3.getElements = function (id) {
if (typeof id == "object") {
    return [id];
} else {
    return document.querySelectorAll(id);
}
};
w3.filterHTML = function(id, sel, filter) {
var a, b, c, i, ii, iii, hit;
a = w3.getElements(id);
for (i = 0; i < a.length; i++) {
    b = a[i].querySelectorAll(sel);
    for (ii = 0; ii < b.length; ii++) {
    hit = 0;
    if (b[ii].innerText.toUpperCase().indexOf(filter.toUpperCase()) > -1) {
        hit = 1;
    }
    c = b[ii].getElementsByTagName("*");
    for (iii = 0; iii < c.length; iii++) {
        if (c[iii].innerText.toUpperCase().indexOf(filter.toUpperCase()) > -1) {
        hit = 1;
        }
    }
    if (hit == 1) {
        b[ii].style.display = "";
    } else {
        b[ii].style.display = "none";
    }
    }
}
};
w3.sortHTML = function(id, sel, sortvalue) {
var a, b, i, ii, y, bytt, v1, v2, cc, j;
a = w3.getElements(id);
for (i = 0; i < a.length; i++) {
    for (j = 0; j < 2; j++) {
    cc = 0;
    y = 1;
    while (y == 1) {
        y = 0;
        b = a[i].querySelectorAll(sel);
        for (ii = 0; ii < (b.length - 1); ii++) {
        bytt = 0;
        if (sortvalue) {
            v1 = b[ii].querySelector(sortvalue).innerText;
            v2 = b[ii + 1].querySelector(sortvalue).innerText;
        } else {
            v1 = b[ii].innerText;
            v2 = b[ii + 1].innerText;
        }
        v1 = v1.toLowerCase();
        v2 = v2.toLowerCase();
        if ((j == 0 && (v1 > v2)) || (j == 1 && (v1 < v2))) {
            bytt = 1;
            break;
        }
        }
        if (bytt == 1) {
        b[ii].parentNode.insertBefore(b[ii + 1], b[ii]);
        y = 1;
        cc++;
        }
    }
    if (cc > 0) {break;}
    }
}
};
w3.slideshow = function (sel, ms, func) {
var i, ss, x = w3.getElements(sel), l = x.length;
ss = {};
ss.current = 1;
ss.x = x;
ss.ondisplaychange = func;
if (!isNaN(ms) || ms == 0) {
    ss.milliseconds = ms;
} else {
    ss.milliseconds = 1000;
}
ss.start = function() {
    ss.display(ss.current)
    if (ss.ondisplaychange) {ss.ondisplaychange();}
    if (ss.milliseconds > 0) {
    window.clearTimeout(ss.timeout);
    ss.timeout = window.setTimeout(ss.next, ss.milliseconds);
    }
};
ss.next = function() {
    ss.current += 1;
    if (ss.current > ss.x.length) {ss.current = 1;}
    ss.start();
};
ss.previous = function() {
    ss.current -= 1;
    if (ss.current < 1) {ss.current = ss.x.length;}
    ss.start();
};
ss.display = function (n) {
    w3.styleElements(ss.x, "display", "none");
    w3.styleElement(ss.x[n - 1], "display", "block");
}
ss.start();
return ss;
};
w3.includeHTML = function(cb) {
var z, i, elmnt, file, xhttp;
z = document.getElementsByTagName("*");
for (i = 0; i < z.length; i++) {
    elmnt = z[i];
    file = elmnt.getAttribute("w3-include-html");
    if (file) {
    xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4) {
        if (this.status == 200) {elmnt.innerHTML = this.responseText;}
        if (this.status == 404) {elmnt.innerHTML = "Page not found.";}
        elmnt.removeAttribute("w3-include-html");
        w3.includeHTML(cb);
        }
    }      
    xhttp.open("GET", file, true);
    xhttp.send();
    return;
    }
}
if (cb) cb();
};
w3.getHttpData = function (file, func) {
w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
    func(this.responseText);
    }
});
};
w3.getHttpObject = function (file, func) {
w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
    func(JSON.parse(this.responseText));
    }
});
};
w3.displayHttp = function (id, file) {
w3.http(file, function () {
    if (this.readyState == 4 && this.status == 200) {
    w3.displayObject(id, JSON.parse(this.responseText));
    }
});
};
w3.http = function (target, readyfunc, xml, method) {
var httpObj;
if (!method) {method = "GET"; }
if (window.XMLHttpRequest) {
    httpObj = new XMLHttpRequest();
} else if (window.ActiveXObject) {
    httpObj = new ActiveXObject("Microsoft.XMLHTTP");
}
if (httpObj) {
    if (readyfunc) {httpObj.onreadystatechange = readyfunc;}
    httpObj.open(method, target, true);
    httpObj.send(xml);
}
};
w3.getElementsByAttribute = function (x, att) {
var arr = [], arrCount = -1, i, l, y = x.getElementsByTagName("*"), z = att.toUpperCase();
l = y.length;
for (i = -1; i < l; i += 1) {
    if (i == -1) {y[i] = x;}
    if (y[i].getAttribute(z) !== null) {arrCount += 1; arr[arrCount] = y[i];}
}
return arr;
};  
w3.dataObject = {},
w3.displayObject = function (id, data) {
var htmlObj, htmlTemplate, html, arr = [], a, l, rowClone, x, j, i, ii, cc, repeat, repeatObj, repeatX = "";
htmlObj = document.getElementById(id);
htmlTemplate = init_template(id, htmlObj);
html = htmlTemplate.cloneNode(true);
arr = w3.getElementsByAttribute(html, "w3-repeat");
l = arr.length;
for (j = (l - 1); j >= 0; j -= 1) {
    cc = arr[j].getAttribute("w3-repeat").split(" ");
    if (cc.length == 1) {
    repeat = cc[0];
    } else {
    repeatX = cc[0];
    repeat = cc[2];
    }
    arr[j].removeAttribute("w3-repeat");
    repeatObj = data[repeat];
    if (repeatObj && typeof repeatObj == "object" && repeatObj.length != "undefined") {
    i = 0;
    for (x in repeatObj) {
        i += 1;
        rowClone = arr[j];
        rowClone = w3_replace_curly(rowClone, "element", repeatX, repeatObj[x]);
        a = rowClone.attributes;
        for (ii = 0; ii < a.length; ii += 1) {
        a[ii].value = w3_replace_curly(a[ii], "attribute", repeatX, repeatObj[x]).value;
        }
        (i === repeatObj.length) ? arr[j].parentNode.replaceChild(rowClone, arr[j]) : arr[j].parentNode.insertBefore(rowClone, arr[j]);
    }
    } else {
    console.log("w3-repeat must be an array. " + repeat + " is not an array.");
    continue;
    }
}
html = w3_replace_curly(html, "element");
htmlObj.parentNode.replaceChild(html, htmlObj);
function init_template(id, obj) {
    var template;
    template = obj.cloneNode(true);
    if (w3.dataObject.hasOwnProperty(id)) {return w3.dataObject[id];}
    w3.dataObject[id] = template;
    return template;
}
function w3_replace_curly(elmnt, typ, repeatX, x) {
    var value, rowClone, pos1, pos2, originalHTML, lookFor, lookForARR = [], i, cc, r;
    rowClone = elmnt.cloneNode(true);
    pos1 = 0;
    while (pos1 > -1) {
    originalHTML = (typ == "attribute") ? rowClone.value : rowClone.innerHTML;
    pos1 = originalHTML.indexOf("{{", pos1);
    if (pos1 === -1) {break;}
    pos2 = originalHTML.indexOf("}}", pos1 + 1);
    lookFor = originalHTML.substring(pos1 + 2, pos2);
    lookForARR = lookFor.split("||");
    value = undefined;
    for (i = 0; i < lookForARR.length; i += 1) {
        lookForARR[i] = lookForARR[i].replace(/^\s+|\s+$/gm, ''); //trim
        if (x) {value = x[lookForARR[i]];}
        if (value == undefined && data) {value = data[lookForARR[i]];}
        if (value == undefined) {
        cc = lookForARR[i].split(".");
        if (cc[0] == repeatX) {value = x[cc[1]]; }
        }
        if (value == undefined) {
        if (lookForARR[i] == repeatX) {value = x;}
        }
        if (value == undefined) {
        if (lookForARR[i].substr(0, 1) == '"') {
            value = lookForARR[i].replace(/"/g, "");
        } else if (lookForARR[i].substr(0,1) == "'") {
            value = lookForARR[i].replace(/'/g, "");
        }
        }
        if (value != undefined) {break;}
    }
    if (value != undefined) {
        r = "{{" + lookFor + "}}";
        if (typ == "attribute") {
        rowClone.value = rowClone.value.replace(r, value);
        } else {
        w3_replace_html(rowClone, r, value);
        }
    }
    pos1 = pos1 + 1;
    }
    return rowClone;
}
function w3_replace_html(a, r, result) {
    var b, l, i, a, x, j;
    if (a.hasAttributes()) {
    b = a.attributes;
    l = b.length;
    for (i = 0; i < l; i += 1) {
        if (b[i].value.indexOf(r) > -1) {b[i].value = b[i].value.replace(r, result);}
    }
    }
    x = a.getElementsByTagName("*");
    l = x.length;
    a.innerHTML = a.innerHTML.replace(r, result);
}
};
"@


if($HTMLReport -eq "")
{
    $HTMLReport =$(join-path -Path $PSScriptRoot -ChildPath "\EntraID_Role_Report_$($TenantId).html") 
}
else {
    
}

$CSVFile =$(join-path -Path $PSScriptRoot -ChildPath "\EntraID_Role_Report_$($TenantId).csv")    

$dicEntraLicensing = @{"ADV_COMMS"="Advanced Communications";
"CDSAICAPACITY"="AI Builder Capacity add-on";
"SPZA_IW"="App Connect IW";
"Microsoft_Cloud_App_Security_App_Governance_Add_On"="App governance add-on to Microsoft Defender for Cloud Apps";
"MCOMEETADV"="Microsoft 365 Audio Conferencing";
"MCOMEETADV_FACULTY"="Microsoft 365 Audio Conferencing for faculty";
"AAD_BASIC"="Microsoft Entra Basic";
"AAD_PREMIUM"="Microsoft Entra ID P1";
"AAD_PREMIUM_FACULTY"="Microsoft Entra ID P1 for faculty";
"AAD_PREMIUM_USGOV_GCCHIGH"="Microsoft Entra ID P1_USGOV_GCCHIGH";
"AAD_PREMIUM_P2"="Microsoft Entra ID P2";
"RIGHTSMANAGEMENT"="Azure Information Protection Plan 1";
"RIGHTSMANAGEMENT_CE"="Azure Information Protection Plan 1";
"RIGHTSMANAGEMENT_CE_GOV"="Azure Information Protection Premium P1 for Government";
"RIGHTSMANAGEMENT_CE_USGOV_GCCHIGH"="Azure Information Protection Premium P1_USGOV_GCCHIGH";
"OFFICEBASIC"="Basic Collaboration";
"SMB_APPS"="Business Apps (free)";
"CAREERCOACH_FACULTY"="Career Coach for faculty";
"CAREERCOACH_STUDENTS"="Career Coach for students";
"Clipchamp_Premium"="Clipchamp Premium";
"Clipchamp_Standard"="Clipchamp Standard";
"CDS_FILE_CAPACITY"="Common Data Service for Apps File Capacity";
"CDS_DB_CAPACITY"="Common Data Service Database Capacity";
"CDS_DB_CAPACITY_GOV"="Common Data Service Database Capacity for Government";
"CDS_LOG_CAPACITY"="Common Data Service Log Capacity";
"MCOPSTNC"="Communications Credits";
"CMPA_addon"="Compliance Manager Premium Assessment Add-On";
"CMPA_addon_GCC"="Compliance Manager Premium Assessment Add-On for GCC";
"Defender_Threat_Intelligence"="Defender Threat Intelligence";
"MESSAGING_GCC_TEST"="Digital Messaging for GCC Test SKU";
"CRMSTORAGE"="Dynamics 365 - Additional Database Storage (Qualified Offer)";
"CRMINSTANCE"="Dynamics 365 - Additional Production Instance (Qualified Offer)";
"CRMTESTINSTANCE"="Dynamics 365 - Additional Non-Production Instance (Qualified Offer)";
"SOCIAL_ENGAGEMENT_APP_USER"="Dynamics 365 AI for Market Insights (Preview)";
"DYN365_ASSETMANAGEMENT"="Dynamics 365 Asset Management Addl Assets";
"DYN365_BUSCENTRAL_ADD_ENV_ADDON"="Dynamics 365 Business Central Additional Environment Addon";
"DYN365_BUSCENTRAL_DB_CAPACITY"="Dynamics 365 Business Central Database Capacity";
"DYN365_BUSCENTRAL_ESSENTIAL"="Dynamics 365 Business Central Essentials";
"DYN365_FINANCIALS_ACCOUNTANT_SKU"="Dynamics 365 Business Central External Accountant";
"PROJECT_MADEIRA_PREVIEW_IW_SKU"="Dynamics 365 Business Central for IWs";
"DYN365_BUSCENTRAL_PREMIUM"="Dynamics 365 Business Central Premium";
"DYN365_BUSCENTRAL_TEAM_MEMBER"="Dynamics 365 Business Central Team Members";
"DYN365_RETAIL_TRIAL"="Dynamics 365 Commerce Trial";
"DYN365_ENTERPRISE_PLAN1"="Dynamics 365 Plan 1 Enterprise Edition";
"DYN365_CUSTOMER_INSIGHTS_ATTACH"="Dynamics 365 Customer Insights Attach";
"DYN365_CUSTOMER_INSIGHTS_BASE"="Dynamics 365 Customer Insights Standalone";
"DYN365_CUSTOMER_INSIGHTS_VIRAL"="Dynamics 365 Customer Insights vTrial";
"DYN365_CS_CHAT_GOV"="Dynamics 365 for Customer Service Chat for Government";
"DYN365_CS_MESSAGING_GOV"="Dynamics 365 for Customer Service Digital Messaging add-on for Government";
"DYN365_CS_OC_MESSAGING_VOICE_GOV"="Dynamics 365 Customer Service Digital Messaging and Voice Add-in for Government";
"DYN365_CS_OC_MESSAGING_VOICE_GOV_TEST"="Dynamics 365 Customer Service Digital Messaging and Voice Add-in for Government for Test";
"Dynamics_365_Customer_Service_Enterprise_viral_trial"="Dynamics 365 Customer Service Enterprise Viral Trial";
"D365_CUSTOMER_SERVICE_ENT_ATTACH"="Dynamics 365 for Customer Service Enterprise Attach to Qualifying Dynamics 365 Base Offer A";
"DYN365_AI_SERVICE_INSIGHTS"="Dynamics 365 Customer Service Insights Trial";
"FORMS_PRO"="Dynamics 365 Customer Voice Trial";
"DYN365_CUSTOMER_SERVICE_PRO"="Dynamics 365 Customer Service Professional";
"DYN365_CUSTOMER_VOICE_BASE"="Dynamics 365 Customer Voice";
"Forms_Pro_AddOn"="Dynamics 365 Customer Voice Additional Responses";
"DYN365_CUSTOMER_VOICE_ADDON"="Dynamics 365 Customer Voice Additional Responses";
"Forms_Pro_USL"="Dynamics 365 Customer Voice USL";
"CRMSTORAGE_GCC"="Dynamics 365 Enterprise Edition - Additional Database Storage for Government";
"CRMTESTINSTANCE_NOPREREQ"="Dynamics 365 - Additional Non-Production Instance for Government";
"CRMTESTINSTANCE_GCC"="Dynamics 365 Enterprise Edition - Additional Non-Production Instance for Government";
"CRM_ONLINE_PORTAL"="Dynamics 365 Enterprise Edition - Additional Portal (Qualified Offer)";
"CRM_ONLINE_PORTAL_GCC"="Dynamics 365 Enterprise Edition - Additional Portal for Government";
"CRM_ONLINE_PORTAL_NOPREREQ"="Dynamics 365 Enterprise Edition - Additional Portal for Government";
"CRMINSTANCE_GCC"="Dynamics 365 Enterprise Edition - Additional Production Instance for Government";
"Dynamics_365_Field_Service_Enterprise_viral_trial"="Dynamics 365 Field Service Viral Trial";
"DYN365_FINANCE"="Dynamics 365 Finance";
"DYN365_ENTERPRISE_CASE_MANAGEMENT"="Dynamics 365 for Case Management Enterprise Edition";
"D365_ENTERPRISE_CASE_MANAGEMENT_GOV"="Dynamics 365 for Case Management, Enterprise Edition for Government";
"DYN365_ENTERPRISE_CASE_MANAGEMENT_GOV"="Dynamics 365 for Case Management, Enterprise Edition for Government";
"Dynamics_365_Customer_Service_Enterprise_admin_trial"="Dynamics 365 Customer Service Enterprise Admin";
"DYN365_ENTERPRISE_CUSTOMER_SERVICE"="Dynamics 365 for Customer Service Enterprise Edition";
"DYN365_ENTERPRISE_CUSTOMER_SERVICE_GOV"="Dynamics 365 for Customer Service, Enterprise Edition for Government";
"D365_ENTERPRISE_CUSTOMER_SERVICE_GOV"="Dynamics 365 for Customer Service Enterprise for Government";
"DYN365_CS_CHAT"="Dynamics 365 for Customer Service Chat";
"D365_CUSTOMER_SERVICE_PRO_ATTACH"="Dynamics 365 for Customer Service Professional Attach to Qualifying Dynamics 365 Base Offer";
"D365_FIELD_SERVICE_ATTACH"="Dynamics 365 for Field Service Attach to Qualifying Dynamics 365 Base Offer";
"DYN365_ENTERPRISE_FIELD_SERVICE"="Dynamics 365 for Field Service Enterprise Edition";
"D365_FIELD_SERVICE_CONTRACTOR_GOV"="Dynamics 365 Field Service Contractor for Government";
"CRM_AUTO_ROUTING_ADDON"="Dynamics 365 Field Service, Enterprise Edition - Resource Scheduling Optimization";
"D365_ENTERPRISE_FIELD_SERVICE_GOV"="Dynamics 365 for Field Service for Government";
"DYN365_ENTERPRISE_FIELD_SERVICE_GOV"="Dynamics 365 for Field Service Enterprise Edition for Government";
"DYN365_FINANCIALS_BUSINESS_SKU"="Dynamics 365 for Financials Business Edition";
"Dynamics_365_Guides_vTrial"="Dynamics 365 Guides vTrial";
"CRM_HYBRIDCONNECTOR"="Dynamics 365 Hybrid Connector";
"DYN365_MARKETING_APPLICATION_ADDON"="Dynamics 365 for Marketing Additional Application";
"DYN365_MARKETING_CONTACT_ADDON_T3"="Dynamics 365 for Marketing Additional Contacts Tier 3";
"DYN365_MARKETING_SANDBOX_APPLICATION_ADDON"="Dynamics 365 for Marketing Additional Non-Prod Application";
"DYN365_MARKETING_CONTACT_ADDON_T5"="Dynamics 365 for Marketing Addnl Contacts Tier 5";
"DYN365_MARKETING_APP_ATTACH"="Dynamics 365 for Marketing Attach";
"DYN365_BUSINESS_MARKETING"="Dynamics 365 for Marketing Business Edition";
"D365_MARKETING_USER"="Dynamics 365 for Marketing USL";
"Dyn365_Operations_Activity"="Dynamics 365 Operations ? Activity";
"DYN365_PROJECT_OPERATIONS_ATTACH"="Dynamics 365 Project Operations Attach";
"D365_ENTERPRISE_PROJECT_SERVICE_AUTOMATION_GOV"="Dynamics 365 for Project Service Automation for Government";
"DYN365_ENTERPRISE_PROJECT_SERVICE_AUTOMATION_GOV"="Dynamics 365 for Project Service Automation Enterprise Edition for Government";
"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"="Dynamics 365 for Sales and Customer Service Enterprise Edition";
"DYN365_ENTERPRISE_SALES"="Dynamics 365 for Sales Enterprise Edition";
"DYN365_ENTERPRISE_SALES_GOV"="Dynamics 365 for Sales, Enterprise Edition for Government";
"D365_ENTERPRISE_SALES_GOV"="Dynamics 365 for Sales Enterprise for Government";
"Dynamics_365_Sales_Field_Service_and_Customer_Service_Partner_Sandbox"="Dynamics 365 Sales, Field Service and Customer Service Partner Sandbox";
"DYN365_SALES_PREMIUM"="Dynamics 365 Sales Premium";
"D365_SALES_PRO"="Dynamics 365 for Sales Professional";
"D365_SALES_PRO_GOV"="Dynamics 365 for Sales Professional for Government";
"D365_SALES_PRO_IW"="Dynamics 365 for Sales Professional Trial";
"D365_SALES_PRO_ATTACH"="Dynamics 365 for Sales Professional Attach to Qualifying Dynamics 365 Base Offer";
"DYN365_SCM"="Dynamics 365 for Supply Chain Management";
"SKU_Dynamics_365_for_HCM_Trial"="Dynamics 365 for Talent";
"DYN365_ENTERPRISE_TEAM_MEMBERS"="Dynamics 365 for Team Members Enterprise Edition";
"DYN365_ENTERPRISE_TEAM_MEMBERS_GOV"="Dynamics 365 for Team Members Enterprise Edition for Government";
"GUIDES_USER"="Dynamics 365 Guides";
"Dynamics_365_for_Operations_Devices"="Dynamics 365 Operations - Device";
"Dynamics_365_for_Operations_Sandbox_Tier2_SKU"="Dynamics 365 Operations - Sandbox Tier 2:Standard Acceptance Testing";
"Dynamics_365_for_Operations_Sandbox_Tier4_SKU"="Dynamics 365 Operations - Sandbox Tier 4:Standard Performance Testing";
"DYN365_ENTERPRISE_P1_IW"="Dynamics 365 P1 Trial for Information Workers";
"DYN365_PROJECT_OPERATIONS"="Dynamics 365 Project Operations";
"DYN365_REGULATORY_SERVICE"="Dynamics 365 Regulatory Service - Enterprise Edition Trial";
"MICROSOFT_REMOTE_ASSIST"="Dynamics 365 Remote Assist";
"MICROSOFT_REMOTE_ASSIST_HOLOLENS"="Dynamics 365 Remote Assist HoloLens";
"D365_SALES_ENT_ATTACH"="Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer";
"Dynamics_365_Sales_Premium_Viral_Trial"="Dynamics 365 Sales Premium Viral Trial";
"DYN365_SCM_ATTACH"="Dynamics 365 for Supply Chain Management Attach to Qualifying Dynamics 365 Base Offer";
"Dynamics_365_Hiring_SKU"="Dynamics 365 Talent: Attract";
"DYNAMICS_365_ONBOARDING_SKU"="Dynamics 365 Talent: Onboard";
"DYN365_TEAM_MEMBERS"="Dynamics 365 Team Members_wDynamicsRetail";
"Dynamics_365_for_Operations"="Dynamics 365 UNF OPS Plan ENT Edition";
"EMS_EDU_FACULTY"="Enterprise Mobility + Security A3 for Faculty";
"EMS"="Enterprise Mobility + Security E3";
"EMSPREMIUM"="Enterprise Mobility + Security E5";
"EMSPREMIUM_USGOV_GCCHIGH"="Enterprise Mobility + Security E5_USGOV_GCCHIGH";
"EMS_GOV"="Enterprise Mobility + Security G3 GCC";
"EMSPREMIUM_GOV"="Enterprise Mobility + Security G5 GCC";
"EOP_ENTERPRISE_PREMIUM"="Exchange Enterprise CAL Services (EOP, DLP)";
"EXCHANGESTANDARD"="Exchange Online (Plan 1)";
"EXCHANGESTANDARD_STUDENT"="Exchange Online (Plan 1) for Students";
"EXCHANGESTANDARD_ALUMNI"="Exchange Online (Plan 1) for Alumni with Yammer";
"EXCHANGEENTERPRISE"="Exchange Online (PLAN 2)";
"EXCHANGEENTERPRISE_FACULTY"="Exchange Online (Plan 2) for Faculty";
"EXCHANGEENTERPRISE_GOV"="Exchange Online (Plan 2) for GCC";
"EXCHANGEARCHIVE_ADDON"="Exchange Online Archiving for Exchange Online";
"EXCHANGEARCHIVE"="Exchange Online Archiving for Exchange Server";
"EXCHANGEESSENTIALS"="Exchange Online Essentials (ExO P1 Based)";
"EXCHANGE_S_ESSENTIALS"="Exchange Online Essentials";
"EXCHANGEDESKLESS"="Exchange Online Kiosk";
"EXCHANGESTANDARD_GOV"="Exchange Online (Plan 1) for GCC";
"EXCHANGETELCO"="Exchange Online POP";
"EOP_ENTERPRISE"="Exchange Online Protection";
"INTUNE_A"="Intune";
"INTUNE_EDU"="Intune for Education";
"AX7_USER_TRIAL"="Microsoft Dynamics AX7 User Trial";
"MFA_STANDALONE"="Microsoft Azure Multi-Factor Authentication";
"THREAT_INTELLIGENCE"="Microsoft Defender for Office 365 (Plan 2)";
"M365EDU_A1"="Microsoft 365 A1";
"M365EDU_A3_FACULTY"="Microsoft 365 A3 for faculty";
"M365EDU_A3_STUDENT"="Microsoft 365 A3 for students";
"M365EDU_A3_STUUSEBNFT"="Microsoft 365 A3 student use benefits";
"Microsoft_365_A3_Suite_features_for_faculty"="Microsoft 365 A3 Suite features for faculty";
"M365EDU_A3_STUUSEBNFT_RPA1"="Microsoft 365 A3 - Unattended License for students use benefit";
"M365EDU_A5_FACULTY"="Microsoft 365 A5 for Faculty";
"M365EDU_A5_STUDENT"="Microsoft 365 A5 for students";
"M365EDU_A5_STUUSEBNFT"="Microsoft 365 A5 student use benefits";
"M365_A5_SUITE_COMPONENTS_FACULTY"="Microsoft 365 A5 Suite features for faculty";
"M365EDU_A5_NOPSTNCONF_STUUSEBNFT"="Microsoft 365 A5 without Audio Conferencing for students use benefit";
"O365_BUSINESS"="Microsoft 365 Apps for Business";
"SMB_BUSINESS"="Microsoft 365 Apps for Business";
"OFFICESUBSCRIPTION"="Microsoft 365 Apps for enterprise";
"OFFICE_PROPLUS_DEVICE1"="Microsoft 365 Apps for enterprise (device)";
"OFFICESUBSCRIPTION_FACULTY"="Microsoft 365 Apps for Faculty";
"OFFICESUBSCRIPTION_STUDENT"="Microsoft 365 Apps for Students";
"MCOMEETADV_GOV"="Microsoft 365 Audio Conferencing for GCC";
"MCOACBYOT_AR_GCCHIGH_USGOV_GCCHIGH"="Microsoft 365 Audio Conferencing - GCCHigh Tenant (AR)_USGOV_GCCHIGH";
"MCOMEETADV_USGOV_GCCHIGH"="Microsoft 365 Audio Conferencing_USGOV_GCCHIGH";
"MCOMEETACPEA"="Microsoft 365 Audio Conferencing Pay-Per-Minute - EA";
"O365_BUSINESS_ESSENTIALS"="Microsoft 365 Business Basic";
"SMB_BUSINESS_ESSENTIALS"="Microsoft 365 Business Basic";
"Microsoft_365_Business_Basic_EEA_(no_Teams)"="Microsoft 365 Business Basic EEA (no Teams)";
"O365_BUSINESS_PREMIUM"="Microsoft 365 Business Standard";
"Microsoft_365_Business_Standard_EEA_(no_Teams)"="Microsoft 365 Business Standard EEA (no Teams)";
"Office_365_w/o_Teams_Bundle_Business_Standard"="Microsoft 365 Business Standard EEA (no Teams)";
"SMB_BUSINESS_PREMIUM"="Microsoft 365 Business Standard - Prepaid Legacy";
"SPB"="Microsoft 365 Business Premium";
"Office_365_w/o_Teams_Bundle_Business_Premium"="Microsoft 365 Business Premium EEA (no Teams)";
"BUSINESS_VOICE_MED2_TELCO"="Microsoft 365 Business Voice (US)";
"MCOPSTN_5"="Microsoft 365 Domestic Calling Plan (120 Minutes)";
"MCOPSTN_1_GOV"="Microsoft 365 Domestic Calling Plan for GCC";
"SPE_E3"="Microsoft 365 E3";
"O365_w/o Teams Bundle_M3"="Microsoft 365 E3 EEA (no Teams)";
"Microsoft_365_E3_EEA_(no_Teams)_Unattended_License"="Microsoft 365 E3 EEA (no Teams) - Unattended License";
"O365_w/o Teams Bundle_M3_(500_seats_min)_HUB"="Microsoft 365 E3 EEA (no Teams) (500 seats min)_HUB";
"Microsoft_365_E3_Extra_Features"="Microsoft 365 E3 Extra Features";
"SPE_E3_RPA1"="Microsoft 365 E3 - Unattended License";
"Microsoft_365_E3"="Microsoft 365 E3 (500 seats min)_HUB";
"SPE_E3_USGOV_DOD"="Microsoft 365 E3_USGOV_DOD";
"SPE_E3_USGOV_GCCHIGH"="Microsoft 365 E3_USGOV_GCCHIGH";
"SPE_E5"="Microsoft 365 E5";
"Microsoft_365_E5"="Microsoft 365 E5 (500 seats min)_HUB";
"DEVELOPERPACK_E5"="Microsoft 365 E5 Developer (without Windows and Audio Conferencing)";
"INFORMATION_PROTECTION_COMPLIANCE"="Microsoft 365 E5 Compliance";
"O365_w/o_Teams_Bundle_M5"="Microsoft 365 E5 EEA (no Teams)";
"O365_w/o_Teams_Bundle_M5_(500_seats_min)_HUB"="Microsoft 365 E5 EEA (no Teams) (500 seats min)_HUB";
"Microsoft_365_E5_EEA_(no_Teams)_with_Calling_Minutes"="Microsoft 365 E5 EEA (no Teams) with Calling Minutes";
"Microsoft_365_E5_EEA_(no_Teams)_without_Audio_Conferencing"="Microsoft 365 E5 EEA (no Teams) without Audio Conferencing";
"Microsoft_365_E5_EEA_(no_Teams)without_Audio_Conferencing(500_seats_min)_HUB"="Microsoft 365 E5 EEA (no Teams) without Audio Conferencing (500 seats min)_HUB";
"IDENTITY_THREAT_PROTECTION"="Microsoft 365 E5 Security";
"IDENTITY_THREAT_PROTECTION_FOR_EMS_E5"="Microsoft 365 E5 Security for EMS E5";
"SPE_E5_CALLINGMINUTES"="Microsoft 365 E5 with Calling Minutes";
"SPE_E5_NOPSTNCONF"="Microsoft 365 E5 without Audio Conferencing";
"Microsoft_365_E5_without_Audio_Conferencing"="Microsoft 365 E5 without Audio Conferencing (500 seats min)_HUB";
"M365_F1"="Microsoft 365 F1";
"Microsoft_365_F1_EEA_(no_Teams)"="Microsoft 365 F1 EEA (no Teams)";
"SPE_F1"="Microsoft 365 F3";
"Microsoft_365_F3_EEA_(no_Teams)"="Microsoft 365 F3 EEA (no Teams)";
"SPE_F5_COMP"="Microsoft 365 F5 Compliance Add-on";
"SPE_F5_COMP_AR_D_USGOV_DOD"="Microsoft 365 F5 Compliance Add-on AR DOD_USGOV_DOD";
"SPE_F5_COMP_AR_USGOV_GCCHIGH"="Microsoft 365 F5 Compliance Add-on AR_USGOV_GCCHIGH";
"SPE_F5_COMP_GCC"="Microsoft 365 F5 Compliance Add-on GCC";
"SPE_F5_SEC"="Microsoft 365 F5 Security Add-on";
"SPE_F5_SECCOMP"="Microsoft 365 F5 Security + Compliance Add-on";
"FLOW_FREE"="Microsoft Power Automate Free";
"M365_E5_SUITE_COMPONENTS"="Microsoft 365 E5 Extra Features";
"M365_F1_COMM"="Microsoft 365 F1";
"SPE_E5_USGOV_GCCHIGH"="Microsoft 365 E5_USGOV_GCCHIGH";
"M365_F1_GOV"="Microsoft 365 F3 GCC";
"M365_G3_GOV"="Microsoft 365 G3 GCC";
"M365_G3_RPA1_GOV"="Microsoft 365 G3 - Unattended License for GCC";
"M365_G5_GCC"="Microsoft 365 GCC G5";
"M365_G5_GOV"="Microsoft 365 GCC G5 w/o WDATP/CAS Unified";
"Microsoft365_Lighthouse"="Microsoft 365 Lighthouse";
"M365_SECURITY_COMPLIANCE_FOR_FLW"="Microsoft 365 Security and Compliance for Firstline Workers";
"MICROSOFT_BUSINESS_CENTER"="Microsoft Business Center";
"Microsoft_Cloud_for_Sustainability_vTrial"="Microsoft Cloud for Sustainability vTrial";
"ADALLOM_STANDALONE"="Microsoft Cloud App Security";
"Microsoft_365_Copilot"="Microsoft Copilot for Microsoft 365";
"WIN_DEF_ATP"="Microsoft Defender for Endpoint";
"Microsoft_Defender_for_Endpoint_F2"="Microsoft Defender for Endpoint F2";
"DEFENDER_ENDPOINT_P1"="Microsoft Defender for Endpoint P1";
"DEFENDER_ENDPOINT_P1_EDU"="Microsoft Defender for Endpoint P1 for EDU";
"MDATP_XPLAT"="Microsoft Defender for Endpoint P2_XPLAT";
"MDATP_Server"="Microsoft Defender for Endpoint Server";
"ATP_ENTERPRISE_FACULTY"="Microsoft Defender for Office 365 (Plan 1) Faculty";
"CRMPLAN2"="Microsoft Dynamics CRM Online Basic";
"ATA"="Microsoft Defender for Identity";
"ATP_ENTERPRISE_GOV"="Microsoft Defender for Office 365 (Plan 1) GCC";
"ATP_ENTERPRISE_USGOV_GCCHIGH"="Microsoft Defender for Office 365 (Plan 1)_USGOV_GCCHIGH";
"THREAT_INTELLIGENCE_GOV"="Microsoft Defender for Office 365 (Plan 2) GCC";
"TVM_Premium_Standalone"="Microsoft Defender Vulnerability Management";
"TVM_Premium_Add_on"="Microsoft Defender Vulnerability Management Add-on";
"CRMSTANDARD"="Microsoft Dynamics CRM Online";
"CRMPLAN2_GCC"="Microsoft Dynamics CRM Online Basic for Government";
"CRMSTANDARD_GCC"="Microsoft Dynamics CRM Online for Government";
"Microsoft_Entra_ID_Governance"="Microsoft Entra ID Governance";
"POWER_BI_STANDARD"="Microsoft Fabric (Free)";
"POWER_BI_STANDARD_FACULTY"="Microsoft Fabric (Free) for faculty";
"POWER_BI_STANDARD_STUDENT"="Microsoft Fabric (Free) for student";
"IT_ACADEMY_AD"="Microsoft Imagine Academy";
"INTUNE_A_D"="Microsoft Intune Device";
"INTUNE_A_D_GOV"="Microsoft Intune Device for Government";
"INTUNE_A_GOV"="Microsoft Intune Government";
"INTUNE_A_VL_USGOV_GCCHIGH"="Microsoft Intune Plan 1 A VL_USGOV_GCCHIGH";
"Microsoft_Intune_Suite"="Microsoft Intune Suite";
"POWERAPPS_VIRAL"="Microsoft Power Apps Plan 2 Trial";
"FLOW_P2"="Microsoft Power Automate Plan 2";
"INTUNE_SMB"="Microsoft Intune SMB";
"POWERAPPS_DEV"="Microsoft PowerApps for Developer";
"POWERFLOW_P2"="Microsoft Power Apps Plan 2 (Qualified Offer)";
"DYN365_ENTERPRISE_RELATIONSHIP_SALES"="Microsoft Relationship Sales solution";
"STREAM"="Microsoft Stream";
"STREAM_P2"="Microsoft Stream Plan 2";
"STREAM_STORAGE"="Microsoft Stream Storage Add-On (500 GB)";
"Microsoft_Cloud_for_Sustainability_USL"="Microsoft Sustainability Manager USL Essentials";
"Microsoft_Teams_Audio_Conferencing_select_dial_out"="Microsoft Teams Audio Conferencing with dial-out to USA/CAN";
"TEAMS_FREE"="Microsoft Teams (Free)";
"Microsoft_Teams_Calling_Plan_pay_as_you_go_(country_zone_1_US)"="Microsoft Teams Calling Plan pay-as-you-go (country zone 1 - US)";
"MCOPSTN_6"="Microsoft Teams Domestic Calling Plan (240 min)";
"Microsoft_Teams_EEA_New"="Microsoft Teams EEA";
"Teams_Ess"="Microsoft Teams Essentials";
"TEAMS_ESSENTIALS_AAD"="Microsoft Teams Essentials (Microsoft Entra identity)";
"TEAMS_EXPLORATORY"="Microsoft Teams Exploratory";
"MCOEV"="Microsoft Teams Phone Standard";
"MCOEV_DOD"="Microsoft Teams Phone Standard for DOD";
"MCOEV_FACULTY"="Microsoft Teams Phone Standard for Faculty";
"MCOEV_GOV"="Microsoft Teams Phone Standard for GCC";
"MCOEV_GCCHIGH"="Microsoft Teams Phone Standard for GCCHIGH";
"MCOEVSMB_1"="Microsoft Teams Phone Standard for Small and Medium Business";
"MCOEV_STUDENT"="Microsoft Teams Phone Standard for Student";
"MCOEV_TELSTRA"="Microsoft Teams Phone Standard for TELSTRA";
"MCOEV_USGOV_DOD"="Microsoft Teams Phone Standard_System_USGOV_DOD";
"MCOEV_USGOV_GCCHIGH"="Microsoft Teams Phone Standard_USGOV_GCCHIGH";
"PHONESYSTEM_VIRTUALUSER"="Microsoft Teams Phone Resoure Account";
"PHONESYSTEM_VIRTUALUSER_GOV"="Microsoft Teams Phone Resource Account for GCC";
"PHONESYSTEM_VIRTUALUSER_USGOV_GCCHIGH"="Microsoft Teams Phone Resource Account_USGOV_GCCHIGH";
"Microsoft_Teams_Premium"="Microsoft Teams Premium Introductory Pricing";
"Microsoft_Teams_Rooms_Basic"="Microsoft Teams Rooms Basic";
"Microsoft_Teams_Rooms_Basic_FAC"="Microsoft Teams Rooms Basic for EDU";
"Microsoft_Teams_Rooms_Basic_without_Audio_Conferencing"="Microsoft Teams Rooms Basic without Audio Conferencing";
"Microsoft_Teams_Rooms_Pro"="Microsoft Teams Rooms Pro";
"Microsoft_Teams_Rooms_Pro_FAC"="Microsoft Teams Rooms Pro for EDU";
"Microsoft_Teams_Rooms_Pro_GCC"="Microsoft Teams Rooms Pro for GCC";
"Microsoft_Teams_Rooms_Pro_without_Audio_Conferencing"="Microsoft Teams Rooms Pro without Audio Conferencing";
"MEETING_ROOM"="Microsoft Teams Rooms Standard";
"MEETING_ROOM_GOV"="Microsoft Teams Rooms Standard for GCC";
"MEETING_ROOM_GOV_NOAUDIOCONF"="Microsoft Teams Rooms Standard for GCC without Audio Conferencing";
"MCOCAP"="Microsoft Teams Shared Devices";
"MCOCAP_GOV"="Microsoft Teams Shared Devices for GCC";
"MS_TEAMS_IW"="Microsoft Teams Trial";
"EXPERTS_ON_DEMAND"="Microsoft Threat Experts - Experts on Demand";
"WORKPLACE_ANALYTICS"="Microsoft Workplace Analytics";
"Microsoft_Viva_Goals"="Microsoft Viva Goals";
"Viva_Glint_Standalone"="Microsoft Viva Glint";
"VIVA"="Microsoft Viva Suite";
"MEE_FACULTY"="Minecraft Education Faculty";
"MEE_STUDENT"="Minecraft Education Student";
"OFFICE365_MULTIGEO"="Multi-Geo Capabilities in Office 365";
"NONPROFIT_PORTAL"="Nonprofit Portal";
"STANDARDWOFFPACK_FACULTY"="Office 365 A1 for Faculty";
"STANDARDWOFFPACK_IW_FACULTY"="Office 365 A1 Plus for Faculty";
"STANDARDWOFFPACK_STUDENT"="Office 365 A1 for Students";
"STANDARDWOFFPACK_IW_STUDENT"="Office 365 A1 Plus for Students";
"ENTERPRISEPACKPLUS_FACULTY"="Office 365 A3 for Faculty";
"ENTERPRISEPACKPLUS_STUDENT"="Office 365 A3 for Students";
"ENTERPRISEPREMIUM_FACULTY"="Office 365 A5 for faculty";
"ENTERPRISEPREMIUM_STUDENT"="Office 365 A5 for students";
"EQUIVIO_ANALYTICS"="Office 365 Advanced Compliance";
"ATP_ENTERPRISE"="Microsoft Defender for Office 365 (Plan 1)";
"SHAREPOINTSTORAGE_GOV"="Office 365 Extra File Storage for GCC";
"TEAMS_COMMERCIAL_TRIAL"="Microsoft Teams Commercial Cloud";
"ADALLOM_O365"="Office 365 Cloud App Security";
"SHAREPOINTSTORAGE"="Office 365 Extra File Storage";
"STANDARDPACK"="Office 365 E1";
"Office_365_w/o_Teams_Bundle_E1"="Office 365 E1 EEA (no Teams)";
"STANDARDWOFFPACK"="Office 365 E2";
"ENTERPRISEPACK"="Office 365 E3";
"O365_w/o_Teams_Bundle_E3"="Office 365 E3 EEA (no Teams)";
"DEVELOPERPACK"="Office 365 E3 Developer";
"ENTERPRISEPACK_USGOV_DOD"="Office 365 E3_USGOV_DOD";
"ENTERPRISEPACK_USGOV_GCCHIGH"="Office 365 E3_USGOV_GCCHIGH";
"ENTERPRISEWITHSCAL"="Office 365 E4";
"ENTERPRISEPREMIUM"="Office 365 E5";
"Office_365_w/o_Teams_Bundle_E5"="Office 365 E5 EEA (no Teams)";
"Office_365_E5_EEA_(no_Teams)_without_Audio_Conferencing"="Office 365 E5 EEA (no Teams) without Audio Conferencing";
"ENTERPRISEPREMIUM_NOPSTNCONF"="Office 365 E5 without Audio Conferencing";
"DESKLESSPACK"="Office 365 F3";
"Office_365_F3_EEA_(no_Teams)"="Office 365 F3 EEA (no Teams)";
"DESKLESSPACK_USGOV_GCCHIGH"="Office 365 F3_USGOV_GCCHIGH";
"STANDARDPACK_GOV"="Office 365 G1 GCC";
"ENTERPRISEPACK_GOV"="Office 365 G3 GCC";
"ENTERPRISEPACKWITHOUTPROPLUS_GOV"="Office 365 G3 without Microsoft 365 Apps GCC";
"ENTERPRISEPREMIUM_GOV"="Office 365 G5 GCC";
"ENTERPRISEPREMIUM_NOPBIPBX_GOV"="Office 365 GCC G5 without Power BI and Phone System";
"ENTERPRISEPREMIUM_NOPSTNCONF_NOPBI_GOV"="Office 365 GCC G5 without Audio Conferencing";
"EQUIVIO_ANALYTICS_GOV"="Office 365 Advanced Compliance for GCC";
"MIDSIZEPACK"="Office 365 Midsize Business";
"LITEPACK"="Office 365 Small Business";
"LITEPACK_P2"="Office 365 Small Business Premium";
"OFFICEMOBILE_SUBSCRIPTION_GOV_TEST"="Office Mobile Apps for Office 365 for GCC";
"WACONEDRIVESTANDARD"="OneDrive for Business (Plan 1)";
"WACONEDRIVEENTERPRISE"="OneDrive for Business (Plan 2)";
"POWERFLOWGCC_TEST"="PowerApps & Flow GCC Test - O365 & Dyn365 Plans";
"POWERAPPS_INDIVIDUAL_USER"="Power Apps and Logic Flows";
"POWERAPPS_PER_APP_IW"="Power Apps per app baseline access";
"POWERAPPS_PER_APP"="Power Apps per app Plan";
"POWERAPPS_PER_APP_NEW"="Power Apps per app Plan (1 app or portal)";
"POWERAPPS_PER_APP_BD_ONLY_GCC"="Power Apps Per App BD Only for GCC";
"Power_Apps_per_app_plan_(1_app_or_portal)_BD_Only_GCC"="Power Apps per app plan (1 app or website) BD Only ? GCC";
"POWERAPPS_PER_APP_GCC_NEW"="Power Apps per app plan (1 app or website) for Government";
"POWERAPPS_PER_APP_GCC"="Power Apps per app plan for Government";
"POWERAPPS_PER_USER_BD_ONLY"="Power Apps Per User BD Only";
"POWERAPPS_PER_USER"="Power Apps per user Plan";
"POWERAPPS_PER_USER_GCC"="Power Apps per user Plan for Government";
"POWERAPPS_P1_GOV"="Power Apps Plan 1 for Government";
"POWERAPPS_PORTALS_LOGIN_T2"="Power Apps Portals login capacity add-on Tier 2 (10 unit min)";
"POWERAPPS_PORTALS_LOGIN_T2_GCC"="Power Apps Portals login capacity add-on Tier 2 (10 unit min) for Government";
"POWERAPPS_PORTALS_LOGIN_T3"="Power Apps Portals login capacity add-on Tier 3 (50 unit min)";
"POWERAPPS_PORTALS_PAGEVIEW"="Power Apps Portals page view capacity add-on";
"POWERAPPS_PORTALS_PAGEVIEW_GCC"="Power Apps Portals page view capacity add-on for Government";
"FLOW_BUSINESS_PROCESS"="Power Automate per flow plan";
"FLOW_BUSINESS_PROCESS_GCC"="Power Automate per flow plan for Government";
"FLOW_PER_USER"="Power Automate per user plan";
"FLOW_PER_USER_DEPT"="Power Automate per user plan dept";
"FLOW_PER_USER_GCC"="Power Automate per user plan for Government";
"POWERAUTOMATE_ATTENDED_RPA"="Power Automate per user with attended RPA plan";
"FLOW_P1_GOV"="Power Automate Plan 1 for Government (Qualified Offer)";
"POWERAUTOMATE_ATTENDED_RPA_GCC"="Power Automate Premium for Government";
"POWERAUTOMATE_UNATTENDED_RPA"="Power Automate unattended RPA add-on";
"POWERAUTOMATE_UNATTENDED_RPA_GCC"="Power Automate unattended RPA add-on for Government";
"POWER_BI_INDIVIDUAL_USER"="Power BI";
"POWER_BI_ADDON"="Power BI for Office 365 Add-On";
"PBI_PREMIUM_P1_ADDON"="Power BI Premium P1";
"PBI_PREMIUM_P1_ADDON_GCC"="Power BI Premium P1 GCC";
"PBI_PREMIUM_PER_USER"="Power BI Premium Per User";
"PBI_PREMIUM_PER_USER_ADDON"="Power BI Premium Per User Add-On";
"PBI_PREMIUM_PER_USER_ADDON_GCC"="Power BI Premium Per User Add-On for GCC";
"PBI_PREMIUM_PER_USER_ADDON_CE_GCC"="Power BI Premium Per User Add-On for GCC";
"PBI_PREMIUM_PER_USER_FACULTY"="Power BI Premium Per User for Faculty";
"PBI_PREMIUM_PER_USER_DEPT"="Power BI Premium Per User Dept";
"PBI_PREMIUM_PER_USER_GCC"="Power BI Premium Per User for Government";
"POWER_BI_PRO"="Power BI Pro";
"POWER_BI_PRO_CE"="Power BI Pro CE";
"POWER_BI_PRO_DEPT"="Power BI Pro Dept";
"POWER_BI_PRO_FACULTY"="Power BI Pro for Faculty";
"POWERBI_PRO_GOV"="Power BI Pro for GCC";
"Power_Pages_authenticated_users_T1_100_users/per_site/month_capacity_pack"="Power Pages authenticated users T1 100 users/per site/month capacity pack";
"Power Pages authenticated users T1_CN_CN"="Power Pages authenticated users T1 100 users/per site/month capacity pack CN_CN";
"Power_Pages_authenticated_users_T1_100_users/per_site/month_capacity_pack_GCC"="Power Pages authenticated users T1 100 users/per site/month capacity pack_GCC";
"Power_Pages_authenticated_users_T1_100_users/per_site/month_capacity_pack_USGOV_DOD"="Power Pages authenticated users T1 100 users/per site/month capacity pack_USGOV_DOD";
"Power_Pages_authenticated_users_T1_100_users/per_site/month_capacity_pack_USGOV_GCCHIGH"="Power Pages authenticated users T1 100 users/per site/month capacity pack_USGOV_GCCHIGH";
"Power_Pages_authenticated_users_T2_min_100_units_100_users/per_site/month_capacity_pack"="Power Pages authenticated users T2 min 100 units - 100 users/per site/month capacity pack";
"Power Pages authenticated users T2_CN_CN"="Power Pages authenticated users T2 min 100 units - 100 users/per site/month capacity pack CN_CN";
"Power_Pages_authenticated_users_T2_min_100_units_100_users/per_site/month_capacity_pack_GCC"="Power Pages authenticated users T2 min 100 units - 100 users/per site/month capacity pack_GCC";
"Power_Pages_authenticated_users_T2_min_100_units_100_users/per_site/month_capacity_pack_USGOV_DOD"="Power Pages authenticated users T2 min 100 units - 100 users/per site/month capacity pack_USGOV_DOD";
"Power_Pages_authenticated_users_T2_min_100_units_100_users/per_site/month_capacity_pack_USGOV_GCCHIGH"="Power Pages authenticated users T2 min 100 units - 100 users/per site/month capacity pack_USGOV_GCCHIGH";
"Power_Pages_authenticated_users_T3_min_1,000_units_100_users/per_site/month_capacity_pack"="Power Pages authenticated users T3 min 1,000 units - 100 users/per site/month capacity pack";
"Power Pages authenticated users T3_CN_CN"="Power Pages authenticated users T3 min 1,000 units - 100 users/per site/month capacity pack CN_CN";
"Power_Pages_authenticated_users_T3_min_1,000_units_100_users/per_site/month_capacity_pack_GCC"="Power Pages authenticated users T3 min 1,000 units - 100 users/per site/month capacity pack_GCC";
"Power_Pages_authenticated_users_T3_min_1,000_units_100_users/per_site/month_capacity_pack_USGOV_DOD"="Power Pages authenticated users T3 min 1,000 units - 100 users/per site/month capacity pack_USGOV_DOD";
"Power_Pages_authenticated_users_T3_min_1,000_units_100_users/per_site/month_capacity_pack_USGOV_GCCHIGH"="Power Pages authenticated users T3 min 1,000 units - 100 users/per site/month capacity pack_USGOV_GCCHIGH";
"Power_Pages_vTrial_for_Makers"="Power Pages vTrial for Makers";
"VIRTUAL_AGENT_BASE"="Power Virtual Agent";
"VIRTUAL_AGENT_BASE_GCC"="Power Virtual Agent for GCC";
"VIRTUAL_AGENT_USL"="Power Virtual Agent User License";
"VIRTUAL_AGENT_USL_GCC"="Power Virtual Agent User License for GCC";
"CCIBOTS_PRIVPREV_VIRAL"="Power Virtual Agents Viral Trial";
"PRIVACY_MANAGEMENT_RISK"="Privacy Management ? risk";
"PRIVACY_MANAGEMENT_RISK_EDU"="Privacy Management - risk for EDU";
"PRIVACY_MANAGEMENT_RISK_GCC"="Privacy Management - risk GCC";
"PRIVACY_MANAGEMENT_RISK_USGOV_DOD"="Privacy Management - risk_USGOV_DOD";
"PRIVACY_MANAGEMENT_RISK_USGOV_GCCHIGH"="Privacy Management - risk_USGOV_GCCHIGH";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_1_V2"="Privacy Management - subject rights request (1)";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_1_EDU_V2"="Privacy Management - subject rights request (1) for EDU";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_1_V2_GCC"="Privacy Management - subject rights request (1) GCC";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_1_V2_USGOV_DOD"="Privacy Management - subject rights request (1) USGOV_DOD";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_1_V2_USGOV_GCCHIGH"="Privacy Management - subject rights request (1) USGOV_GCCHIGH";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_10_V2"="Privacy Management - subject rights request (10)";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_10_EDU_V2"="Privacy Management - subject rights request (10) for EDU";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_10_V2_GCC"="Privacy Management - subject rights request (10) GCC";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_10_V2_USGOV_DOD"="Privacy Management - subject rights request (10) USGOV_DOD";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_10_V2_USGOV_GCCHIGH"="Privacy Management - subject rights request (10) USGOV_GCCHIGH";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_50"="Privacy Management - subject rights request (50)";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_50_V2"="Privacy Management - subject rights request (50)";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_50_EDU_V2"="Privacy Management - subject rights request (50) for EDU";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_100_V2"="Privacy Management - subject rights request (100)";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_100_EDU_V2"="Privacy Management - subject rights request (100) for EDU";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_100_V2_GCC"="Privacy Management - subject rights request (100) GCC";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_100_V2_USGOV_DOD"="Privacy Management - subject rights request (100) USGOV_DOD";
"PRIVACY_MANAGEMENT_SUB_RIGHTS_REQ_100_V2_USGOV_GCCHIGH"="Privacy Management - subject rights request (100) USGOV_GCCHIGH";
"PROJECTCLIENT"="Project for Office 365";
"PROJECTESSENTIALS"="Project Online Essentials";
"PROJECTESSENTIALS_FACULTY"="Project Online Essentials for Faculty";
"PROJECTESSENTIALS_GOV"="Project Online Essentials for GCC";
"PROJECTPREMIUM"="Project Online Premium";
"PROJECTONLINE_PLAN_1"="Project Online Premium without Project Client";
"PROJECTONLINE_PLAN_2"="Project Online with Project for Office 365";
"PROJECT_P1"="Project Plan 1";
"PROJECT_PLAN1_DEPT"="Project Plan 1 (for Department)";
"PROJECTPROFESSIONAL"="Project Plan 3";
"PROJECT_PLAN3_DEPT"="Project Plan 3 (for Department)";
"PROJECTPROFESSIONAL_FACULTY"="Project Plan 3 for Faculty";
"PROJECTPROFESSIONAL_GOV"="Project Plan 3 for GCC";
"Project_Professional_TEST_GCC"="Project Plan 3 for GCC TEST";
"PROJECTPROFESSIONAL_USGOV_GCCHIGH"="Project Plan 3_USGOV_GCCHIGH";
"PROJECTPREMIUM_FACULTY"="Project Plan 5 for faculty";
"PROJECTPREMIUM_GOV"="Project Plan 5 for GCC";
"PROJECTONLINE_PLAN_1_FACULTY"="Project Plan 5 without Project Client for Faculty";
"RIGHTSMANAGEMENT_ADHOC"="Rights Management Adhoc";
"RMSBASIC"="Rights Management Service Basic Content Protection";
"DYN365_IOT_INTELLIGENCE_ADDL_MACHINES"="Sensor Data Intelligence Additional Machines Add-in for Dynamics 365 Supply Chain Management";
"DYN365_IOT_INTELLIGENCE_SCENARIO"="Sensor Data Intelligence Scenario Add-in for Dynamics 365 Supply Chain Management";
"SHAREPOINTSTANDARD"="SharePoint Online (Plan 1)";
"SHAREPOINTENTERPRISE"="SharePoint Online (Plan 2)";
"Intelligent_Content_Services"="SharePoint Syntex";
"MCOIMP"="Skype for Business Online (Plan 1)";
"MCOSTANDARD"="Skype for Business Online (Plan 2)";
"MCOPSTN2"="Skype for Business PSTN Domestic and International Calling";
"MCOPSTN1"="Skype for Business PSTN Domestic Calling";
"MCOPSTN5"="Skype for Business PSTN Domestic Calling (120 Minutes)";
"MCOPSTNPP"="Skype for Business PSTN Usage Calling Plan";
"Operator_Connect_Mobile"="Teams Phone Mobile";
"MCOTEAMS_ESSENTIALS"="Teams Phone with Calling Plan";
"Teams_Premium_(for_Departments)"="Teams Premium (for Departments)";
"MCOPSTNEAU2"="TELSTRA Calling for O365";
"UNIVERSAL_PRINT"="Universal Print";
"VISIO_PLAN1_DEPT"="Visio Plan 1";
"VISIO_PLAN2_DEPT"="Visio Plan 2";
"VISIOONLINE_PLAN1"="Visio Online Plan 1";
"VISIOCLIENT_FACULTY"="Visio Plan 2 for Faculty";
"VISIOCLIENT"="Visio Online Plan 2";
"VISIOCLIENT_GOV"="Visio Plan 2 for GCC";
"VISIOCLIENT_USGOV_GCCHIGH"="Visio Plan 2_USGOV_GCCHIGH";
"Viva_Goals_User_led"="Viva Goals User-led";
"TOPIC_EXPERIENCES"="Viva Topics";
"WIN_ENT_E5"="Windows 10/11 Enterprise E5 (Original)";
"WIN10_ENT_A3_FAC"="Windows 10/11 Enterprise A3 for faculty";
"WIN10_ENT_A3_STU"="Windows 10/11 Enterprise A3 for students";
"WIN10_ENT_A5_FAC"="Windows 10/11 Enterprise A5 for faculty";
"WIN10_PRO_ENT_SUB"="WINDOWS 10/11 ENTERPRISE E3";
"WIN10_VDA_E3"="WINDOWS 10/11 ENTERPRISE E3";
"WIN10_VDA_E5"="Windows 10/11 Enterprise E5";
"WINE5_GCC_COMPAT"="Windows 10/11 Enterprise E5 Commercial (GCC Compatible)";
"E3_VDA_only"="Windows 10/11 Enterprise VDA";
"CPC_B_1C_2RAM_64GB"="Windows 365 Business 1 vCPU 2 GB 64 GB";
"CPC_B_2C_4RAM_128GB"="Windows 365 Business 2 vCPU 4 GB 128 GB";
"CPC_B_2C_4RAM_256GB"="Windows 365 Business 2 vCPU 4 GB 256 GB";
"CPC_B_2C_4RAM_64GB"="Windows 365 Business 2 vCPU 4 GB 64 GB";
"CPC_B_2C_8RAM_128GB"="Windows 365 Business 2 vCPU 8 GB 128 GB";
"CPC_B_2C_8RAM_256GB"="Windows 365 Business 2 vCPU 8 GB 256 GB";
"CPC_B_4C_16RAM_128GB"="Windows 365 Business 4 vCPU 16 GB 128 GB";
"CPC_B_4C_16RAM_128GB_WHB"="Windows 365 Business 4 vCPU 16 GB 128 GB (with Windows Hybrid Benefit)";
"CPC_B_4C_16RAM_256GB"="Windows 365 Business 4 vCPU 16 GB 256 GB";
"CPC_B_4C_16RAM_512GB"="Windows 365 Business 4 vCPU 16 GB 512 GB";
"CPC_B_8C_32RAM_128GB"="Windows 365 Business 8 vCPU 32 GB 128 GB";
"CPC_B_8C_32RAM_256GB"="Windows 365 Business 8 vCPU 32 GB 256 GB";
"CPC_B_8C_32RAM_512GB"="Windows 365 Business 8 vCPU 32 GB 512 GB";
"CPC_E_1C_2GB_64GB"="Windows 365 Enterprise 1 vCPU 2 GB 64 GB";
"CPC_E_2C_4GB_128GB"="Windows 365 Enterprise 2 vCPU 4 GB 128 GB";
"CPC_E_2C_4GB_256GB"="Windows 365 Enterprise 2 vCPU 4 GB 256 GB";
"CPC_E_2C_4GB_64GB"="Windows 365 Enterprise 2 vCPU 4 GB 64 GB";
"CPC_E_2C_8GB_128GB"="Windows 365 Enterprise 2 vCPU 8 GB 128 GB";
"CPC_E_2C_8GB_256GB"="Windows 365 Enterprise 2 vCPU 8 GB 256 GB";
"CPC_E_4C_16GB_128GB"="Windows 365 Enterprise 4 vCPU 16 GB 128 GB";
"CPC_E_4C_16GB_256GB"="Windows 365 Enterprise 4 vCPU 16 GB 256 GB";
"CPC_E_4C_16GB_512GB"="Windows 365 Enterprise 4 vCPU 16 GB 512 GB";
"CPC_E_8C_32GB_128GB"="Windows 365 Enterprise 8 vCPU 32 GB 128 GB";
"CPC_E_8C_32GB_256GB"="Windows 365 Enterprise 8 vCPU 32 GB 256 GB";
"CPC_E_8C_32GB_512GB"="Windows 365 Enterprise 8 vCPU 32 GB 512 GB";
"CPC_LVL_1"="Windows 365 Enterprise 2 vCPU 4 GB 128 GB (Preview)";
"Windows_365_S_2vCPU_4GB_64GB"="Windows 365 Shared Use 2 vCPU 4 GB 64 GB";
"Windows_365_S_2vCPU_4GB_128GB"="Windows 365 Shared Use 2 vCPU 4 GB 128 GB";
"Windows_365_S_2vCPU_4GB_256GB"="Windows 365 Shared Use 2 vCPU 4 GB 256 GB";
"Windows_365_S_2vCPU_8GB_128GB"="Windows 365 Shared Use 2 vCPU 8 GB 128 GB";
"Windows_365_S_2vCPU_8GB_256GB"="Windows 365 Shared Use 2 vCPU 8 GB 256 GB";
"Windows_365_S_4vCPU_16GB_128GB"="Windows 365 Shared Use 4 vCPU 16 GB 128 GB";
"Windows_365_S_4vCPU_16GB_256GB"="Windows 365 Shared Use 4 vCPU 16 GB 256 GB";
"Windows_365_S_4vCPU_16GB_512GB"="Windows 365 Shared Use 4 vCPU 16 GB 512 GB";
"Windows_365_S_8vCPU_32GB_128GB"="Windows 365 Shared Use 8 vCPU 32 GB 128 GB";
"Windows_365_S_8vCPU_32GB_256GB"="Windows 365 Shared Use 8 vCPU 32 GB 256 GB";
"Windows_365_S_8vCPU_32GB_512GB"="Windows 365 Shared Use 8 vCPU 32 GB 512 GB";
"WINDOWS_STORE"="Windows Store for Business";
"WSFB_EDU_FACULTY"="Windows Store for Business EDU Faculty";
"Workload_Identities_Premium_CN"="Workload Identities Premium"}

Add-Type -Assembly System.Drawing

#List of possible colors for data points
$arrColors = new-object System.Collections.ArrayList

[VOID]$arrColors.add([System.Drawing.Color]::DeepSkyBlue)
[VOID]$arrColors.add([System.Drawing.Color]::BlueViolet)
[VOID]$arrColors.add([System.Drawing.Color]::Turquoise)
[VOID]$arrColors.add([System.Drawing.Color]::DarkTurquoise)
[VOID]$arrColors.add([System.Drawing.Color]::LimeGreen)
[VOID]$arrColors.add([System.Drawing.Color]::Aquamarine)
[VOID]$arrColors.add([System.Drawing.Color]::Aqua)
[VOID]$arrColors.add([System.Drawing.Color]::SpringGreen)
[VOID]$arrColors.add([System.Drawing.Color]::SteelBlue)
[VOID]$arrColors.add([System.Drawing.Color]::Navy)
[VOID]$arrColors.add([System.Drawing.Color]::Teal)
[VOID]$arrColors.add([System.Drawing.Color]::MidnightBlue)
[VOID]$arrColors.add([System.Drawing.Color]::LightSlateGray)
[VOID]$arrColors.add([System.Drawing.Color]::LightSteelBlue)
[VOID]$arrColors.add([System.Drawing.Color]::DimGray)
[VOID]$arrColors.add([System.Drawing.Color]::DarkCyan)
[VOID]$arrColors.add([System.Drawing.Color]::Magenta)
80..300 |%{[VOID]$arrColors.Add([System.Drawing.Color]::$(([System.Drawing.Color] | gm -Static -MemberType Properties)[$_].Name))}

#$VerbosePreference = "Continue"
#List of possible colors for data points
$arrPrivGraphColors = new-object System.Collections.ArrayList

[VOID]$arrPrivGraphColors.add([System.Drawing.Color]::Crimson)


#==========================================================================
# Function		: New-DoughnutChartMutpleDataPoints
# Arguments     : Chart Object, Serie Name, Legend Text, CSV Data,Background color, Color 1, Color 1, Number of Chart Areas in the same Chart Object
# Returns   	: 
# Description   : Draw Doughnut Chart Object in Chart Area
#==========================================================================
Function New-DoughnutChartMutpleDataPoints
{
    param(
    $chart1,$Legend,$CSV,[string]$BackColor,$arrColors,$ChartCount)
    
    $SerieName = "Serie1"
    $Arial = new-object System.Drawing.FontFamily("Arial")
    $Font = new-object System.Drawing.Font($Arial,12 ,[System.Drawing.FontStyle]::Bold)

    $ChartCounterPosition = $chart1.ChartAreas.count
    $ChartElementPosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),12,((100/($ChartCount*0.98))-5),100)
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = $SerieName
    $chartarea.Position = $ChartElementPosition
    $chartarea.BackColor = $BackColor
    [void]$chart1.ChartAreas.Add($chartarea)


    if($Legend -ne "")
    {
        $NewTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $NewTitle.Text = $Legend
        $NewTitle.Name = "ChartTitle"+$chart1.Titles.count
        $TitlePosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),(100-(135-(2.5*$ChartCount))),((100/($ChartCount*0.98))-5),100)
        $NewTitle.Position= $TitlePosition
        [void]$Chart1.Titles.Add($NewTitle);
        $chart1.Titles[($chart1.Titles.count-1)].Font = $Font 
        $chart1.Titles[($chart1.Titles.count-1)].ForeColor = [System.Drawing.Color]::White
        $Chart1.Titles[($chart1.Titles.count-1)].DockedToChartArea = $chartarea.Name
    }

        

    [void]$chart1.Series.Add($SerieName)
    $chart1.Series[$SerieName].ChartType = "Doughnut"
    $chart1.Series[$SerieName].SetCustomProperty("DoughnutRadius","50")
    $chart1.Series[$SerieName].IsVisibleInLegend = $true
    $chart1.Series[$SerieName].chartarea = $chartarea.Name
 
    $chart1.Series[$SerieName].LabelForeColor = [System.Drawing.Color]::White
    $chart1.Series[$SerieName].BorderColor = $BackColor
    $chart1.Series[$SerieName].BorderWidth = 5
    $chart1.Series[$SerieName].Font = $Font
    $chart1.Series[$SerieName].IsValueShownAsLabel = $true
    $chart1.Series[$SerieName].IsXValueIndexed = $true

    $colHeaders = ( $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')

    $i = 0    
    $Total = 0
    #For all headers create a data point
    Foreach ($ColumnName in $colHeaders )
    {
        $CSV | ForEach-Object {
            $Point = $chart1.Series[$SerieName].Points.addxy( "$ColumnName" , ($_.$ColumnName)) 
            $Total = $Total + $_.$ColumnName
        }
        $chart1.Series[$SerieName].Points[$Point].Color = $arrColors[$i]
        if($i -ge $($arrColors.count -1))
        {
            $i = 0
        }
        else {
            $i++
        }
        
    }    


    $TextAnno = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
    $TextAnno.Text = $Total
    $TextAnno.Width = ((100/($ChartCount*0.98))-5)
    $TextAnno.X = (0+($ChartCounterPosition * ((100/$ChartCount))))
    $TextAnno.Y = 53
    $TextAnno.Font = "Segoe UI Black,20pt"
    $TextAnno.ForeColor = [System.Drawing.Color]::White
    $TextAnno.BringToFront()
    [void]$chart1.Annotations.Add($TextAnno)
    

}
#==========================================================================
# Function		: Create-ChartDoughnutMutpleDataPoints
# Arguments     : Chart Title, picuture file, CSV data , Background Color, Doughnut Color
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream.
# Requires      : Function New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Create-ChartDoughnutMutpleDataPoints {
    param([string]$ChartTitle,$TitlelinGraph,$picturefile,$CSV,[string]$Backcolor,$arrColors,$ChartCount = 2)

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    ## Chart Object 
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 200 *$ChartCount
    $chart1.Height = (240 + (5 * $ChartCount))
    $chart1.BackColor = $Backcolor
    ## Title
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,20pt"
    $chart1.Titles[0].Alignment = "topLeft"
    $chart1.Titles[0].ForeColor = [System.Drawing.Color]::White


    ## Legend 
    $Legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.name = "Legend1"
    $Legend.Font = "Arial,12pt"
    $Legend.ForeColor = [System.Drawing.Color]::White
    $Legend.BackColor = 0
    $Legend.MaximumAutoSize = 100
    $Legend.IsDockedInsideChartArea = $false
    $Legend.TextWrapThreshold = 55
    $Legend.Alignment = [System.Drawing.StringAlignment]::Near

    $chart1.Legends.Add($Legend)
    ## Data Series

    New-DoughnutChartMutpleDataPoints $chart1 $TitlelinGraph $CSV $Backcolor $arrColors $ChartCount
       
    # Save Chart
    $chart1.SaveImage($picturefile,"png")

}

#==========================================================================
# Function		: Add-DoughnutGraph
# Arguments     : Adds doughnut graph as a picuture file from CSV data 
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream input the data in a HTML string
# Requires      : Create-ChartDoughnutMutpleDataPoints,New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Add-DoughnutGraph
{
    param($Data,$GraphTitle,$DoughnutTitle,$arrColors, $ChartCount = 2, $BackColor = "#131313")
$PNGFileName = New-Object System.IO.MemoryStream


#Adapting the width of the chart to fit the text and legend
$DataLegends = $data | Get-Member -MemberType noteproperty | Select-Object -property name |  Sort-Object { $_.name.length } -Descending:$true

foreach($legend in $DataLegends)
{
    $length = $legend.name.tostring().length
    if($length -gt 30)
    {
        $ChartCount = 2 + (($length - 18)*.04)
        break
    }    
    elseif($length -gt 18)
    {
        $ChartCount = 2 + (($length - 18)*0.048)
        break
    }
    else
    {
        $ChartCount = 2 + (($length - 18)*0.03)
    }
}

Create-ChartDoughnutMutpleDataPoints -Backcolor $BackColor -ChartTitle $GraphTitle -TitlelinGraph $DoughnutTitle -picturefile $PNGFileName -CSV $Data  -arrColors $arrColors -ChartCount $ChartCount 

$Stat = [convert]::ToBase64String($PNGFileName.ToArray())


$strHTMLGraph = "<img src=""data:image/png;base64,$Stat "" />"

return $strHTMLGraph

}
#==========================================================================
# Function		: New-DoughnutChartMutpleDataPoints
# Arguments     : Chart Object, Serie Name, Legend Text, CSV Data,Background color, Color 1, Color 1, Number of Chart Areas in the same Chart Object
# Returns   	: 
# Description   : Draw Doughnut Chart Object in Chart Area
#==========================================================================
Function New-BigDoughnutChartMutpleDataPoints
{
    param(
    $chart1,$Legend,$CSV,[string]$BackColor,$arrColors,$ChartCount)
    
    $SerieName = "Serie1"
    $Arial = new-object System.Drawing.FontFamily("Arial")
    $Font = new-object System.Drawing.Font($Arial,12 ,[System.Drawing.FontStyle]::Bold)

    $ChartCounterPosition = $chart1.ChartAreas.count
    $ChartElementPosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),12,((100/($ChartCount*0.98))-5),100)
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = $SerieName
    $chartarea.Position = $ChartElementPosition
    $chartarea.BackColor = $BackColor
    [void]$chart1.ChartAreas.Add($chartarea)


    if($Legend -ne "")
    {
        $NewTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $NewTitle.Text = $Legend
        $NewTitle.Name = "ChartTitle"+$chart1.Titles.count
        $TitlePosition = new-object System.Windows.Forms.DataVisualization.Charting.ElementPosition((0+($ChartCounterPosition * ((100/$ChartCount)))),(100-(135-(2.5*$ChartCount))),((100/($ChartCount*0.98))-5),100)
        $NewTitle.Position= $TitlePosition
        [void]$Chart1.Titles.Add($NewTitle);
        $chart1.Titles[($chart1.Titles.count-1)].Font = $Font 
        $chart1.Titles[($chart1.Titles.count-1)].ForeColor = [System.Drawing.Color]::White
        $Chart1.Titles[($chart1.Titles.count-1)].DockedToChartArea = $chartarea.Name
    }

        

    [void]$chart1.Series.Add($SerieName)
    $chart1.Series[$SerieName].ChartType = "Doughnut"
    $chart1.Series[$SerieName].SetCustomProperty("DoughnutRadius","50")
    $chart1.Series[$SerieName].IsVisibleInLegend = $true
    $chart1.Series[$SerieName].chartarea = $chartarea.Name
 
    $chart1.Series[$SerieName].LabelForeColor = [System.Drawing.Color]::White
    $chart1.Series[$SerieName].BorderColor = $BackColor
    $chart1.Series[$SerieName].BorderWidth = 5
    $chart1.Series[$SerieName].Font = $Font
    $chart1.Series[$SerieName].IsValueShownAsLabel = $true
    $chart1.Series[$SerieName].IsXValueIndexed = $true

    $colHeaders = ( $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')

    $i = 0    
    $Total = 0
    #For all headers create a data point
    Foreach ($ColumnName in $colHeaders )
    {
        $CSV | ForEach-Object {
            $Point = $chart1.Series[$SerieName].Points.addxy( "$ColumnName" , ($_.$ColumnName)) 
            $Total = $Total + $_.$ColumnName
        }
        $chart1.Series[$SerieName].Points[$Point].Color = $arrColors[$i]
        if($i -ge $($arrColors.count -1))
        {
            $i = 0
        }
        else {
            $i++
        }
        
    }    


    $TextAnno = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
    $TextAnno.Text = $Total
    $TextAnno.Width = ((100/($ChartCount*0.98))-5)
    $TextAnno.X = (0+($ChartCounterPosition * ((100/$ChartCount))))
    $TextAnno.Y = 58
    $TextAnno.Font = "Segoe UI Black,20pt"
    $TextAnno.ForeColor = [System.Drawing.Color]::White
    $TextAnno.BringToFront()
    [void]$chart1.Annotations.Add($TextAnno)
    

}
#==========================================================================
# Function		: Create-BigChartDoughnutMutpleDataPoints
# Arguments     : Chart Title, picuture file, CSV data , Background Color, Doughnut Color
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream.
# Requires      : Function New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Create-BigChartDoughnutMutpleDataPoints {
    param([string]$ChartTitle,$TitlelinGraph,$picturefile,$CSV,[string]$Backcolor,$arrColors,$ChartCount = 2)

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")


    ## Chart Object 
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 800
    $chart1.Height = 600
    $chart1.BackColor = $Backcolor
    ## Title
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,20pt"
    $chart1.Titles[0].Alignment = "topLeft"
    $chart1.Titles[0].ForeColor = [System.Drawing.Color]::White


    ## Legend 
    $Legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.name = "Legend1"
    $Legend.Font = "Arial,12pt"
    $Legend.ForeColor = [System.Drawing.Color]::White
    $Legend.BackColor = $Backcolor
    $Legend.MaximumAutoSize = 100
    $Legend.IsDockedInsideChartArea = $false
    $Legend.TextWrapThreshold = 30
    $Legend.Alignment = [System.Drawing.StringAlignment]::Near

    $chart1.Legends.Add($Legend)
    ## Data Series

    New-BigDoughnutChartMutpleDataPoints $chart1 $TitlelinGraph $CSV $Backcolor $arrColors -ChartCount $ChartCount
       
    # Save Chart
    $chart1.SaveImage($picturefile,"png")

}

#==========================================================================
# Function		: Add-BigDoughnutGraph
# Arguments     : Adds doughnut graph as a picuture file from CSV data 
# Returns   	: 
# Description   : Create Chart Object and save to png file/filestream input the data in a HTML string
# Requires      : Create-ChartDoughnutMutpleDataPoints,New-DoughnutChartMutpleDataPoints
#==========================================================================
Function Add-BigDoughnutGraph
{
    param($Data,$GraphTitle,$DoughnutTitle,$ChartCount = 2,$arrColors,$Backcolor = "#131313")
$PNGFileName = New-Object System.IO.MemoryStream


Create-BigChartDoughnutMutpleDataPoints $GraphTitle $DoughnutTitle $PNGFileName $Data $Backcolor $arrColors -ChartCount $ChartCount 

$Stat = [convert]::ToBase64String($PNGFileName.ToArray())


$strHTMLGraph = "<img src=""data:image/png;base64,$Stat "" />"

return $strHTMLGraph

}
#==========================================================================
# Function		: Parse-JWToken
# Arguments     : Token as string
# Returns   	: 
# Description   : Decodes a JWT token. This was taken from link below. Thanks to Vasil Michev.
#==========================================================================
function Parse-JWToken {
    <#
    .DESCRIPTION
    Decodes a JWT token. This was taken from link below. Thanks to Vasil Michev.
    .LINK
    https://www.michev.info/Blog/Post/2140/decode-jwt-access-and-id-tokens-via-powershell
    #>
    param(
        [Parameter(Mandatory = $True)]
        [string]$Token
    )

    #Validate as per https://tools.ietf.org/html/rfc7519
    #Access and ID tokens are fine, Refresh tokens will not work
    if (-not $Token.Contains(".") -or -not $Token.StartsWith("eyJ")) {
        Write-Error "Invalid token" -ErrorAction Stop
    }
 
    #Header
    $tokenheader = $Token.Split(".")[0].Replace('-', '+').Replace('_', '/')

    #Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenheader.Length % 4) {
        Write-Verbose "Invalid length for a Base-64 char array or string, adding ="
        $tokenheader += "="
    }

    Write-Verbose "Base64 encoded (padded) header: $tokenheader"

    #Convert from Base64 encoded string to PSObject all at once
    Write-Verbose "Decoded header:"
    $header = ([System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($tokenheader)) | convertfrom-json)
 
    #Payload
    $tokenPayload = $Token.Split(".")[1].Replace('-', '+').Replace('_', '/')

    #Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenPayload.Length % 4) {
        Write-Verbose "Invalid length for a Base-64 char array or string, adding ="
        $tokenPayload += "="
    }
    
    Write-Verbose "Base64 encoded (padded) payoad: $tokenPayload"

    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)


    $tokenArray = ([System.Text.Encoding]::ASCII.GetString($tokenByteArray) | ConvertFrom-Json)

    #Converts $header and $tokenArray from PSCustomObject to Hashtable so they can be added together.
    #I would like to use -AsHashTable in convertfrom-json. This works in pwsh 6 but for some reason Appveyor isnt running tests in pwsh 6.
    $headerAsHash = @{}
    $tokenArrayAsHash = @{}
    $header.psobject.properties | ForEach-Object { $headerAsHash[$_.Name] = $_.Value }
    $tokenArray.psobject.properties | ForEach-Object { $tokenArrayAsHash[$_.Name] = $_.Value }
    $output = $headerAsHash + $tokenArrayAsHash

    Return $output

    
}

function Get-EntraIDToken {

    ############################################################################

    <#
    .SYNOPSIS

        Get an access token for use with the API cmdlets.


    .DESCRIPTION

        Uses MSAL.ps to obtain an access token. Has an option to refresh an existing token.


    .EXAMPLE

        Get-EntraIDToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f.


    .EXAMPLE

        Get-EntraIDToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -ForceRefresh

        Refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f.


    .EXAMPLE

        Get-EntraIDToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -LoginHint Bob@Contoso.com

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f and user Bob@Contoso.com.


    .EXAMPLE

        Get-EntraIDToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -InterActive

        Gets or refreshes an access token for making API calls for the tenant ID
        b446a536-cb76-4360-a8bb-6593cf4d9c7f. Ensures a pop-up box appears.

    #>

    ############################################################################

    [CmdletBinding(DefaultParameterSetName="InterActive")]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [guid]$TenantId,

        #Force a token refresh
        [Parameter(Position=1,ParameterSetName="ForceRefresh")]
        [switch]$ForceRefresh,

        #The user's upn used for the login hint
        [Parameter(Position=2,ParameterSetName="InterActive")]
        [string]$LoginHint,

        #Force a pop-up box
        [Parameter(Position=3,ParameterSetName="InterActive")]
        [switch]$InterActive,

        #get an Entra ID Graph token
        [Parameter(Position=4)]
        [switch]$AadGraph,

        #get an Entra ID Graph token
        [Parameter(Position=5)]
        [String[]]$Scopes
    )


    ############################################################################


    #Get an access token using the PowerShell client ID
    #$ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894" 
    $ClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
    #$RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $Authority = "https://login.microsoftonline.com/$TenantId"

    if ($AadGraph) {

        $Scopes = "https://graph.windows.net/.default"

    }
    else {
    
        $Scopes = "https://graph.microsoft.com/.default"

    }
    

    if ($ForceRefresh) {

        Write-Verbose -Message "$(Get-Date -f T) - Attempting to refresh an existing access token"

        #Attempt to refresh access token
        try {

            $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Scopes $Scopes -ForceRefresh

        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token refreshed - new expiry: $(($Response).ExpiresOn.UtcDateTime)"

            return $Response

        }
        else {
            
            Write-Warning -Message "$(Get-Date -f T) - Failed to refresh Access Token - try re-running the cmdlet again"

        }

    }
    elseif ($LoginHint) {

        Write-Verbose -Message "$(Get-Date -f T) - Checking token cache with -LoginHint for $LoginHint"

        #Run this to obtain an access token - should prompt on first run to select the account used for future operations
        try {

            if ($InterActive) {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -LoginHint $LoginHint -Scopes $Scopes -Interactive 
                
            } 
            else {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -LoginHint $LoginHint -Scopes $Scopes 

            }
        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained for: $(($Response).Account.Username) ($(($Response).Account.HomeAccountId.ObjectId))"
            #Write-Verbose -Message "$(Get-Date -f T) - API Access Token scopes: $(($Response).Scopes)"

            return $Response

        }
        else {

            Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
            Write-Warning -Message "$(Get-Date -f T) - If the problem persists, use `$Error[0] for more detail on the error or start a new PowerShell session"

        }

    }
    else {

        Write-Verbose -Message "$(Get-Date -f T) - Checking token cache with -Prompt"

        #Run this to obtain an access token - should prompt on first run to select the account used for future operations
        try {

            if ($InterActive) {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount -Interactive -Scopes $Scopes 

            }
            else {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount -Scopes $Scopes 

            }

        }
        catch {}

        #Error handling for token acquisition
        if ($Response) {

            Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained for: $(($Response).Account.Username) ($(($Response).Account.HomeAccountId.ObjectId))"
            #Write-Verbose -Message "$(Get-Date -f T) - API Access Token scopes: $(($Response).Scopes)"

            return $Response

        }
        else {

            Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
            Write-Warning -Message "$(Get-Date -f T) - If the problem persists, run Connect-AzureADIR with the -UserUpn parameter"

        }

    }


}   


function Get-AzureADIRHeader {

    ############################################################################

    <#
    .SYNOPSIS

        Uses a supplied Access Token to construct a header for a an API call.


    .DESCRIPTION

        Uses a supplied Access Token to construct a header for a an API call with 
        Invoke-WebRequest.

        Can supply the ConsistencyLevel = Eventual parameter for performing Count
        activities.


    .EXAMPLE

        Get-AzureADIRHeader -Token $Token

        Constructs a header with an obtained token for using with Invoke-WebRequest.


    .EXAMPLE

        Get-AzureADIRHeader -Token $Token -ConsistencyLevelEventual

        Constructs a header with an obtained token for using with Invoke-WebRequest.

        Uses the optional -ConsistencyLevelEventual switch for use in conjunction with
        the count call.

    #>

    ############################################################################
    
    [CmdletBinding()]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [string]$Token,

        #Switch to include ConsistencyLevel = Eventual for $count operations
        [Parameter(Position=1)]
        [switch]$ConsistencyLevelEventual

        )

    ############################################################################

    if ($ConsistencyLevelEventual) {

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";
                "ConsistencyLevel" = "eventual";

        }

    }
    else {

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";

        }

    }

}   #end function

function Get-EntraIDSKUs {
    ############################################################################

   <#
   .SYNOPSIS

       Get the subscriptions and the licences for the Entra ID Tenant


   .DESCRIPTION

       Returns the licences in the tenant

   .EXAMPLE

       Get-EntraIDSKU $Token $Header $TenantID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header

   )
   
   


   Write-Verbose -Message "$(Get-Date -f T) - Looking up role definitions for all roles"

   #All Users 

   $Url = "https://graph.microsoft.com/v1.0/subscribedSkus?`&`$Select=subscriptionIds,skuPartNumber,consumedUnits,prepaidUnits,servicePlans"
   
   try {

       # Convert the content in the response from Json and expand all values
      $skus = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {
       $StatusCode = [int]$_.Exception.Response.StatusCode
       if ($StatusCode -eq 400) {
           Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
       } elseif ($StatusCode -eq 500) {
           Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
       } else {
           Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
       }
   }


   
$Subscriptions = New-Object System.Collections.ArrayList


#$skus = $results.value | Select-Object subscriptionIds, skuPartNumber, consumedUnits, prepaidUnits, servicePlans

Foreach ($sku in $Skus) {

    $AADP1 = $sku.servicePlans | Where-Object { $_.servicePlanId -like "41781fb2-bc02-4b7c-bd55-b576c07bb09d" }
    $AADP2 = $sku.servicePlans | Where-Object { $_.servicePlanId -like "eec0eb4f-6444-4f95-aba0-50c24d67f998" }


    If ($null -eq $AADP1.provisioningStatus) {
        $StatusAADP1 = "n/a"
    }
    else {
        $StatusAADP1 = $AADP1.provisioningStatus
    }

    If ($null -eq $AADP2.provisioningStatus) {
        $StatusAADP2 = "n/a"
    }
    else {
        $StatusAADP2 = $AADP2.provisioningStatus
    }

    [string]$temp = $sku.subscriptionIds

    $details = @{
    
        SubscriptionId   = $temp
        SubscriptionName = $sku.skuPartNumber
        TotalLicenses    = $sku.prepaidUnits.enabled
        AssignedLicenses = $sku.consumedUnits
        EntraIDP1      = $StatusAADP1
        EntraIDP2      = $StatusAADP2

    }
     
    [VOID]$Subscriptions.add((New-Object psobject -Property $details))
}

   Return $Subscriptions 


} # End Function

function Get-RoleDefinitionsRoleManagement {
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-RoleDefinitionsRoleManagement $Token $Header $TenantID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up role definitions for all roles"

   #All Users 

   $Url = "https://graph.microsoft.com/beta/roleManagement/directory/roleDefinitions?"
      
   try {

       # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

    $StatusCode = [int]$_.Exception.Response.StatusCode
    if ($StatusCode -eq 400) {
        Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
    } elseif ($StatusCode -eq 403) {
        Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
        Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
    } elseif ($StatusCode -eq 500) {
            Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
    } else {
        Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
    }
}

   Return $ResponseData 


} # End Function

function Get-PIMRoleAssignments{
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-PIMRoleAssignments $Token $Header $TenantID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
        [Parameter(Mandatory)]
        $Token,

        [Parameter(Mandatory)]
        $Header,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $TenantID

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up role assignments for all roles"

   #All Users 

   $Url = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?`$expand=*"
   
   try {

       # Convert the content in the response from Json and expand all values
      $roleEligibilityScheduleInstances = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    #Get all role assignments by principalID
    $Url = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$select=principalId"

    try {
 
        # Convert the content in the response from Json and expand all values
       $principalIds= (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value
 
    }    
     
    catch {
 
         $StatusCode = [int]$_.Exception.Response.StatusCode
         if ($StatusCode -eq 400) {
             Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
         } elseif ($StatusCode -eq 403) {
             Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
             Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
         } elseif ($StatusCode -eq 500) {
                 Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
         } else {
             Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
         }
     }     

    $CollectionroleAssignmentScheduleInstances = New-Object System.Collections.ArrayList

    foreach($principalId in  $(($principalIds | Select-Object -Property principalID -Unique).principalID))
    {
        $Url = "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignmentScheduleInstances?`$filter=principalid eq '" + $principalId + "'&`$expand=*"
    
        try {
    
            # Convert the content in the response from Json and expand all values
        $roleAssignmentScheduleInstances = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value
    
        }    
        
        catch {
    
            $StatusCode = [int]$_.Exception.Response.StatusCode
            if ($StatusCode -eq 400) {
                Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
            } elseif ($StatusCode -eq 403) {
                Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
                Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
            } elseif ($StatusCode -eq 500) {
                    Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
            } else {
                Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
            }
        }   

        if($roleAssignmentScheduleInstances)
        {
            foreach($roleAssignmentScheduleInstance in $roleAssignmentScheduleInstances)
            {
                [VOID]$CollectionroleAssignmentScheduleInstances.add($roleAssignmentScheduleInstance)
            }

        }

    }

     $ResponseData = $CollectionroleAssignmentScheduleInstances + @($roleEligibilityScheduleInstances)

     if ($ResponseData) {
         Return $ResponseData 
     } 


} # End Function


function Get-RoleEligibilityScheduleRequests{
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-RoleEligibilityScheduleRequests $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
        [Parameter(Mandatory)]
        $Token,

        [Parameter(Mandatory)]
        $Header,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $TenantID,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $ObjectID
   
    )
    
    
   $ResponseData = $null



   $Url = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleRequests?`$filter=(Id%20eq%20'" + $ObjectID + "')"
   
   try {

       # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    if ($ResponseData) {
        Return $ResponseData 
    } 


} # End Function


function Get-RoleAssignmentScheduleRequests{
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-RoleAssignmentScheduleRequests $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
        [Parameter(Mandatory)]
        $Token,

        [Parameter(Mandatory)]
        $Header,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $TenantID,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $ObjectID
   
    )
    
    
   $ResponseData = $null



   $Url = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?`$filter=(Id%20eq%20'" + $ObjectID + "')"
   
   try {

       # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    if ($ResponseData) {
        Return $ResponseData 
    } 


} # End Function

function Get-RoleAssignments{
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-RoleAssignments $Token $Header $TenantID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up role assignments for all roles"

   #All Users 

   $Url = "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignments?"
   
   try {

       # Convert the content in the response from Json and expand all values
      $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction Stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   Return $ResponseData 


} # End Function

function Get-AdministrativeUnit {
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for Administrative Unit using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find administrativeunit

   .EXAMPLE

       Get-AdministrativeUnit $Token $Header $TenantID $ObjectID $Properties


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter()]
       [string]
       $ObjectID,

       [Parameter()]
       [string]
       $Properties
   )
   
   

    $ResponseData = $null
   

   #Search Administrative Units
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'&`$Select='$Properties'"
   }
   else {    
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "administrativeunit"
       Return $ResponseData 
   }    
} # End Function


function FindGlobalObject {
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGlobalObject $Token $Header $TenantID $ObjectID $Properties


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter()]
       [string]
       $ObjectID,

       [Parameter()]
       [string]
       $Properties
   )
   
   

    $ResponseData = $null
   
    #All Users 
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/users?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(givenName,'$ObjectID') OR startsWith(surName,'$ObjectID') OR startsWith(mail,'$ObjectID') OR startsWith(userPrincipalName,'$ObjectID') OR id eq '$ObjectID'&`$Select='$Properties'"
    }
    else {
        $Url = "https://graph.microsoft.com/beta/users?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(givenName,'$ObjectID') OR startsWith(surName,'$ObjectID') OR startsWith(mail,'$ObjectID') OR startsWith(userPrincipalName,'$ObjectID') OR id eq '$ObjectID'"     
    }
   
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   
   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "User"
       Return $ResponseData 
   }

   #All Groups
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(mail,'$ObjectID')&`$Select='$Properties'"
   }
   else {   
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=startsWith(displayName,'$ObjectID') OR startsWith(mail,'$ObjectID')"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }
    
   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "Group"
       Return $ResponseData 
   }


   #All Groups ID
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=id eq '$ObjectID'&`$Select='$Properties'"
   }
   else {      
        $Url = "https://graph.microsoft.com/beta/groups?`$filter=id eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }


   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "Group"        
       
        Return $ResponseData 
   }


    #All Applications
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/applications?`$filter=startsWith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'&`$Select='$Properties'"
    }
    else {    
        $Url = "https://graph.microsoft.com/beta/applications?`$filter=startsWith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'"
    }
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "application"
       Return $ResponseData 
   }

    #All Devices
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=startswith(displayName,'$ObjectID')&`$Select='$Properties'"
    }
    else {     
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=startswith(displayName,'$ObjectID')"
    }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }

   #Device DeviceID
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=deviceId eq '$ObjectID'&`$Select='$Properties'"
    }
    else {    
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=deviceId eq '$ObjectID'"
    }
    try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }

   #Device id
    if($Properties)
    {
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=id eq '$ObjectID'&`$Select='$Properties'"
    }
    else {       
        $Url = "https://graph.microsoft.com/beta/devices?`$filter=id eq '$ObjectID'"
    }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "device"
       Return $ResponseData 
   }
   
   #Role TemplateId
   if($Properties)
   {
    $Url = "https://graph.microsoft.com/beta/roleManagement/directory/roleDefinitions?"+"$"+"filter=templateID+eq+'$ObjectID'&`$Select='$Properties'"
   }
   else {     
        $Url = "https://graph.microsoft.com/beta/roleManagement/directory/roleDefinitions?"+"$"+"filter=templateID+eq+'$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       Return $ResponseData 
   }
   
   ##Search  all ServicePrincipals
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/servicePrincipals?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'&`$Select='$Properties'"
   }
   else {     
    $Url = "https://graph.microsoft.com/beta/servicePrincipals?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID' OR appId eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "application"
       Return $ResponseData 
   }

   #Search Administrative Units
   if($Properties)
   {
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'&`$Select='$Properties'"
   }
   else {    
        $Url = "https://graph.microsoft.com/beta/administrativeUnits?"+"$"+"filter=startswith(displayName,'$ObjectID') OR id eq '$ObjectID'"
   }
   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 404) {
            #Do nothing, the object is simply not found.              
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

   if ($ResponseData) {
       add-member -InputObject $ResponseData -MemberType NoteProperty -Name "type" -Value "administrativeunit"
       Return $ResponseData 
   }    
} # End Function

function Get-GroupOwner {
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-GroupOwner $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter(Mandatory)]
       $ObjectID

   )
   
   
   $ResponseData = $null

   #Get Owner
    $Url = "https://graph.microsoft.com/beta/groups/$ObjectID/owners?" + "$"+" orderby=displayName asc&"+"$"+"count=true"
    try {

        # Convert the content in the response from Json and expand all values
        $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }
    ####################### - Owner
   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function

function Get-GroupMembers {
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       FindGloGet-GroupMembersbalObject $Token $Header $TenantID $ObjectID


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header,

       [Parameter(Mandatory)]
       $TenantID,

       [Parameter(Mandatory)]
       $ObjectID

   )
   
   
   $ResponseData = $null

   #Get Owner
    $Url = "https://graph.microsoft.com/beta/groups/$ObjectID/transitiveMembers?" +"$"+" orderby=displayName asc&"+"$"+"count=true"
    try {

        # Convert the content in the response from Json and expand all values
        $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }
    ####################### - Owner
   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function

function Get-GroupPIMStatus {
    ############################################################################
  
   <#
   .SYNOPSIS
  
       Search Entra ID for a group and check if it's PIM enabled
  
  
   .DESCRIPTION
  
       Function to return the PIM activations status for a group
  
   .EXAMPLE
  
       Get-GroupPIMStatus $Token $Header $TenantID $ObjectID
  
  
   #>
  
   ############################################################################
  
    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,
  
       [Parameter(Mandatory)]
       $Header,
  
       [Parameter(Mandatory=$true, 
       ValueFromPipeline=$true,
       ValueFromPipelineByPropertyName=$true, 
       ValueFromRemainingArguments=$false, 
       ParameterSetName='Default')]
       [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]
       [string] 
       $TenantID,
  
       [Parameter(Mandatory=$true, 
       ValueFromPipeline=$true,
       ValueFromPipelineByPropertyName=$true, 
       ValueFromRemainingArguments=$false, 
       ParameterSetName='Default')]
       [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]
       [string] 
       $ObjectID
  
   )
   
   
   $ResponseData = $null

    $url = "https://graph.microsoft.com/beta/privilegedAccess/aadGroups/resources?`$filter=(Id%20eq%20'" + $ObjectID + "')"
  
    try {
  
        # Convert the content in the response from Json and expand all values
        $ResponseData = ((Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value).status
  
    }    
    catch {
  
          $StatusCode = [int]$_.Exception.Response.StatusCode
          if ($StatusCode -eq 400) {
              Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
          } elseif ($StatusCode -eq 403) {
              Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
              Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red           
          } elseif ($StatusCode -eq 500) {
                  Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
          } else {
              Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
          }
      }
  
   if ($ResponseData) {
       Return $ResponseData 
   }
  
  } # End Function


function Get-PIMGroupMembers {
############################################################################

<#
.SYNOPSIS

    Search Entra ID for a group and its eligible or active members


.DESCRIPTION

    Function to get eligible or active members of a group

.EXAMPLE

    Get-PIMGroupMembers $Token $Header $TenantID $ObjectID


#>

############################################################################

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Token,

        [Parameter(Mandatory)]
        $Header,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $TenantID,

        [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $ObjectID

    )


    $ResponseData = $null


    # Get active assignments
    $Url = "https://graph.microsoft.com/beta/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '" + $ObjectID + "'"
    try {

        # Convert the content in the response from Json and expand all values
        $assignmentResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

            $StatusCode = [int]$_.Exception.Response.StatusCode
            if ($StatusCode -eq 400) {
                Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
            } elseif ($StatusCode -eq 403) {
                Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
                Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red           
            } elseif ($StatusCode -eq 500) {
                    Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
            } else {
                Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
            }
        }

    # Get active assignments
    $Url = "https://graph.microsoft.com/beta/identityGovernance/privilegedAccess/group/EligibilitySchedules?`$filter=groupId eq '" + $ObjectID + "'"
    try {

        # Convert the content in the response from Json and expand all values
        $eligibilityResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

            $StatusCode = [int]$_.Exception.Response.StatusCode
            if ($StatusCode -eq 400) {
                Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
            } elseif ($StatusCode -eq 403) {
                Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
                Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red           
            } elseif ($StatusCode -eq 500) {
                    Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
            } else {
                Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
            }
        }      

    $ResponseData = @($assignmentResponseData) + @($eligibilityResponseData)

    if ($ResponseData) {
        Return $ResponseData 
    } 

} # End Function



#==========================================================================
# Function		: ConvertTo-ObjectArrayListFromPsCustomObject  
# Arguments     : Defined Object
# Returns   	: Custom Object List
# Description   : Convert a defined object to a custom, this will help you  if you got a read-only object 
# 
#==========================================================================
function ConvertTo-ObjectArrayListFromPsCustomObject 
{ 
     param ( 
         [Parameter(  
             Position = 0,   
             Mandatory = $true,   
             ValueFromPipeline = $true,  
             ValueFromPipelineByPropertyName = $true  
         )] $psCustomObject
     ); 
     
     process {
 
        $myCustomArray = New-Object System.Collections.ArrayList
     
         foreach ($myPsObject in $psCustomObject) { 
             $hashTable = @{}; 
             $myPsObject | Get-Member -MemberType *Property | ForEach-Object { 
                 $hashTable.($_.name) = $myPsObject.($_.name); 
             } 
             $Newobject = new-object psobject -Property  $hashTable
             [void]$myCustomArray.add($Newobject)
         } 
         return $myCustomArray
     } 
 }# End function
 Write-Verbose -Message "$(Get-Date -f T) - Authenticate"
Function Get-TenantInformation
{
    ############################################################################

   <#
   .SYNOPSIS

       Search Entra ID for objects using ObjectID,Id,DeviceID,TemplateID 


   .DESCRIPTION

       Function to find user,group,application,device,application and administrativeunit

   .EXAMPLE

       Get-TenantInformation $Token $Header


   #>

   ############################################################################

    [CmdletBinding()]
   param (
       [Parameter(Mandatory)]
       $Token,

       [Parameter(Mandatory)]
       $Header

   )
   
   

   $ResponseData = $null

   Write-Verbose -Message "$(Get-Date -f T) - Looking up Tenant Information"

   #Get Domains 
   $Url = "https://graph.microsoft.com/beta/organization"

   try {

       # Convert the content in the response from Json and expand all values
       $ResponseData = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

   }    
   catch {

    $StatusCode = [int]$_.Exception.Response.StatusCode
    if ($StatusCode -eq 400) {
        Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
    } elseif ($StatusCode -eq 403) {
        Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
        Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
    } elseif ($StatusCode -eq 500) {
            Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
    } else {
        Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
    }
   }

   
   if ($ResponseData) {
       Return $ResponseData 
   }

} # End Function

Function Get-ModifiedRoleManagementPolicies
{
    ############################################################################

   <#
   .SYNOPSIS

       Return all role management policies that have once been edited 


   .DESCRIPTION

       Function to report  all role management policies that have once been edited 

   .EXAMPLE

       Get-ModifiedRoleManagementPolicies -Header $Header


   #>

   ############################################################################
    Param
    (        
 
        [Parameter(Mandatory)]
        $Header
        
     
    )

    $Url = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies?`$filter=scopeId eq '/' and scopeType eq 'Directory' and lastModifiedDateTime ne null"
    
    try {

    $ModifiedroleManagementPolicies = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    return $ModifiedroleManagementPolicies
  

}

Function Get-RoleManagementPolicyAssignments 
{
    ############################################################################

   <#
   .SYNOPSIS

       Return all role m anagement policy assignments 


   .DESCRIPTION

       Function to report  all role m anagement policy assignments 

   .EXAMPLE

       Get-RoleManagementPolicyAssignments -Header $Header


   #>

   ############################################################################
    Param
    (        
 
        [Parameter(Mandatory)]
        $Header
        
     
    )

    $Url = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '/' and scopeType eq 'Directory'"

    try {

    $roleManagementPolicyAssignments = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red             
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    return $roleManagementPolicyAssignments
  

}
Function Get-RoleManagementPolicySettings
{
    ############################################################################

   <#
   .SYNOPSIS

       Search for role definition id to check if the role requires an approval


   .DESCRIPTION

       Function report if the role requires an approval in the role management policy for a specific role using its role defintion id

   .EXAMPLE

       Get-RoleManagementPolicySettings -roleDefinitionsid 790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b


   #>

   ############################################################################
    Param
    (        
        # Entra ID Role Definition ID
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true, 
                    ValueFromRemainingArguments=$false, 
                    Position=0,
                    ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $policyId,
        [Parameter(Mandatory)]
        $Token,
        [Parameter(Mandatory)]
        $Header
        
     
    )

   $InternalError = $false

    $Url = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/$($policyId)/rules"

    try {

    $objroleManagementPolicyRules = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    }    
    catch {

        $StatusCode = [int]$_.Exception.Response.StatusCode
        if ($StatusCode -eq 400) {
            Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
        } elseif ($StatusCode -eq 403) {
            Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
            Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
        } elseif ($StatusCode -eq 500) {
                Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red   
                $InternalError = $true          
        } else {
            Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
        }
    }

    # If internal Error try again in verbose mode
    if($InternalError)
    {
        Write-Host "Second try in verbose mode" -ForegroundColor Yellow
        try {

            $objroleManagementPolicyRules = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$true -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value
        
            }    
            catch {
        
                $StatusCode = [int]$_.Exception.Response.StatusCode
                if ($StatusCode -eq 400) {
                    Write-Host "Error: The remote server returned an error: (400) Bad Request." -ForegroundColor Red 
                } elseif ($StatusCode -eq 403) {
                    Write-Host "Error: Access denied! You do not have access to the API: $($Url) in the Tenant: $($TenantID)" -ForegroundColor Red
                    Write-Host "Current Permissions: $((Parse-JWToken $Header.Authorization.split(" ")[-1]).scp)" -ForegroundColor Red
                } elseif ($StatusCode -eq 500) {
                        Write-Host "Error: InternalServerError: Something went wrong on the backend!" -ForegroundColor Red 
                        Write-host "This Uri could not be executed: `n $Url"  

                } else {
                    Write-Host "Error: Expected 200, got $([int]$StatusCode)" -ForegroundColor Red        
                }
            }
    }

    
    $Enablement_EndUser_Assignment = ($objroleManagementPolicyRules | ?{$_.id -eq 'Enablement_EndUser_Assignment'})
    
    $EnablementReq_MFA = $Enablement_EndUser_Assignment.enabledRules -contains "MultiFactorAuthentication"
    $EnablementReq_Justification = $Enablement_EndUser_Assignment.enabledRules -contains "Justification"
    $EnablementReq_Ticketing= $Enablement_EndUser_Assignment.enabledRules -contains "Ticketing"

    $Expiration_EndUser_Assignment = ($objroleManagementPolicyRules | ?{$_.id -eq 'Expiration_EndUser_Assignment'})
    $maximumAssignmentDuration = $Expiration_EndUser_Assignment.maximumDuration

    $Approval_EndUser_Assignment = ($objroleManagementPolicyRules | ?{$_.id -eq 'Approval_EndUser_Assignment'}).setting
    $ApprovalRequired = $Approval_EndUser_Assignment.isApprovalRequired



    
    if($Approval_EndUser_Assignment.approvalStages)
    {
        if(@($Approval_EndUser_Assignment.approvalStages.primaryApprovers).count -gt 0)
        {
            $UserApprovers = New-Object System.Collections.ArrayList
            $GroupApprovers = New-Object System.Collections.ArrayList
            foreach($Approver in $Approval_EndUser_Assignment.approvalStages.primaryApprovers )
            {
                if($Approver.userid)
                {
                    $objRoleApprover = New-Object pscustomobject
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "userID" -Value $Approver.userID
                    $ApproverObject = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $($Approver.userID) | select-object -Property displayName,userPrincipalName)
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "displayName" -Value $ApproverObject.displayName
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "description" -Value $Approver.description
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "type" -Value $($Approver.'@odata.type'.split(".")[-1])
                    [VOID]$UserApprovers.Add($objRoleApprover)
                }
                else {
                    $objRoleApprover = New-Object pscustomobject
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "userID" -Value $Approver.groupId
                    $ApproverObject = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $($Approver.groupId) | select-object -Property displayName,userPrincipalName)
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "displayName" -Value $ApproverObject.displayName
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "description" -Value $Approver.description
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "type" -Value $($Approver.'@odata.type'.split(".")[-1])
                    [VOID]$GroupApprovers.Add($objRoleApprover)
                }
            }
            
        }
    }

    
    $PolicyObject = New-Object pscustomobject
    
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "EnablementReq_MFA" -Value $EnablementReq_MFA
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "EnablementReq_Justification" -Value $EnablementReq_Justification
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "EnablementReq_Ticketing" -Value $EnablementReq_Ticketing
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "maximumAssignmentDuration" -Value $maximumAssignmentDuration
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "ApprovalRequired" -Value $ApprovalRequired    
    if($UserApprovers)
    {
        $UserApproversList = [string]$($i = 0; $UserApprovers | %{if($i -eq 0 ){$_.displayname}else{","+ $_.displayName} ; $i++})
    }
    else {
        $UserApproversList = "None"
    }
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "UserApprovers" -Value $UserApproversList
    if($GroupApprovers)
    {
        $GroupApproversList = [string]$($i = 0; $GroupApprovers | %{if($i -eq 0 ){$_.displayname}else{","+ $_.displayName} ; $i++})
    }
    else {
        $GroupApproversList = "None"
    }    
    add-member -InputObject $PolicyObject -MemberType NoteProperty -Name "GroupApprovers" -Value $GroupApproversList
    

    return $PolicyObject

}

Function Get-RoleApprovers
{
    ############################################################################

   <#
   .SYNOPSIS

       Search for role definition id to get the role approvers


   .DESCRIPTION

       Function report settings in the role approvers for a specific role using its role defintion id

   .EXAMPLE

       Get-RoleApprovers -roleDefinitionsid 790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b


   #>

   ############################################################################
    Param
    (        
    # Entra ID Tenant ID
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true, 
                    ValueFromRemainingArguments=$false, 
                    Position=0,
                    ParameterSetName='Default')]
        [ValidatePattern('^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')]                    
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string] 
        $roleDefinitionsid
    )

    $Url = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '/' and scopeType eq 'Directory'"
    $roleManagementPolicyAssignments = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    $objroleManagementPolicyAssignments = $roleManagementPolicyAssignments | Where-Object{$_.roleDefinitionId -eq $roleDefinitionsid}



    $Url = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/$($objroleManagementPolicyAssignments.policyId)/rules"

    $objroleManagementPolicyRules = (Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $Url -Verbose:$false -ErrorAction stop ).Content | ConvertFrom-Json | Select-Object -ExpandProperty Value

    $Approval_EndUser_Assignment = ($objroleManagementPolicyRules | ?{$_.id -eq 'Approval_EndUser_Assignment'}).setting
    $ApprovalRequired = $Approval_EndUser_Assignment.isApprovalRequired
    
    if($ApprovalRequired)
    {
        $Approvers = New-Object System.Collections.ArrayList
        if($Approval_EndUser_Assignment.approvalStages)
        {
            if(@($Approval_EndUser_Assignment.approvalStages.primaryApprovers).count -gt 0)
            {
                foreach($Approver in $Approval_EndUser_Assignment.approvalStages.primaryApprovers )
                {
                    $objRoleApprover = New-Object pscustomobject
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "userID" -Value $Approver.userID
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "description" -Value $Approver.description
                    add-member -InputObject $objRoleApprover -MemberType NoteProperty -Name "type" -Value $($Approver.'@odata.type'.split(".")[-1])
                    [VOID]$Approvers.Add($objRoleApprover)
                }
                
            }
            else {
                Write-Host "Approval required for $RoleName but its missing approvers"
            }
        }
    }
    else {
        $Approvers = ""
    }
    return $Approvers

}


Function Enum-Group
{

    param(
        $Group,
        $RoleMembers,
        $SubscriptionsInfo ,
        $NoPIM,
        $Token,
        $TenantID,
        $Header


    )

    $GroupID = $Group.principalId

    #If PIM is enabled , check if the group is PIM enabled
    if(($SubscriptionsInfo | Where-Object{$_.EntraIDP2 -eq "Success"}) -and (-not($NoPIM)))
    {
        $PIMEnabled = Get-GroupPIMStatus -Token $Token -Header $Header -TenantID $TenantID -ObjectID $GroupID
    }

    if($PIMEnabled)
    {
        # Get Just-in-Time memberships for the group
        $GroupMembers = Get-PIMGroupMembers -Token $Token -Header $Header -TenantID $TenantID -ObjectID $GroupID

    }
    else {

        # Get standard memberships
        $GroupMembers = Get-GroupMembers -Token $Token -Header $Header -TenantID $TenantID -ObjectID $GroupID
    }
    
    if($GroupMembers)
    {
        foreach($GroupMember in $GroupMembers)
        {
            $GroupMemberObject = $(ConvertTo-ObjectArrayListFromPsCustomObject  $Group)
            if($PIMEnabled)
            {
                $GroupMemberid = $GroupMember.principalId
                $GroupMemberObject.Status = $GroupMember.status
                $GroupMemberObject.GroupPIM = "True"
                if($GroupMember.assignmentType)
                {
                    $GroupMemberObject.GroupPIMAssignmentType =  $($TextInfo.ToTitleCase($GroupMember.assignmentType))
                }else{
                    $GroupMemberObject.GroupPIMAssignmentType =  "Eligible"
                }
                
                $GroupMemberObject.startDateTime = $GroupMember.scheduleInfo.startDateTime
                $GroupMemberObject.endDateTime = $GroupMember.scheduleInfo.expiration.endDateTime
                # Since the member is added through a group the membertype should be "Group", like in the portal
                if($GroupMember.accessId -eq "owner")
                {
                    $GroupMemberObject.memberType = "Owner"    
                }
                else
                {
                    $GroupMemberObject.memberType = "Group"
                }
                
                
            
            }
            else {
                $GroupMemberObject.GroupPIM = "False"
                $GroupMemberObject.GroupPIMAssignmentType = ""
                $GroupMemberid = $GroupMember.id
                $GroupMemberObject.startDateTime = $Group.startDateTime
                $GroupMemberObject.endDateTime = $Group.endDateTime
                $GroupMemberObject.Status = $Group.Status
                $GroupMemberObject.memberType = "Group"
            }
            
            $GroupMemberProperties = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $GroupMemberid) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole
            $GroupMemberObject.NestedGroupID = $GroupID
            $GroupMemberObject.NestedGroupdisplayName = $Group.displayName                     

            if($(($GroupMemberObject | get-member -MemberType NoteProperty ).name.contains("userPrincipalName")))
            {  
                                    
                $GroupMemberObject.UserPrincipalName = $GroupMemberProperties.UserPrincipalName                         
            }
            else
            {
                Add-Member -InputObject $GroupMemberObject -MemberType NoteProperty -Name UserPrincipalName $GroupMemberProperties.userPrincipalName
            }

            $GroupMemberObject.principalId = $GroupMemberid
            $GroupMemberObject.displayName = $GroupMemberProperties.displayName
            $GroupMemberObject.Type = $GroupMemberProperties.type     
            
            if($GroupMemberProperties.type -eq "Group")
            {
                if($GroupMemberProperties.isAssignableToRole)
                {
                    $GroupMemberObject.PrivManagement = $True    
                }
                else {
                    $GroupMemberObject.PrivManagement = $false
                }
            }
       


            if($GroupMemberProperties.isAssignableToRole )
            {
            $GroupMemberObject.isAssignableToRole = $GroupMemberProperties.isAssignableToRole                                
            }
            else {
            $GroupMemberObject.isAssignableToRole = $null
            }
            [VOID]$RoleMembers.Add($GroupMemberObject)
            
            #if its a group , check if it got a owner
            if ($GroupMemberObject.Type -eq "Group")
            {
                $Owners = ""
                $Owners = @(Get-GroupOwner -Token $Token -Header $Header -TenantID $TenantID $GroupMemberObject.principalId)
                if($Owners)
                {
                    foreach($Owner in $Owners)
                    {
                        $OwnerObject = New-Object PSCustomObject
                        $OwnerProperties = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $Owner.id) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole

                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name Role -value $GroupMemberObject.Role
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name AdministrativeUnit -value $GroupMemberObject.AdministrativeUnit
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name roleDefinitionId -value $GroupMemberObject.roleDefinitionId
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name assignmentType -value $GroupMemberObject.assignmentType
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name NestedGroupID -value $GroupMemberObject.principalId
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name NestedGroupdisplayName -value $GroupMemberObject.displayName 
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name isPrivileged -value $GroupMemberObject.isPrivileged
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name isApprovalRequired $GroupMemberObject.isApprovalRequired
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name isBuiltIn $GroupMemberObject.isBuiltIn
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name UserPrincipalName $OwnerProperties.userPrincipalName
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name principalId $Owner.id
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name displayName $Owner.displayName
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name startDateTime $GroupMemberObject.startDateTime
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name endDateTime $GroupMemberObject.endDateTime         
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name Type $($TextInfo.ToTitleCase($Owner.'@odata.type'.split(".")[-1]))
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name Status $GroupMemberObject.Status
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name memberType "Owner"    
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name isAssignableToRole $OwnerProperties.isAssignableToRole   
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name GroupPIM $GroupMemberObject.GroupPIM
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name GroupPIMAssignmentType $GroupMemberObject.GroupPIMAssignmentType
                        Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name CollectionDateTime $GroupMemberObject.CollectionDateTime

                        if($OwnerProperties.type -eq "Group")
                        {
                            if($OwnerProperties.isAssignableToRole)
                            {
                                Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name PrivManagement $True    
                            }
                            else {
                                Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name PrivManagement $false
                            }
                        }
                        else {
                            Add-Member -InputObject $OwnerObject -MemberType NoteProperty -Name PrivManagement $GroupMemberObject.PrivManagement
                        }            
            
                        [VOID]$RoleMembers.Add($OwnerObject)

                        if($OwnerProperties.type -eq "Group")
                        {
                            Write-Verbose "Getting group membership of group: $($GroupMemberObject.DisplayName)"   
                            Enum-Group -Group $OwnerProperties -RoleMembers $RoleMembers -SubscriptionsInfo $SubscriptionsInfo -NoPIM $NoPIM -Token $Token -TenantID $TenantID -Header $Header
                        }                        
                    }
                
                }
                Write-Verbose "Getting group membership of group: $($GroupMemberObject.DisplayName)"   
                Enum-Group -Group $GroupMemberObject -RoleMembers $RoleMembers -SubscriptionsInfo $SubscriptionsInfo -NoPIM $NoPIM -Token $Token -TenantID $TenantID -Header $Header
            } #end if a group                       
        }
    }
        
    
}

$scopes = "Agreement.Read.All",`
"AdministrativeUnit.Read.All",`
"Directory.Read.All",`
"email",`
"EntitlementManagement.Read.All",`
"Group.Read.All",`
"IdentityProvider.Read.All",`
"openid",`
"Organization.Read.All",`
"PrivilegedAccess.Read.AzureAD",`          #Not required at the moment
"PrivilegedAccess.Read.AzureADGroup",`      #Required for https://graph.microsoft.com/beta/privilegedAccess/aadGroups/roleAssignments?
"PrivilegedAccess.Read.AzureResources",`   #Not required at the moment
"PrivilegedAssignmentSchedule.Read.AzureADGroup",`
"PrivilegedEligibilitySchedule.Read.AzureADGroup",`
"profile",`
"RoleAssignmentSchedule.Read.Directory",`
"RoleAssignmentSchedule.ReadWrite.Directory",`
"RoleEligibilitySchedule.Read.Directory",`
"RoleManagement.Read.All",`
"RoleManagement.Read.Directory",`
"RoleManagement.ReadWrite.Directory",`
"RoleManagementAlert.Read.Directory",`
"RoleManagementPolicy.Read.Directory",`
"RoleManagementPolicy.Read.AzureADGroup",`
"User.Read",`
"User.Read.All",`
"AgreementAcceptance.Read",`
"AgreementAcceptance.Read.All",`
"AuditLog.Read.All",`
"Policy.Read.All"

# Starts Authentication
if($ClearCache)
{
    Clear-MsalTokenCache
}
if($SkipLogin)
{

    $LoginToken = Get-MsalToken -ClientId "14d82eec-204b-4c2f-b7e8-296a70dab67e" -Scopes $Scopes -TenantId $TenantID -ForceRefresh
}
else
{

    $LoginToken = Get-MsalToken -ClientId "14d82eec-204b-4c2f-b7e8-296a70dab67e" -Scopes $Scopes -TenantId $TenantID

}

if($LoginToken)
{
    $Token =  ($LoginToken).AccessToken
    $Header = Get-AzureADIRHeader -Token $Token
    
    
}
else {
    Write-host "Access Token is missing!" -ForegroundColor Red
    break;
}

if($Token -and $Header)
{
    
    ##Roles
    Write-Verbose "Getting all role definitions"
    $Roles = Get-RoleDefinitionsRoleManagement -Token $Token -Header $Header -TenantID $TenantID

}
else {
    break;
}



$SubscriptionsInfo = Get-EntraIDSKUs $Token $Header

$TextInfo = (Get-Culture).TextInfo


    #$PrivRolesAssignments
    if($Roles)
    {
        

        if(($SubscriptionsInfo | ?{$_.EntraIDP2 -eq "Success"}) -and (-not($NoPIM)))
        {
            Write-Verbose "PIM active"
            Write-Verbose "Getting all role assignments"            
            $PrivRolesAssignments = Get-PIMRoleAssignments -Token $Token -Header $Header -TenantID $TenantID
            
            Write-Verbose "Getting all modified role management policies" 
            $ModfiedroleManagementPolicies = Get-ModifiedRoleManagementPolicies -Header $Header

            $roleManagementPolicyAssignments = Get-RoleManagementPolicyAssignments $Header

            Write-Verbose "Getting all role policy assignments (Rule settings)"
            foreach($Role in $Roles)
            {
                Write-Verbose "Getting role policy assignments (Rule settings): $($Role.DisplayName)"            
                if($roleManagementPolicyAssignments)
                {
                    # Get the policy id for the specific role
                    $objroleManagementPolicy = ($roleManagementPolicyAssignments | Where-Object{$_.roleDefinitionId -eq $($Role.templateId)}).policyId
                    
                    #Return the isApprovalRequired value
                    $PolicyObject = Get-RoleManagementPolicySettings -Token $Token -Header $Header -policyId $objroleManagementPolicy
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name isApprovalRequired -Value $PolicyObject.ApprovalRequired
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name EnablementReq_MFA -Value $PolicyObject.EnablementReq_MFA
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name maximumAssignmentDuration -Value $PolicyObject.maximumAssignmentDuration
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name UserApprovers -Value $PolicyObject.UserApprovers
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name GroupApprovers -Value $PolicyObject.GroupApprovers
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name EnablementReq_Justification -Value $PolicyObject.EnablementReq_Justification
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name EnablementReq_Ticketing -Value $PolicyObject.EnablementReq_Ticketing
                    Add-Member -InputObject $Role -MemberType NoteProperty -Name RulesModified -Value $(if($ModfiedroleManagementPolicies.id -contains $objroleManagementPolicy){$true}else{$false})
                }          
            }
            
        }
        else {
            Write-Verbose "PIM not active"
            Write-Verbose "Getting all role assignments"
            $PrivRolesAssignments = Get-RoleAssignments -Token $Token -Header $Header -TenantID $TenantID
        }

        if($PrivRolesAssignments)
        {
            $RoleMembers = New-Object System.Collections.ArrayList
            $arrRoleDefinitionId = ($PrivRolesAssignments | select-object -Unique RoleDefinitionId).RoleDefinitionId
            Write-Verbose "Collecting membership information"

            foreach($RoleDefinitionId in $arrRoleDefinitionId)
            {
                

                # Array of members for the same role
                $Members = @($PrivRolesAssignments | Where-Object{$_.RoleDefinitionId -eq $RoleDefinitionId})

                if(($SubscriptionsInfo | ?{$_.EntraIDP2 -eq "Success"}) -and (-not($NoPIM)))
                {
                    # For the first member with this role definition id get the role definition object
                    $RoleObject = $Roles | Where-Object{$_.templateId -eq @($Members)[0].roleDefinition.templateId}
                  
                }
                else {
                    # Get the role object from the list of role definitions that match the templateID
                    $RoleObject = ($Roles | Where-Object{$RoleDefinitionId -eq $_.templateId})
                }                
                
                Write-Verbose "Collecting role information for $($RoleObject.displayname)"

                $Members| ForEach-Object{

                    # If PIM is enabled on the tenant , check if the role requires an approval
                    if(($SubscriptionsInfo | Where-Object{$_.EntraIDP2 -eq "Success"}) -and (-not($NoPIM)))
                    {
                        $MemberObject = $_ | select-object -Property id,principalId,roleDefinitionId,startDateTime,endDateTime,memberType,assignmentType,directoryScopeId,appScopeId
                        $MemberObject = $(ConvertTo-ObjectArrayListFromPsCustomObject  $MemberObject)
                        $MemberProperties = $_.principal
                        $RoleScheduleRequests = $null
                        if($_.roleEligibilityScheduleId)
                        {
                            $MemberObject.assignmentType = "Eligible"
                            # Check if the schedule id is a GUID, else it will not have a request to it
                            if( $_.roleEligibilityScheduleId -match '^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')
                            {
                                $RoleScheduleRequests = Get-RoleEligibilityScheduleRequests -Token $Token -Header $Header -TenantID $TenantID -ObjectID $_.roleEligibilityScheduleId
                            }
                        }
                        else {
                            # Check if the schedule id is a GUID, else it will not have a request to it
                            if( $_.roleAssignmentScheduleId -match '^[A-Za-z0-9]{4}([A-Za-z0-9]{4}\-?){4}[A-Za-z0-9]{12}$')
                            {                            
                                $RoleScheduleRequests = Get-RoleAssignmentScheduleRequests -Token $Token -Header $Header -TenantID $TenantID -ObjectID $_.roleAssignmentScheduleId
                            }
                        }

                        if($RoleScheduleRequests)
                        {

                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name justification $RoleScheduleRequests.justification
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name status $RoleScheduleRequests.Status
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name RequestcreatedDateTime $RoleScheduleRequests.createdDateTime

                            if($RoleScheduleRequests.createdBy.user.id)
                            {
                                $RequestCreator = ""
                                $RequestCreator = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $($RoleScheduleRequests.createdBy.user.id) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole)
                                Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name Creator $RequestCreator.userPrincipalName
                                Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name CreatorId $RoleScheduleRequests.createdBy.user.id
                            }
                        }
                        else {
                            
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name justification $null
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name status $null
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name Creator $null
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name CreatorId $null                            
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name RequestcreatedDateTime $null
                        }

                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name type $($TextInfo.ToTitleCase($MemberProperties.'@odata.type'.split(".")[-1]))

                    }
                    else {
                        #If PIM is not enabled the member object needs to be created with the same type of attributes

                        #$isApprovalRequired only works if PIM is activated
                        $isApprovalRequired = ""

                        $principalId = $_.principalId
                        $MemberObject = New-Object PSCustomObject
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name id -value $_.id
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name principalId -value $_.principalId
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name linkedEligibleRoleAssignmentId -value $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name directoryScopeId $_.directoryScopeId
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name resourceScope $_.resourceScope
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name appScopeId $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name startDateTime $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name endDateTime "Permanent"
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name memberType "Direct"
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name assignmentType "Assigned"
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name status $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name justification $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name RequestcreatedDateTime $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name Creator $null
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name CreatorId $null
                        
                        # Search for the principal to get additional data
                        $MemberProperties = (FindGlobalObject -Token $Token -Header $Header -TenantID $TenantID -ObjectID $principalId) | select-object -Property displayName,type,userPrincipalName,isAssignableToRole
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name type $MemberProperties.type
                    }                
                   
                    if($MemberObject.directoryScopeId -eq "/")
                    {
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name AdministrativeUnit "None"
                    }
                    else
                    {
                        $objAdministrativeUnit = Get-AdministrativeUnit -Token $Token -Header $Header -TenantID $TenantID -ObjectID $($MemberObject.directoryScopeId.split("/")[-1])
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name AdministrativeUnit $objAdministrativeUnit.displayname
                    }
                    

                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name Role $RoleObject.displayName
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name displayName $MemberProperties.displayName
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name userPrincipalName $MemberProperties.userPrincipalName
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name isAssignableToRole $MemberProperties.isAssignableToRole
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name isPrivileged $RoleObject.isPrivileged
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name isBuiltIn $RoleObject.isBuiltIn
                    if($MemberProperties.type -eq "Group")
                    {
                        if($MemberProperties.isAssignableToRole)
                        {
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name PrivManagement $True    
                        }
                        else {
                            Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name PrivManagement $false
                        }
                    }
                    else {
                        Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name PrivManagement $True
                    }
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name isApprovalRequired $RoleObject.isApprovalRequired
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name NestedGroupID -value ""
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name NestedGroupdisplayName -value ""
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name GroupPIM -value "False"
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name GroupPIMAssignmentType -value ""
                    Add-Member -InputObject $MemberObject -MemberType NoteProperty -Name CollectionDateTime $ReportDateTimeFileUTC
                    [VOID]$RoleMembers.Add($MemberObject)
                } 
                
            }

            # Do additional checks if the member is a group
            if ($RoleMembers | Where-Object{$_.Type -eq "Group"})
            {
                $MemberTypeGroup = $RoleMembers | Where-Object{$_.Type -eq "Group"}
                foreach ($Group in $MemberTypeGroup)
                {
                    Write-Verbose "Getting group membership of group: $($Group.DisplayName)"   
                    Enum-Group -Group $Group -RoleMembers $RoleMembers -SubscriptionsInfo $SubscriptionsInfo -NoPIM $NoPIM -Token $Token -TenantID $TenantID -Header $Header

                }
            }


                     

            #$RoleMembers | select-object -Property roleDefinitionId, Role, id,displayName,userPrincipalName,principalId,type, CollectionDateTime, startDateTime,endDateTime,appScopeId,directoryScopeId,AdministrativeUnit, status, assignmentType, memberType,justification, isAssignableToRole, isPrivileged, isBuiltIn, isApprovalRequired, NestedGroupID, NestedGroupdisplayName, GroupPIM, GroupPIMAssignmentType, Creator, CreatorId | export-csv -Path $CSVFile -NoClobber -NoTypeInformation 

            $strHTMLTextCurrent = $strHTMLTextCurrent + '</script>'
            $strFontColor = "#F4A100"
            $OrganizationData = Get-TenantInformation $Token $Header 
            $DefaultDomainName = (($OrganizationData).verifiedDomains | Where-Object{$_.isDefault}).Name
            $strHTMLTextCurrent = $strHTMLTextCurrent + "<h1><font color='$strFontColor'>$ToolName - $($DefaultDomainName)</font></h1>"
            $strFontColor = "#ffffff"
            $strHTMLTextCurrent = $strHTMLTextCurrent + '<div class="tab_sum">'
$SummaryTabs = @"
<button class="tablinks_sum" onclick="openSummary(event, 'Summary')">Summary</button>
<button class="tablinks_sum" onclick="openSummary(event, 'All Roles')">All Roles</button>
<button class="tablinks_sum" onclick="openSummary(event, 'Assigned Accounts')">Assigned Accounts</button>
"@            
            $strHTMLTextCurrent = $strHTMLTextCurrent + $SummaryTabs
            $strHTMLTextCurrent = $strHTMLTextCurrent + '</div>'
            $strHTMLTextCurrent = $strHTMLTextCurrent + '<div id="Summary" class="tabcontent_sum" style="display: block;">'
            $strHTMLTextCurrent = $strHTMLTextCurrent + '<table id="TopTable"><tr id="TopTable"><td id="TopTable">'
            $strHTMLTextCurrent = $strHTMLTextCurrent + '<h3>Tenant Information</h3>'
            #Table with Tenant Information
            $TenantData = New-Object PSCustomObject
            Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Tenant ID" -Value $TenantId
            
            Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Creation Date" -Value $(($OrganizationData).createdDateTime)
            $IntialDomainName = (($OrganizationData).verifiedDomains | Where-Object{$_.isInitial}).Name
            Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Initial Domain" -Value $IntialDomainName
                 Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Default Domain" -Value $DefaultDomainName
            $P1License = ""
            $P1License = ($SubscriptionsInfo | ?{$_.EntraIDP1 -eq "Success"}).SubscriptionName
            
            if($P1License)
            {
                If ($dicEntraLicensing.ContainsKey($P1License))
                {
                    $P1License = $dicEntraLicensing.Item($P1License)
                }                
                
                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "P1 License" -Value $P1License
            }
            else {
                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "P1 License" -Value "n/a"
            }
            $P2License = ""
            $P2License = ($SubscriptionsInfo | ?{$_.EntraIDP2 -eq "Success"}).SubscriptionName
            
            if($P2License)
            {
                If ($dicEntraLicensing.ContainsKey($P2License))
                {
                    $P2License = $dicEntraLicensing.Item($P2License)
                }

                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "P2 License" -Value $P2License
            }
            else {
                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "P2 License" -Value "n/a"
            }        
            if($SubscriptionsInfo | ?{$_.EntraIDP2 -eq "Success"})
            {
                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "PIM" -Value "Enabled"
            }
            else {
                Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "PIM" -Value "Disabled"
            }
            Add-Member -InputObject $TenantData -MemberType NoteProperty -Name "Total Licenses" -Value $(($SubscriptionsInfo | ?{$_.EntraIDP2 -eq "Success"}).TotalLicenses)            
            $TenantTable = ($TenantData | ConvertTo-Html -Fragment -As List)
            $strHTMLTextCurrent = $strHTMLTextCurrent + $TenantTable + "`n"
            ### Summaries all privileged users in the Tenant
            $strHTMLTextCurrent = $strHTMLTextCurrent + '<p>'
            # create array with assignments
            $AssignmentData = New-Object PSCustomObject
            Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name 'Highly Privileged' -Value $(@(($RoleMembers | Where-object{($_.isPrivileged -eq "True")} | Select-Object -Unique principalId)).count)

            # Add graph for assignments eligible/active
            $strHTMLTextWithGraph = Add-DoughnutGraph -Data $AssignmentData -GraphTitle "Highly Privileged Accounts" -DoughnutTitle "" -arrColors $arrPrivGraphColors -BackColor $ThemeBackGrounColor
            $strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLTextWithGraph
            $strHTMLTextCurrent = $strHTMLTextCurrent + '</p>'
            $strHTMLTextCurrent = $strHTMLTextCurrent + '</td><td id="TopTable">' + "`n"
            ### Graph of all role memberships
            $RoleSummaryData = New-Object PSCustomObject
            $RoleNames = $(($RoleMembers | Select-Object -Property Role -Unique).Role | Sort-Object -Property Role)
            Foreach ($RoleName in $RoleNames)
            {
                Add-Member -InputObject $RoleSummaryData -MemberType NoteProperty -Name $RoleName -Value $((@($RoleMembers | Where-object{$_.Role -eq $RoleName})).count)
            }
            $strHTMLTextWithGraph = Add-BigDoughnutGraph -Data $RoleSummaryData -GraphTitle "Role Population" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
            $strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLTextWithGraph + "`n"
            $strHTMLTextCurrent = $strHTMLTextCurrent + '</td><td id="TopTable">' + "`n"
                 ### Table of all role memberships
            $TableAllRole = ($RoleMembers |  Group-Object -Property role | Select-Object -Property @{N="Members";E={$_.count}},@{N="Role";E={$_.name}} | ConvertTo-Html -Fragment)
            $TableAllRole = $TableAllRole -replace $TableAllRole[1], ""
            $strHTMLTextCurrent = $strHTMLTextCurrent + $TableAllRole + "`n"
            $strHTMLTextCurrent = $strHTMLTextCurrent + "</td></tr></table>" + "`n"
            $strHTMLTextCurrent = $strHTMLTextCurrent + "<p>Click on the buttons inside the tabbed menu:</p>" + "`n"

            $strTab = "<div class='tab'>" + "`n"
            Foreach ($RoleName in $(($RoleMembers | Sort-Object Role | Select-Object -Property Role -Unique).Role))
            {
$RoleTab = @"
<button class="tablinks" onclick="openRole(event, '$RoleName')">$RoleName </button>
"@
$RoleTabPriv = @"
<button class="tablinks" onclick="openRole(event, '$RoleName')">$RoleName <span class="priv-box" data-bind="visible: settings.item.isPrivileged">PRIVILEGED</span></button>
"@
                $IsPriv = $(($RoleMembers | where-object{$_.Role -eq $RoleName} | Select-Object -Property isPrivileged).isPrivileged)
                if($IsPriv -eq "True")
                {
                    $strTab = $strTab + $RoleTabPriv + "`n"
                }
                else {
                    $strTab = $strTab + $RoleTab + "`n"
                }
                
            }
            $strTab = $strTab + "</div>" + "`n"
    
            $strHTMLTextCurrent = $strHTMLTextCurrent + $strTab
    
            $strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLContent
    
            $iCount = 0
            Foreach ($RoleName in $RoleNames)
            {
                #Add Role Name in Header
                $strHTMLGraphs = "<h1><font color='$strFontColor'>$RoleName</font></h1>`n"

                $strHTMLGraphs = $strHTMLGraphs + '<table id="RoleInfoTbl" ><tr id="RoleInfoTbl"><td id="RoleInfoTbl">'

                #Table with Role Information
                $RoleDetailTable = New-Object PSCustomObject
                
                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Activation maximum duration (hours)" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).maximumAssignmentDuration)

                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "MFA On activation" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).EnablementReq_MFA)
                
                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Justification on Activation" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).EnablementReq_Justification)
                
                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Require approval" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).isApprovalRequired)

                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "User Approvers" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).UserApprovers)

                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Group Approvers" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).GroupApprovers)

                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Ticket on Activation" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).EnablementReq_Ticketing)

                Add-Member -InputObject $RoleDetailTable -MemberType NoteProperty -Name "Settings Modified" -Value $(($Roles | Where-object{$_.DisplayName -eq $RoleName}).RulesModified)
    
                $HTMLRoleDetailTable = ($RoleDetailTable | ConvertTo-Html -Fragment -As List)
                                
                $strHTMLGraphs = $strHTMLGraphs + $HTMLRoleDetailTable 

                $strHTMLGraphs = $strHTMLGraphs + '</td></tr></table>'
            

                # Add array with for membertype
                $MemberTypeData = New-Object PSCustomObject
                $MemberTypes = (($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property membertype)
                $MemberTypesNames = ($MemberTypes | Select-Object -property membertype -Unique).membertype
                foreach($MemberType in $MemberTypesNames)
                {
                    Add-Member -InputObject $MemberTypeData -MemberType NoteProperty -Name $MemberType -Value $(@(($MemberTypes | Where-object{$_.membertype -eq $MemberType})).count)
                }
    
                # Add graph for memberships
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $MemberTypeData -GraphTitle "Memberships" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph

    
                # create array with assignments
                $AssignmentData = New-Object PSCustomObject
                Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name Eligible -Value $(@(($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.assignmentType -eq "Eligible")})).count)
                Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name Active -Value $(@(($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.assignmentType -eq "Assigned")})).count)
    
                # Add graph for assignments eligible/active
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $AssignmentData -GraphTitle "Assignments" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph
    
   
                # Create array of types
                $TypeData = New-Object PSCustomObject
                $Types = (($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property type)
                $TypesNames = ($Types | Select-Object -property type -Unique).type
                foreach($Type in $TypesNames)
                {
                    Add-Member -InputObject $TypeData -MemberType NoteProperty -Name $Type -Value $(@(($Types | Where-object{$_.type -eq $Type})).count)
                }
    
                # Add graph for object types
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $TypeData -GraphTitle "Member object types" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph + "`n"
                
                # Create array of unique assignments
                $UniqueMemberData = New-Object PSCustomObject
                Add-Member -InputObject $UniqueMemberData -MemberType NoteProperty -Name "Unique" -Value $(@((($RoleMembers | Where-object{($_.Role -eq $RoleName)}) | Select-Object -property principalId -Unique)).count)
                # Add graph for unique assignments
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $UniqueMemberData -GraphTitle "Unique assignments" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph + "`n"
                
                # Create array of nested members
                $NestedMemberData = New-Object PSCustomObject
                Add-Member -InputObject $NestedMemberData -MemberType NoteProperty -Name "Direct" -Value $(@((($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.NestedGroupdisplayName -eq "")}) | Select-Object -property principalId -Unique)).count)
                Add-Member -InputObject $NestedMemberData -MemberType NoteProperty -Name "Nested" -Value $(@((($RoleMembers | Where-object{($_.Role -eq $RoleName) -and ($_.NestedGroupdisplayName -ne "")}) | Select-Object -property principalId -Unique)).count)
                # Add graph for unique assignments
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $NestedMemberData -GraphTitle "Nested members" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLGraphs = $strHTMLGraphs + $strHTMLTextWithGraph + "`n"
    
                $strHTMLRoleTable = ""
                $strHTMLGraphs = $strHTMLGraphs + '<table id="RoleMemberTblStyle" ><tr id="RoleMemberTblStyle"><td id="RoleMemberTblStyle">'
                
                $strHTMLRoleTable = ($RoleMembers | Where-object{$_.Role -eq $RoleName} | Select-Object @{Name = "Display Name"; E = {if($_.PrivManagement -eq $false){$($_.DisplayName + "[WARN]")}else{$_.DisplayName}}},@{Name = "UserPrincipalName"; E = {$_.userPrincipalName}},@{Name = "Subject ID"; E = {$_.principalId}},@{Name = "Admin Unit"; E = {$_.AdministrativeUnit}},@{Name = "Type"; E = {$_.type}},@{Name = "Priv Mgmt"; E = {if($_.PrivManagement -eq $false){$("[RED]" + $_.PrivManagement.toString())}else{$_.PrivManagement.toString()}}},@{Name = "Assignment State"; E = {$_.assignmentType}},@{Name = "Member Type"; E = {$_.memberType}},@{Name = "Status"; E = {$_.Status}},@{Name = "Nested Group"; E = {$_.NestedGroupdisplayName}},@{Name = "Start Time"; E = {$_.startDateTime}},@{Name = "End Time"; E = {if($_.endDateTime){$_.endDateTime}else{"Permanent"}}} | ConvertTo-Html -Fragment).replace("<table>",'<table id="myTable' + $iCount + '">')
                $strHTMLRoleTable = $strHTMLRoleTable.replace("<tr><td>",'<tr class="item"><td>')
                $strHTMLRoleTable = $strHTMLRoleTable.replace("<td>[RED]",'<td class="WARN">')
                $strHTMLRoleTable = $strHTMLRoleTable.replace("[WARN]",'<span class="warn-box" data-bind="visible: settings.item.isPrivileged">WARNING</span>')
                $tableHeaders = ($strHTMLRoleTable | select-string -Pattern "<th>").tostring().split("/")
                $NewtableHeaders = "" 
                $i = 1
                ForEach($tbHead in $tableHeaders)
                {
$NewHeader = @"
<th onclick="w3.sortHTML('#myTable$iCount', '.item', 'td:nth-child($i)')" style="cursor:pointer">
"@        
                    if($tbHead.Length -gt 7)
                    {
                    $NewtableHeaders = $NewtableHeaders + $tbHead.replace("<th>",$NewHeader) +"/"
                    }
                    $i++
                }
                $NewtableHeaders = $NewtableHeaders + "</tr>"
                $strHTMLRoleTable = $strHTMLRoleTable -replace  "^<tr><th.+", $NewtableHeaders
                $strHTMLRoleTable = $strHTMLRoleTable + '</td></tr></table>'
                $strHTMLRoleTable = $strHTMLRoleTable + "`n"
$strHTMLRoleContent = @"
<div id="$RoleName" class="tabcontent">
    $strHTMLGraphs
    $strHTMLRoleTable 
</div>
"@    
                $strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLRoleContent + "`n"
                $iCount++
        
                }
                $strHTMLTextCurrent = $strHTMLTextCurrent + '</div>'

####
                $strHTMLAllRole = ""
                $strHTMLAllRole = $strHTMLAllRole + "<h1><font color='$strFontColor'>All Roles</font></h1>`n"   
                
                # Create array of Roles Require Approval
                $RolesRequireApproval = New-Object PSCustomObject
                Add-Member -InputObject $RolesRequireApproval -MemberType NoteProperty -Name "Privileged Role" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $True) -and ($_.isApprovalRequired -eq $False)}) | Select-Object -property templateId -Unique)).count)
                Add-Member -InputObject $RolesRequireApproval -MemberType NoteProperty -Name "Standard Role" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $False) -and ($_.isApprovalRequired -eq $False)}) | Select-Object -property templateId -Unique)).count)
                # Add graph for unique assignments
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $RolesRequireApproval -GraphTitle "Roles Without Approvals" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLAllRole = $strHTMLAllRole + $strHTMLTextWithGraph + "`n"    
                
                # Create array of Roles Without MFA on Activation
                $RoleNoMFAActication = New-Object PSCustomObject
                Add-Member -InputObject $RoleNoMFAActication -MemberType NoteProperty -Name "Privileged Role" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $True) -and ($_.EnablementReq_MFA -eq $False)}) | Select-Object -property templateId -Unique)).count)
                Add-Member -InputObject $RoleNoMFAActication -MemberType NoteProperty -Name "Standard Role" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $False) -and ($_.EnablementReq_MFA -eq $False)}) | Select-Object -property templateId -Unique)).count)
                # Add graph for unique assignments
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $RoleNoMFAActication -GraphTitle "Roles Without MFA on Activation" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLAllRole = $strHTMLAllRole + $strHTMLTextWithGraph + "`n"          
                
                # Create array of Roles that is custom
                $RoleCustom = New-Object PSCustomObject
                Add-Member -InputObject $RoleCustom -MemberType NoteProperty -Name "Custom Privileged Roles" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $True) -and ($_.isBuiltIn -eq $False)}) | Select-Object -property templateId -Unique)).count)
                Add-Member -InputObject $RoleCustom -MemberType NoteProperty -Name "Custom Roles" -Value $(@((($Roles | Where-object{($_.isPrivileged -eq $False) -and ($_.isBuiltIn -eq $False)}) | Select-Object -property templateId -Unique)).count)
                # Add graph for unique assignments
                $strHTMLTextWithGraph = Add-DoughnutGraph -Data $RoleCustom -GraphTitle "Custom Roles" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
                $strHTMLAllRole = $strHTMLAllRole + $strHTMLTextWithGraph + "`n"                        

                $strHTMLAllRoleTable = ($Roles | Select-Object @{Name = "Role Name"; E = {$_.displayname}},@{Name = "Template ID"; E = {$_.templateId}},@{Name = "Members"; E = {$Role = $_.displayname; ($RoleMembers | Where-object{$_.Role -eq $Role}).count}},@{Name = "Built-in"; E = {$_.isBuiltIn}},@{Name = "Settings Modified"; E = {$_.RulesModified}},@{Name = "Privileged"; E = {$_.isPrivileged}},@{Name = "Activation Hours"; E = {$_.maximumAssignmentDuration}},@{Name = "Require Approvals"; E = {$_.isApprovalRequired}},@{Name = "User Approver"; E = {$_.UserApprovers}},@{Name = "Group Approver"; E = {$_.GroupApprovers}} ,@{Name = "MFA on Activation"; E = {$_.EnablementReq_MFA}} ,@{Name = "Justification on Activation"; E = {$_.EnablementReq_Justification}} ,@{Name = "Ticket on Activation"; E = {$_.EnablementReq_Ticketing}} | ConvertTo-Html -Fragment).replace("<table>",'<table id="allRoleTable' + $iCount + '">')
                $strHTMLAllRoleTable = $strHTMLAllRoleTable.replace("<tr><td>",'<tr class="item"><td>')
                $tableHeaders = ($strHTMLAllRoleTable | select-string -Pattern "<th>").tostring().split("/")
                $NewtableHeaders = "" 
                $i = 1
                ForEach($tbHead in $tableHeaders)
                {
$NewHeader = @"
<th onclick="w3.sortHTML('#allRoleTable$iCount', '.item', 'td:nth-child($i)')" style="cursor:pointer">
"@        
    if($tbHead.Length -gt 7)
    {
    $NewtableHeaders = $NewtableHeaders + $tbHead.replace("<th>",$NewHeader) +"/"
    }
    $i++
}
                $NewtableHeaders = $NewtableHeaders + "</tr>"
                $strHTMLAllRoleTable = $strHTMLAllRoleTable -replace  "^<tr><th.+", $NewtableHeaders
                $strHTMLAllRole = $strHTMLAllRole + $strHTMLAllRoleTable + "`n"
$strHTMLRoleContent = @"
<div id='All Roles' class="tabcontent_sum">
$strHTMLAllRole 
</div>
"@    
                $strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLRoleContent + "`n"
####
$strHTMLPrivAcc = ""
$strHTMLPrivAcc = $strHTMLPrivAcc + "<h1><font color='$strFontColor'>Highly Privileged Accounts</font></h1>`n"   

### Summaries all privileged users in the Tenant
$strHTMLPrivAcc = $strHTMLPrivAcc + '<p>'
# create array with assignments
$AssignmentData = New-Object PSCustomObject
Add-Member -InputObject $AssignmentData -MemberType NoteProperty -Name 'Highly Privileged' -Value $(@(($RoleMembers | Where-object{($_.isPrivileged -eq "True")} | Select-Object -Unique principalId)).count)

# Add graph for all privileged users in the Tenant
$strHTMLTextWithGraph = Add-DoughnutGraph -Data $AssignmentData -GraphTitle "Highly Privileged Accounts" -DoughnutTitle "" -arrColors $arrPrivGraphColors -BackColor $ThemeBackGrounColor
$strHTMLPrivAcc = $strHTMLPrivAcc + $strHTMLTextWithGraph

### Graph of all role memberships
$PrivRoleSum = New-Object PSCustomObject
$RoleNames = $(($RoleMembers | Where-object{($_.isPrivileged -eq "True")} | Select-Object -Property Role -Unique).Role | Sort-Object -Property Role)
Foreach ($RoleName in $RoleNames)
{
    Add-Member -InputObject $PrivRoleSum -MemberType NoteProperty -Name $RoleName -Value $((@($RoleMembers | Where-object{($_.isPrivileged -eq "True")} | Where-object{$_.Role -eq $RoleName})).count)
}
$strHTMLTextWithGraph = Add-DoughnutGraph -Data $PrivRoleSum -GraphTitle "Role Population" -DoughnutTitle "" -arrColors $arrColors -BackColor $ThemeBackGrounColor
$strHTMLPrivAcc = $strHTMLPrivAcc + $strHTMLTextWithGraph 


$strHTMLPrivAcc = $strHTMLPrivAcc + '</p>'

#$PrivAccounts = ($RoleMembers | where-object{$_.isPrivileged -eq "True"} |Sort-Object -Property PrincipalId -Unique) 
$PrivAccounts = ($RoleMembers | Sort-Object -Property PrincipalId -Unique) 
$strHTMLPrivAccTable = ($PrivAccounts | Select-Object @{Name = "Display Name"; E = {$_.displayName}},@{Name = "Principal Id"; E = {$_.principalId}},@{Name = "Role"; E = {$_.Role}},@{Name = "Object Type"; E = {$_.type}},@{Name = "Privileged"; E = {$_.isPrivileged}},@{Name = "Member Type"; E = {$_.memberType}},@{Name = "Administrative Unit"; E = {$_.AdministrativeUnit}} ,@{Name = "Assignment Type"; E = {$_.assignmentType}} ,@{Name = "Approval Required"; E = {$_.isApprovalRequired}} ,@{Name = "Justification"; E = {$_.Justification}},@{Name = "Created"; E = {$_.RequestcreatedDateTime}},@{Name = "Assigned by"; E = {$_.Creator}} | ConvertTo-Html -Fragment).replace("<table>",'<table id="allRoleTable' + $iCount + '">')
$strHTMLPrivAccTable = $strHTMLPrivAccTable.replace("<tr><td>",'<tr class="item"><td>')
$tableHeaders = ($strHTMLPrivAccTable | select-string -Pattern "<th>").tostring().split("/")
$NewtableHeaders = "" 
$i = 1
ForEach($tbHead in $tableHeaders)
{
$NewHeader = @"
<th onclick="w3.sortHTML('#PrivAccTable$iCount', '.item', 'td:nth-child($i)')" style="cursor:pointer">
"@        
if($tbHead.Length -gt 7)
{
$NewtableHeaders = $NewtableHeaders + $tbHead.replace("<th>",$NewHeader) +"/"
}
$i++
}
$NewtableHeaders = $NewtableHeaders + "</tr>"
$strHTMLPrivAccTable = $strHTMLPrivAccTable -replace  "^<tr><th.+", $NewtableHeaders
$strHTMLPrivAcc = $strHTMLPrivAcc + $strHTMLPrivAccTable + "`n"
$strHTMLRoleContent = @"
<div id='Assigned Accounts' class="tabcontent_sum">
$strHTMLPrivAcc 
</div>
"@    
$strHTMLTextCurrent = $strHTMLTextCurrent + $strHTMLRoleContent + "`n"
####               

            $strHTMLTextCurrent = $strHTMLTextCurrent + '</div>'                
            $strHTMLTextCurrent | Out-File -FilePath $("$HTMLReport") -Force
            Write-Host "Report written to:" -ForegroundColor Yellow
            Write-Output $HTMLReport
            if($Show)
            {
                Invoke-Item -Path $HTMLReport
            }

    }
    else {
        
        Write-host "Could not get role assignemnts from EntraID tenant $Tenantid" -ForegroundColor Red 
        Write-host "Verify the access permissions for $((Parse-JWToken $Token).unique_name)" -ForegroundColor Red 

    
    }
}
else {
    
    Write-host "Could not get role definitions from EntraID tenant $Tenantid" -ForegroundColor Red 
    Write-host "Verify the access permissions for $((Parse-JWToken $Token).unique_name)" -ForegroundColor Red 

}