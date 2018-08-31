# Unzip All Files From Release Directory
#
# VERSION 1.0.20180516
# Copyright (c) 2008–2018 G Treasury SS, LLC
#
# [01] 07/09/2018 JML Locked file check loop
# [00] 05/15/2018 JML Original

Param (
    [string]$fromDir = "C:\GTreasury\Admin\PatchDeploy\" + "TEST"
)

# #############################################################################################################
# ### Identifies each zip file in the release to be unpacked
# ### Unzips into directory that strips out first four numeric segements of the name.  
# ###   e.g. 2016_9_0_7154_Delphi_Patch into Delph_Patch
# ### If the destination folder already exists, then it assumes the files have already been unzipped. 
# #############################################################################################################

#static vars
[bool]$didLock = $false
[int]$turn = 1

if ($fromDir) {
    if (!(get-childitem $fromDir | where {$_.Name -like '*.zip'})) {
        echo "No ZIP files found.  Terminating Patch Process"

    }
    foreach ($fl in (get-childitem $fromDir | where {$_.Name -like '*.zip'})) {
        if (Test-path $fromDir\$fl) {
            try {
                $lockedFile = [System.io.File]::Open("$fromDir\$fl", 'append', 'Write', 'None')
                $didLock = $false
                write-output "$fl is not locked by another process."
                $lockedFile.close()
            }
            catch {
                $didLock = $true
            }
            if ($didLock) {
                Write-Output "$fl is locked by another process."             
                do {
                    $turn++
                    try {
                        Write-Output "Waiting for file to unlock attempt# $turn"
                        $lockedFile = [System.io.File]::Open("$fromDir\$fl", 'append', 'Write', 'None')
                        $didLock = $false
                        write-output "$fl is not locked by another process."
                        $lockedFile.close()
                    }
                    catch {
                        $turn++
                        Write-Output "$fl is locked by another process."
                        Start-Sleep -s 1
                    }

                }while (($didLock) -and ($turn -lt 60))
            }
            if ($didLock) {
                throw "Decompression failed! $inFile appears to be encrypted, not compressed or empty."
            }
        }

        if ($fl.PSIsContainer -eq $false) {
            $dirName = ([string]$fl.basename.split("\")[-1].split("_")[4..$fl.basename.split("\")[-1].split("_").length]).replace(" ", "_")
            $dirTest = $fromDir + '\' + $dirName
            if (Test-Path $dirTest) {
                echo 'Packaged already unzipped ' + $dirTest
            }
            else {
                $newDir = $fromDir + '\' + $dirName
                $inFile = $fromDir + '\' + $fl
                New-Item $newDir -type directory
                $unzipper = New-Object -COMObject "shell.application"  
                $zippedFile = $unzipper.NameSpace($inFile) 2>$NULL

                if ($zippedFile.items().count -gt 0) {
                    [int]$fileCount = 0 
                    foreach ($i in $zippedFile.items()) {
                        $thisExtractFileName = $i.name
                        $thisExtract = $newDir + "\" + $thisExtractFileName
                        $cTimer = 0
                        do {
                            $unzipper.NameSpace($newDir).copyhere($i)
                            $cTimer++
                        }
                        while ((!(test-path $thisExtract)) -and ($cTimer -lt 1))

                        if (test-path $thisExtract) {
                            echo "$thisExtractFileName extracted!"
                            $fileCount ++
                        }
                    }
                    if ($fileCount -gt 0) {
                        echo "$inFile appeared to have decompressed successfully into $fileCount files."
                        echo "---------------------------------------------------------------------------------------------------------------------"
                        echo "---------------------------------------------------------------------------------------------------------------------"
                    }
                    else {
                        echo "Unable to unzip $inFile!"
                        echo "---------------------------------------------------------------------------------------------------------------------"
                        echo "---------------------------------------------------------------------------------------------------------------------"
                    }
                }
                else {
                    echo "Decompression failed! $inFile appears to be encrypted, not compressed or empty."
                    echo "---------------------------------------------------------------------------------------------------------------------"
                    echo "---------------------------------------------------------------------------------------------------------------------"

                }
            }
        }
    }
}
else {
    '-fromDir required'
}