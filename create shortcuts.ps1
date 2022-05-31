#########TODO########
#--------------------
#allow -notmatch in file specification
#fix =============== color diplaying error
#add foeder creation for sending items
#add support for sending items to multiple forlders


############################ Functions ##################################

function New-EnvironmentalVariable
{
    [string] $varName = Read-Host "`nEnter the name of the variable to create";
    [string] $varValue =  Read-Host "`nEnter the value of the variable to create";

    #If the console is without Admin priviledges then the host will be assign user without request.
    if(-not $Global:isAdmin){
        $requestedScope = "user";
    }
    #If console has Admin priviledges then the host has the option to decide over user o machine scope.
    elseif($Global:isAdmin){
        Write-Output "";
        $message = "Choose the scope for the environmental variable";
        $user = New-Object System.Management.Automation.Host.ChoiceDescription "&USER", "To use USER scope.";
        $machine = New-Object System.Management.Automation.Host.ChoiceDescription "&MACHINE", "To use MACHINE scope.";
    
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($user, $machine);
        #This creates a better prompt for choosing.
        $scopeOption = $host.ui.PromptForChoice("Scope",$message, $options, -1);

        switch($scopeOption)
        {
            0 {$requestedScope = "user";}
            1 {$requestedScope = "machine";}
        }
    }

    #Environmental funciton requires a name, a value, and a scope (determine or not by the host).
    [System.Environment]::SetEnvironmentVariable($varName, $varValue, $requestedScope)

    Write-Host "`n`nEnvironmental Variable";
    Write-Host "`r----------------------";
    #Prints the env variable to the console.
    [System.Environment]::GetEnvironmentVariables($requestedScope).GetEnumerator() | Where-Object Name -eq $varName;
}

function New-Shortcut([array] $AvailableFiles)
{
    $Global:reqPathToSendFiles = Read-HostPath "`nEnter the complete path in which you want to send the item(s) `n";
    foreach($file in $AvailableFiles)
    {
        #Replace changes the file extension to the one that shortcuts have, which is 'lnk'.
        $fileRenameExtension = [System.IO.Path]::ChangeExtension($file,"lnk");
        $createShortcut = [System.IO.Path]::Combine($Global:reqPathToSendFiles,$fileRenameExtension);
        $arguments = $null;
        $targetPath = [System.IO.Path]::Combine($Global:reqPathToWork,$file);
        $workingDirectory = $Global:reqPathToSendFiles;

        $extension = [System.IO.Path]::GetExtension($file);
        #Sees if any file is a Powershell script to ask the host what to do.
        if($extension -eq '.ps1') {
            do{
               $specialShortcut = Read-Host "`n""$file"" has been detected as a Powershell Script.
               `rDo you want to create a special Powershell Shortcut to run your script 
               `rinstead of a normal file shortcut? [Y/N]";
            }while($specialShortcut -notmatch "[yYnN]")
        }

        #If the host accepts to make a Powershell shortcut, it enters here.
        if($specialShortcut -match "[yY]")
        {
            #Powershell program location, PSHOME contains the installation path.
            if ($PSVersionTable.PSVersion.Major -le 5) {
                $targetPath = "$PSHOME\powershell.exe";
            }
            else {
                #Powershell name changes to pwsh in version 6 and higher.
                $targetPath = "$PSHOME\pwsh.exe";
            }
            #If you use Windows Terminal and have it configure as the default terminal,
            #It doesn't matter if it is a Powershell shortcut, it will still open Windows Terminal.
                            
            #Directory where powershell program will start on in command line.
            #The same as the location of the file/script.
            $workingDirectory = $Global:reqPathToWork;
            
            #Arguments after the targetPath. -noexit to keep the console running, and -command to run script.
            $scriptPath = [System.IO.Path]::Combine($Global:reqPathToWork,$file)
            $arguments = "-NoExit -Command ""& { . '$scriptPath'}""";
        }
    
        #This WScript.Shell object is the one that has the properties to change settings in shortcuts.
        $WshShell = New-Object -comObject WScript.Shell;
        #WScript.Shell object creates a shortcut that requires the .lnk or .url extension.
        $Shortcut = $WshShell.CreateShortcut($createShortcut);
        $Shortcut.TargetPath = $targetPath;
        $Shortcut.WorkingDirectory = $workingDirectory;
        if($null -ne $arguments){
            $Shortcut.Arguments = $arguments;
        }
        $Shortcut.Save();
        #If a shortcut already exists then the function just changes the settings that are specified.
    }

    Write-Host "`n`nShortcuts";
    Write-Host "`r---------" -NoNewline;
    $AvailableFiles = $AvailableFiles.ForEach({[System.IO.Path]::ChangeExtension($_,"lnk")})
    #Prints the shortcuts to the console.
    Get-Item "$Global:reqPathToSendFiles\*" -Include $AvailableFiles;
}

function New-SymbolicLink([array] $AvailableFiles)
{
    $Global:reqPathToSendFiles = Read-HostPath "`nEnter the complete path in which you want to send the item(s) `n";
    foreach($file in $AvailableFiles)
    {
        $pathSym = [System.IO.Path]::Combine($Global:reqPathToSendFiles,$file);
        $targetSym = [System.IO.Path]::Combine($Global:reqPathToWork,$file);

        #Symbolic links require Admin Privileges. 
        #Path determines where the symlink will be created, and target determines what file is 
        #referencing to  make the symlink.
        New-Item -ItemType SymbolicLink -Path $pathSym -Target $targetSym -Force > $null;
    }

    Write-Host "`n`nSymlinks";
    Write-Host "`r--------" -NoNewline;
    #Prints the symbolic links in the console.
    Get-Item "$Global:reqPathToSendFiles\*" -Include $AvailableFiles;
}


function Get-AvailableFiles
{
    Write-Host "`n";
    Write-Host "=====================================================" -ForegroundColor Black -BackgroundColor White;
    
    [array] $AvailableFiles = @();  
    
    #If reqPathToWork has data, then it does not enter here.
    if ($null -ne $Global:reqPathToWork) {
        do {
            $keepFiles = Read-Host "`nDo you want to keep both, the file location and file specification? [Y/N] "
        } while ($keepFiles -notmatch "[yYnN]")
    } 
    
    #If the host decide to change variables after wanting to create something more, it enters here.
    #Also, if it is the first time, it will enter by default because keepFiles is empty.
    if ($keepFiles -notmatch "[yY]") 
    {
        $Global:reqPathToWork = Read-HostPath "`nEnter the complete path where the files are `n";
        $Global:reqFileSpecification = Read-Host "`nEnter any file specification.
            `rName, extension, or both. (wildcards NOT recommended | regex active).
            `rIf nothing, press enter ";
    }
    
    do{
        [bool] $entriesChanged = $false;
        do{
            #Gets all files inside the requested directory, 
            #but only those files that satisfy the file extension that the host request
            [array] $fileNames = (Get-Item "$Global:reqPathToWork\*" |
             Where-Object Name -Match $Global:reqFileSpecification | Select-Object -ExpandProperty Name);

            #If there are no files with the requested path and file extenison, then the function repeats
            #And it is requested to the host to change the values of one of the requests
            #until at least one file appear as usable.
            if($fileNames.Count -eq 0)
            {
                Write-Host "No files with the requests you enter exist here. Change them." -ForegroundColor Red;
                #Gives the host the chance to change one of the entries.
                Get-Entries; 
            }
        }while($fileNames.Count -eq 0)
        
        do{
            Write-Host "`n`nFiles";
            Write-Host "`r-----";
            Out-Host -InputObject $fileNames;
            #Gives the option to the host to create for all the files.
            Write-Host "`n   ALL: To create for all of them";
            Write-Host "`r   [*]: To change the file entries";
            
            #Gets the name of the file that the host wants to use.
            [string]$requestedFileName = Read-Host "`nEnter the FULL name of the file from which you want to create.
                `rSeparate it in commas if you want multiple files.`n";
            
            if($requestedFileName -match ",")
            {
                #It gets the files splited by the commas separators
                #And it trims them for any SPACE character that encounters at the edges of the string.
                [array] $requestedFileNames = ($requestedFileName.Split(",") | ForEach-Object{$_.Trim()});
                
                #It filles the variable with the files prematurely to not add them one by one.
                $AvailableFiles = $requestedFileNames;
                #Use of requestedFileNames in foreach because AvailableFiles could became empty.
                foreach($reqFileName in $requestedFileNames)
                {
                    #If one of the files does not exits. AvailableFiles is emptied.
                    #And the function is escaped to later return in the do..until loop.
                    if($reqFileName -notin $fileNames){
                        Write-Host "`n""$reqFileName"" doesn't exist in the current directory." -ForegroundColor Red;
                        $AvailableFiles = @();
                        break;
                    }
                } 
            }
            #If just one file is specified, it enters here.
            else{
                #It trims for any SPACE character that encounters at the edges of the string.
                $requestedFileName = $requestedFileName.Trim();

                #If host typed 'ALL' all files are passed to AvailableFiles.
                if($requestedFileName.ToUpper() -eq "ALL"){
                    $AvailableFiles = $fileNames;
                }
                #If the file is specified is inside the fileNames list, it passes to AvailableFiles.
                elseif($requestedFileName -in $fileNames){
                    $AvailableFiles = $requestedFileName;
                }
                #If host typed '*' Get-Entries function is invoke,
                #And the funtion will later repeat in the do..while loop that repeats almost the entirety of this funtion.
                elseif($requestedFileName -eq '*'){
                    Get-Entries;
                    $entriesChanged = $true;
                }
                #If the file is specified is not inside the fileNames list, the host is warned,
                #And the funtion will repeat in the do..until loop.
                else{
                    Write-Host "`n""$requestedFileName"" doesn't exist in the current directory." -ForegroundColor Red;
                }
            }
        }until(($AvailableFiles.Length -gt 0) -or $entriesChanged)
    }while($entriesChanged)

    return $AvailableFiles;
}


function Get-Entries
{
    #These lines prints the current host's entries to the console.
    Write-Host "`nEntries" -ForegroundColor Yellow;
    Write-Host "`r-------" -ForegroundColor Yellow;
    Write-Host "Path location of files: $Global:reqPathToWork" -ForegroundColor Yellow;
    Write-Host "File specification: $Global:reqFileSpecification`n" -ForegroundColor Yellow;

    #The host will be able to choose what to change.
    $message = "============What do you want to change?=============";
    $workingPath = New-Object System.Management.Automation.Host.ChoiceDescription "&Working Path", "To change the working path.";
    $fileSp = New-Object System.Management.Automation.Host.ChoiceDescription "File &Specification", "To change the file specification.";
    
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($workingPath, $fileSp);
    #This creates a better prompt for choosing.
    $requestedChangeOption = $host.ui.PromptForChoice("Change",$message, $options, -1);

    switch($requestedChangeOption)
    {
        0 {$Global:reqPathToWork = Read-HostPath "`nEnter the complete path in which you want to work in `n";}
        1 {$Global:reqFileSpecification = Read-Host "`nEnter any file specification.
              `rName, extension, or both. (wildcards NOT recommended | regex active).
              `rIf nothing, press enter ";}
    }
}

function Read-HostPath([string] $message)
{
    #Read-HostPath is similar to Read-Host command, but for testing paths.
    do{
        [string] $pathRequested = Read-Host $message;
    }while([string]::IsNullOrEmpty($pathRequested))

    #Proves that the path entry exits, and if not then it is said to the host to reenter the path.
    if(-not (Test-Path $pathRequested))
    {
        do{
            Write-Host "Your path does not exist" -ForegroundColor Red;
            $pathRequested = Read-Host "Re-enter your complete path";

          #Repeats the function if the path still not exist after reenter.
        }while(-not (Test-Path $pathRequested))
    }
    #And returns the path to the variable that requested.
    return [System.IO.Path]::GetFullPath($pathRequested);
}




# ============================================================================================

################################# Start Main Program #########################################	



#Checks if the current Powershell session is run as Admin.
$hostPriviledges = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent());
$Global:isAdmin = $hostPriviledges.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

#Global variables.
$Global:reqPathToWork, $Global:reqPathToSendFiles, $Global:reqFileSpecification = $null;

do{
    $message = "============What do you want to create?=============";
    $envVar = New-Object System.Management.Automation.Host.ChoiceDescription "&Environmental Variable", "To create an environmental variable.";
    $shortcut = New-Object System.Management.Automation.Host.ChoiceDescription "&Shortcut", "To create a shortcut.";
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($envVar, $shortcut);
    #If host is not Admin, the option for Symlink will not appear, because it requires Admin priviledges.
    if($Global:isAdmin){
      $symlink = New-Object System.Management.Automation.Host.ChoiceDescription "Symbolic &Link", "To create a symbolic link.";
      $options += [System.Management.Automation.Host.ChoiceDescription[]]($symlink);
    }
    #This creates a better prompt for choosing.
    $requestedOption = $host.ui.PromptForChoice("Creation",$message, $options, -1);

    switch($requestedOption)
    {
        0 { #EnvironmentalVariables don't require a file,
            #so Get-AvailableFiles function is not use and no parameters are needed.
            New-EnvironmentalVariable;
        }
        1 { #Shortcuts do require a file to reference.
            [array] $availableFiles = Get-AvailableFiles;
            New-Shortcut $availableFiles;
        }
        2 { #SymbolicLinks do require a file to reference.
            [array] $availableFiles = Get-AvailableFiles;
            New-SymbolicLink $availableFiles;
        }
    }

    #Opens the option to the host to create something else after the just created items.
    Write-Output "";
    $repeatRequest = $(Write-Host "Do you want to create something else? [Y/N]:" -BackgroundColor Black -ForegroundColor Yellow -NoNewline; Read-Host);

  #If the host says yes then almost the entire program will start again.
  #Except for the host initial entries that will stay the same if the host decides to.
}while($repeatRequest -match "[yY]")

#Eliminates all script global variables.
Remove-Variable -Scope Global -Name isAdmin,reqPathToWork,reqPathToSendFiles,reqFileSpecification;

#END of the script