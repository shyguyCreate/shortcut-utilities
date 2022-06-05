#########TODO########
#--------------------
#CHECK{x} #allow -notmatch in file specification
#CHECK{x} #separate global:req messages into variables
#CHECK{x} #fix =============== color diplaying error
#CHECK{x} #add folder creation for sending items
#CHECK{x} #add support for sending items to multiple forlders
#CHECK{x} #add support for copying items
#CHECK{x} #fix item printing at the end of the funtions
#CHECK{x} #add support for match and notmatch in file specification
#CHECK{x} #separate new-env function from other functions
#CHECK{x} #join repetitive code at the start of functions into one function


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
    else{
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


#This function is to not repeat the same code in all the tree functions that are below this one.
function Start-CreationMethod([scriptblock] $Function, [array] $AvailableFiles)
{
    [bool] $createFolderBasedOnName = $false;
    
    $Global:reqPathToSendFiles = Test-CreateFolders $Global:message_reqPathToSendFiles;
    #Created for cases where host chooses to create folders based on names.
    $current_reqPathToSendFiles = $Global:reqPathToSendFiles;

    #Checks for the '?' char at the end that was left at the end if the host said yes to the creation of folders based on names.
    if ($Global:reqPathToSendFiles.EndsWith('?')) {
        $createFolderBasedOnName = $true
    }

    foreach($file in $AvailableFiles)
    {
        if ($createFolderBasedOnName) {
            #Eliminates the extension form the file.
            $tmpFileWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($file)
            #It created the folder for the name specified.
            New-Item -Path $Global:reqPathToSendFiles -Name $tmpFileWithoutExtension -ItemType Directory -Force > $null;

            #Change the variable reqPathToSendFiles to a new one that has the file name in front of it.
            $current_reqPathToSendFiles = [System.IO.Path]::Combine($Global:reqPathToSendFiles,$tmpFileWithoutExtension)
        }

        #Calls the funtion that create the item.
        Invoke-Command -ScriptBlock $Funciton -ArgumentList $current_reqPathToSendFiles,$file
    }
}

function New-Shortcut($current_reqPathToSendFiles, $file)
{
    #Replace changes the file extension to the one that shortcuts have, which is 'lnk'.
    $fileRenameExtension = [System.IO.Path]::ChangeExtension($file,"lnk");
    $createShortcut = [System.IO.Path]::Combine($current_reqPathToSendFiles,$fileRenameExtension);
    $arguments = $null;
    $targetPath = [System.IO.Path]::Combine($Global:reqPathToWork,$file);
    $workingDirectory = $current_reqPathToSendFiles;

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
        }else {
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

    #It saves the file and the path to later print them.
    $Global:createdFiles += $createShortcut;
}


function Copy-Files($current_reqPathToSendFiles, $file)
{
    #Joins the path and the name of the original file
    $filePath = [System.IO.Path]::Combine($Global:reqPathToWork,$file);
    $destinationPath = [System.IO.Path]::Combine($current_reqPathToSendFiles,$file);
    #Creates the copy, -Force is used to avoid errors in the console and -Recurse if the file is a folder.
    Copy-Item -Path $filePath -Destination $destinationPath -Recurse -Force > $null;

    #It saves the file and the path to later print them.
    $Global:createdFiles += $destinationPath;
}


function New-SymbolicLink($current_reqPathToSendFiles, $file)
{
    $pathSym = [System.IO.Path]::Combine($current_reqPathToSendFiles,$file);
    $targetSym = [System.IO.Path]::Combine($Global:reqPathToWork,$file);

    #Symbolic links require Admin Privileges. 
    #Path determines where the symlink will be created, and target determines what file is 
    #referencing to  make the symlink.
    New-Item -ItemType SymbolicLink -Path $pathSym -Target $targetSym -Force > $null;

    #It saves the file and the path to later print them.
    $Global:createdFiles += $pathSym;
}



function Get-AvailableFiles
{
    Write-Host "`n";
    Write-Host "=======================================================j" -ForegroundColor Black -BackgroundColor White;
    Write-Host "";
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
        $Global:reqPathToWork = Read-HostPath $Global:message_reqPathToWork;
        $Global:reqFileSpecification = Read-Host $Global:message_reqFileSpecification;
    }
    
    do{
        [bool] $entriesChanged = $false;
        do{
            #Copies the request to a new variable to modify its value preserving the original input.
            [string] $current_reqFileSpecification = $Global:reqFileSpecification;
            #It splits the variable if there are any commas.
            [array] $current_reqFileSpecification = $current_reqFileSpecification.Split(',');

            foreach ($fileSpecification in $current_reqFileSpecification) 
            {
                #If the host does not specify '?' to exclude, it enters here.
                if ($fileSpecification -notmatch '^(\s+)?\?') 
                {   
                    #Gets all files inside the requested directory, 
                    #but only those files that satisfy the specification that the host request.
                    [array] $fileNames = (Get-Item "$Global:reqPathToWork\*" -Include $fileNames |
                    Where-Object Name -Match $fileSpecification | Select-Object -ExpandProperty Name);
                }
                else #If the host wants to exclude files according to the specification, it enters here.
                {
                    #Eliminates the '?' char that tells to exclude files with that name.
                    $excludeSignIndex = $fileSpecification.IndexOf('?');
                    $excludeFileSpecification = $fileSpecification.Substring($excludeSignIndex + 1);
                    #Gets all files inside the requested directory, 
                    #but only those files that DO NOT satisfy the specification that the host request.
                    [array] $fileNames = (Get-Item "$Global:reqPathToWork\*" -Include $fileNames |
                    Where-Object Name -NotMatch $excludeFileSpecification | Select-Object -ExpandProperty Name);
                }
            }
            
            #If there are no files with the requested path and file extenison, then the function repeats
            #And it is requested to the host to change the values of one of the requests
            #until at least one file appear as usable.
            if($fileNames.Count -eq 0)
            {
                Write-Host "No files with the requests you enter exist here. Change them." -ForegroundColor Red;
                #Gives the host the chance to change one of the entries.
                Update-Entries; 
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
                #If host typed '*' Update-Entries function is invoke,
                #And the funtion will later repeat in the do..while loop that repeats almost the entirety of this funtion.
                elseif($requestedFileName -eq '*'){
                    Update-Entries;
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


function Update-Entries
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
        0 {$Global:reqPathToWork = Read-HostPath $Global:message_reqPathToWork}
        1 {$Global:reqFileSpecification = Read-Host $Global:message_reqFileSpecification}
    }
}


function Read-HostPath([string] $message, [string] $pathRequested = '')
{
    #Read-HostPath is similar to Read-Host command, but for testing paths.
    if ($pathRequested -eq '') {
        [string] $pathRequested = Read-Host $message;
    }

    #Eliminates any '"',"'" or white space at the beginning or end of the string.
    $pathRequested = $pathRequested.Trim('\',"'",' ');

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


function Test-CreateFolders([string] $message)
{
    #Gets the information the host enter.
    do{
        [string] $prompt_reqPathToSendFiles = Read-Host $message;
    }while([string]::IsNullOrEmpty($prompt_reqPathToSendFiles))
    
    #If there is no '$' char, then it will proceed as normal, just the path will be tested.
    if ($prompt_reqPathToSendFiles -ne '$') 
    {
        $pathRequested = Read-HostPath -message $message -pathRequested $prompt_reqPathToSendFiles;
        return $pathRequested
    }

    $message = "============New folder creation=============";
    $createFolders = New-Object System.Management.Automation.Host.ChoiceDescription "&Create folders", "It create new folders.";
    $basedNameFolders = New-Object System.Management.Automation.Host.ChoiceDescription "Folders based on each file &name", "It create folders based on each file name.";
    $bothFolders = New-Object System.Management.Automation.Host.ChoiceDescription "&Both", "It create new folders and folders based on each file name.";
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($createFolders, $basedNameFolders, $bothFolders);
    #This creates a better prompt for choosing.
    $requestedFolderOption = $host.ui.PromptForChoice("Folder",$message, $options, -1);

    $parentPath = Read-HostPath "First enter the complete parent path where the new folders will leave."

    switch($requestedFolderOption)
    {
        0 { [string] $newFolders = Read-Host "Now, enter the folder's name(s) to create
                `rif multiple, separate them with '\' char ";
            #Eliminates any '\' or white space at the beginning or end of the string.
            $newFolders = $newFolders.Trim('\',' ');
            #Joins both paths into a single.
            $parentPath = [System.IO.Path]::Combine($parentPath,$newFolders);
            return $parentPath
        }
        1 { #The '?' sign will later tell the program to create files based on names.
            return $parentPath + "?"
        }
        2 { [string] $newFolders = Read-Host "NOTE: the folders based on the names will be created afterwards.
                `rNow, enter the folder's name(s) to create
                `rif multiple, separate them with '\' char ";
            #Eliminates any '\' or white space at the beginning or end of the string.
            $newFolders = $newFolders.Trim('\',' ');
            #Joins both paths into a single.
            $parentPath = [System.IO.Path]::Combine($parentPath,$newFolders);
            return $parentPath + "?"
        }
    }
}




# ============================================================================================

################################# Start Main Program #########################################	



#Checks if the current Powershell session is run as Admin.
$hostPriviledges = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent());
$Global:isAdmin = $hostPriviledges.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

#Global variables.
$Global:reqPathToWork, $Global:reqPathToSendFiles, $Global:reqFileSpecification = $null;

#Global message variables.
[string] $Global:message_reqPathToWork = "`nEnter the complete path where the files are `n";
[string] $Global:message_reqFileSpecification = "`nEnter any specification to match with the file.
    `rEnter the string for matches, or add '`?' at the beginning for notmatches.
    `rIf you want multiple matches, sepate them with commas and follow the same rules as above.
    `rEnter name, extension, or both. (careful w/ wildcards | regex active).
    `rIf nothing, press enter ";
[string] $Global:message_reqPathToSendFiles = "`nEnter the complete path in which you want to send the item(s).
    `rIf you want to create folder(s), enter only a dollar (`$) sign `n";

#Global variable to save created files.
[array] $Global:createdFiles = @();


do{
    $message = "============What do you want to create?=============";
    $envVar = New-Object System.Management.Automation.Host.ChoiceDescription "&Environmental Variable", "To create an environmental variable.";
    $shortcut = New-Object System.Management.Automation.Host.ChoiceDescription "&Shortcut", "To create a shortcut.";
    $copy = New-Object System.Management.Automation.Host.ChoiceDescription "&File copies", "To copy files.";
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($envVar, $shortcut, $copy);
    #If host is not Admin, the option for Symlink will not appear, because it requires Admin priviledges.
    if($Global:isAdmin){
      $symlink = New-Object System.Management.Automation.Host.ChoiceDescription "Sym&links", "To create a symbolic link.";
      $options += [System.Management.Automation.Host.ChoiceDescription[]]($symlink);
    }
    #This creates a better prompt for choosing.
    $requestedOption = $host.ui.PromptForChoice("Creation",$message, $options, -1);

    #requestedOption is separated into if..else and not paased entirely into one switch because they require different needs.
    #All funtions in the else need to get that avaliable files, so they are join excluding env variables creation function.
    if ($requestedOption -eq 0) {
        #EnvironmentalVariables don't require a file.
        New-EnvironmentalVariable;
    }
    else { 
        #All these funtions do require a file to reference, so it gets them.
        [array] $availableFiles = Get-AvailableFiles;

        switch($requestedOption)
        {
            1 { Start-CreationMethod -Function ${Function:New-Shortcut} -AvailableFiles $availableFiles
                [string] $itemType = "Shortcuts";
            }
            2 { Start-CreationMethod -Function ${Function:Copy-Files} -AvailableFiles $availableFiles
                [string] $itemType = "Files copied";
            }
            3 { Start-CreationMethod -Function ${Function:New-SymbolicLink} -AvailableFiles $availableFiles
                [string] $itemType = "Symlinks";
            }
        }
        #Prints the type of items with a beloow line of '-' chars that has the same number of chars as the string.
        Write-Host "`n`n$itemType";
        Write-Host ("-" * $itemType.Length) -NoNewline;
        #Prints the items for the above funtions.
        Get-Item -Path $Global:createdFiles;
    }

    #Opens the option to the host to create something else after the just created items.
    Write-Output "";
    $repeatRequest = $(Write-Host "Do you want to create something else? [Y/N]:" -BackgroundColor Black -ForegroundColor Yellow -NoNewline; Read-Host);

    #createdFile is empty to not repeat them again if needed to print.
    $Global:createdFiles = @();

  #If the host says yes then almost the entire program will start again.
  #Except for the host initial entries that will stay the same if the host decides to.
}while($repeatRequest -match "[yY]")

#Eliminates all script global variables.
Remove-Variable -Scope Global -Name isAdmin,reqPathToWork,reqPathToSendFiles,reqFileSpecification,createdFiles;
Remove-Variable -Scope Global -Name message_reqPathToWork,message_reqPathToSendFiles,message_reqFileSpecification;

#END of the script