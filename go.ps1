##############################################################################
##
## go.ps1
##
## by ericbeurre
##
##############################################################################

<#
.SYNOPSIS

Bootstrap my powershell environment

.EXAMPLE

go.ps1 vim,ps,xl

#>

Param
(
    [String] $gomode="all"
)

set-executionpolicy remotesigned 
$myhome = $env:userprofile  # bc forcing $home to something else is a pain  
$env:home = $myhome; 
[Environment]::SetEnvironmentVariable("HOME", $env:home, "User"); 


#
# make sure the env vars are set correctly, inconsistent on win32 
#

# this is just for my own psmodules like utils.psm1
$psmods = join-path $myhome ".psmodules"
if (-not $env:psmodulepath -or ($env:psmodulepath.split(";") -notcontains $psmods)) {
    $env:psmodulepath += $env:psmodulepath 
}

# a lot of scripts e.g., virtualenvwrapper-powershell, assume this location is available
$psmdefault = join-path $myhome "Documents\WindowsPowerShell\Modules"
if ($env:psmodulepath.split(";") -notcontains $psmdefault) {
    $env:psmodulepath += $psmdefault  # at the front, this gets psget files, etc.
}
[Environment]::SetEnvironmentVariable("psmodulepath", $env:psmodulepath, "User")
echo "cleaned up powershell psmodule default paths:"
echo $env:psmodulepath

$vimpath = join-path $myhome ".vim"
if (-not $env:vimruntime -or ($env:vimruntime.split(";") -notcontains $vimpath)) {
    $env:vimruntime += $vimpath
    [Environment]::SetEnvironmentVariable("vimruntime", $env:vimruntime, "User")
    echo "pointed vimruntime to userprofile:"
    echo $env:vimruntime
}


#
# linking the files using mklink for files and junction for dirs
#

function LinkFile ([string]$tolink) 
{
$source = join-path $pwd $tolink; 
    
    if ($tolink -eq "_xlstart") 
    {
        $xlappdata = "$env:appdata\Microsoft\Excel\XLSTART"; 
        $target = $xlappdata; 
    } elseif ($tolink -eq "_psprofile.ps1")  {
        $psprof = $profile.currentuserallhosts; 
        $target = $psprof; 
    } else { 
        $dotlink = $tolink -replace "^_", "."; 
        $target = join-path $myhome $dotlink; 
    }

    if (test-path $target) 
    { 
        $bakfile = "$($target).bak"; 
        echo "removing last .bak, copying $target to $bakfile"; 
        if (test-path $bakfile) { rm $bakfile -recurse; }  
        copy-item -path $target -destination $bakfile -recurse; 
        if ((Get-Item $target).psiscontainer) { junction -d $target } else { rm $target }

    }

    if ((get-item $source).psiscontainer) 
    {
        echo "directory junction $target to $source"
        junction -s $target $source 
    } else {
        echo "file junction $target to $source"
        cmd /c mklink /H $target $source 
    }
}

# here we handle the command line gomode, if none goes to default
if ($gomode) 
{ 
    $gomode = $gomode.split(","); 
    echo $gomode;  
} else { 
    $gomode = "all"; 
    echo "All"; 
}

# get our targets 
Switch ($gomode) 
{
    "vim" { $targets += ls (join-path $pwd _vim*) } 
    "ps" { $targets += ls (join-path $pwd _ps*) } 
    "xl" { $targets += ls (join-path $pwd _xl*) } 
    default { $targets = ls (join-path $pwd _*) } 
}

# link the targets  
$targets | where-object {$_.name -notlike "*bash*"} | foreach-object { linkfile $_.name }


#
# download powershell modules 
#

$modules = get-module -list | ForEach-Object { $_.name }
if ($modules -notcontains "psget") 
{
    echo "downloading psget"
    (new-object Net.WebClient).DownloadString("http://psget.net/GetPsGet.ps1") | iex
}
import-module psget

install-module posh-git -destination $psmods -EA silentlycontinue
install-module posh-hg -destination $psmods -EA silentlycontinue
install-module pscx -destination $psmods -EA silentlycontinue
install-module find-string -destination $psmods -EA silentlycontinue
install-module psurl -destination $psmods -EA silentlycontinue
install-module poshcode -destination $psmods -EA silentlycontinue

#BROKEN: get-poshcode cmatrix -Destination $psmods
# "error proxycred" https://getsatisfaction.com/poshcode/topics/error_with_get_poshcode


#
# load virtualenvwrapper-powershell else throw error
#

if ($modules -notcontains "virtualenvwrapper") 
{
    try {
        echo "trying to install virtualenvwrapper ..."
        pip install virtualenvwrapper-powershell
    } catch {
        echo "WARNING: python and/or pip and/or virtualenvwrapper-powershell not installed"
    }
} 
$workonhome = join-path $myhome ".envs"
if (-not $env:workon_home -or $env:workon_home -ne $workonhome) 
{
    $env:workon_hom = $workonhome
    [Environment]::SetEnvironmentVariable("workon_home", $workonhome, "User")
}


# make sure pyflakes is up to date
try 
{
    pip install pyflakes
} catch {
    echo "WARNING: python and/or pip and/or pyflakes not installed, some vim plugins will not work"
}


#
# update all the submodules in .vim directory 
#

cd $vimpath
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

