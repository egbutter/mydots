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


# make sure the env vars are set correctly, inconsistent on win32 
$psmods = join-path $myhome ".psmodules"
if (-not $env:psmodulepath -or ($env:psmodulepath.split(";") -notcontains $psmods)) {
    $env:psmodulepath += $psmods
    [Environment]::SetEnvironmentVariable("psmodulepath", $env:psmodulepath, "User")
    echo "pointed powershell psmodules to userprofile:"
    echo $env:psmodulepath
}

$vimpath = join-path $myhome ".vim"
if (-not $env:vimruntime -or ($env:vimruntime.split(";") -notcontains $vimpath)) {
    $env:vimruntime += $vimpath
    [Environment]::SetEnvironmentVariable("vimruntime", $env:vimruntime, "User")
    echo "pointed vimruntime to userprofile:"
    echo $env:vimruntime
}



$modules = get-module -list
if ($modules -notcontains "psget") 
{
    echo "downloading psget"
    (new-object Net.WebClient).DownloadString("http://psget.net/GetPsGet.ps1") | iex
}
import-module psget

install-module posh-git
install-module posh-hg
install-module pscx
install-module find-string
install-module psurl

install-module poshcode
#BROKEN: get-poshcode cmatrix -Destination $psmods
# "error proxycred" https://getsatisfaction.com/poshcode/topics/error_with_get_poshcode

# load virtualenvwrapper-powershell else throw error
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

# linking the files using mklink for files and junction for dirs
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
        echo "moving $target to $bakfile"; 
        mv $target $bakfile -force; 
    }

    if ($target.psiscontainer) 
    {
        echo "junction -s $target $source"
        junction -s $target $source
    } else {
        echo "cmd /c mklink /H $target $source"
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
$targets | foreach-object { linkfile $_.name }


# update all the submodules in .vim directory 
cd $vimpath
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

