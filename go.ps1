## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)

set-executionpolicy remotesigned
$myhome = $env:userprofile  # bc forcing $home to something else is a pain


# explicitly set currentuserallhosts to %userprofile%
$psprof = $profile.currentuserallhosts
if (test-path $psprof) { move $psprof "$($psprof).bak" }
cmd /c mklink /H $psprof "$($myhome)\.psprofile.ps1"
echo "pointed powershell currentuserallhosts to userprofile"


# make sure the env vars are set correctly, inconsistent on win32 
$psmods = join-path $myhome ".psmodules"
if ($env:psmodulepath.split(";") -notcontains $psmods) {
    $env:psmodulepath += $psmods
    [Environment]::SetEnvironmentVariable("psmodulepath", $psmods, "User")
    echo "pointed powershell psmodules to userprofile"
}

$vimpath = join-path $myhome ".vim"
if ($env:vimruntime.split(";") -notcontains $vimpath) {
    $env:vimruntime += $vimpath
    [Environment]::SetEnvironmentVariable("vimruntime", $vimpath, "User")
    echo "pointed vimruntime to userprofile"
}

$xlpath = join-path $myhome ".xlstart"
$xlappdata "$env:appdata\Microsoft\Excel\XLSTART"
if (test-path $xlappdata) 
{
    move $xlappdata "$($xlappdata).bak"
    junction -s $xlpath $xlappdata
    echo "pointed XLSTART to userprofile"
}


(new-object Net.WebClient).DownloadString("http://psget.net/GetPsGet.ps1") | iex
import-module psget

install-module posh-git
install-module posh-hg
install-module pscx
install-module poshcode
install-module find-string
install-module psurl

#BROKEN: get-poshcode cmatrix -Destination $psmods
# "error proxycred" https://getsatisfaction.com/poshcode/topics/error_with_get_poshcode


function LinkFile ([string]$tolink) {

    $source="$(join-path $pwd $tolink)"

    $dotlink = $tolink -replace "^_", "."
    $target="$(join-path $myhome $dotlink)"

    if (test-path $target) { move $target "$($target).bak" }

    if ($target.psiscontainer) {
        echo "junction -s $target $source"
        junction -s $target $source
    } else {
        echo "cmd /c mklink /H $target $source"
        cmd /c mklink /H $target $source
    }
}

# here we handle the command line args
Switch -regex ($args) 
{
    "vim" {
        Foreach ( $i in "_vim*" ) {
            LinkFile $i
        }
    }  
    "ps" {
        Foreach ($i in "_ps*") {
            LinkFile $i
        }
    }
    default {
        Foreach ($i in "_*") {
            LinkFile $i
        }
    }
}

# load virtualenvwrapper-powershell else throw error
try {
    import-module virtualenvwrapper
} catch {
    try {
        echo "trying to install virtualenvwrapper ..."
        pip install virtualenvwrapper-powershell
    } catch {
        echo "WARNING: python and/or pip and/or virtualenvwrapper-powershell not installed"
    }
}

# make sure pyflakes is up to date
try {
    pip install pyflakes
} catch {
    echo "WARNING: python and/or pip and/or pyflakes not installed, some vim plugins will not work"
}

# update all the submodules
cd "$($myhome)\vimfiles"
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

set-location $myhome -PassThru
set-variable -name HOME -value $env:USERPROFILE -force  # may not persist sessions
