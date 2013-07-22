## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)

set-executionpolicy remotesigned
set-variable -name HOME -value $env:USERPROFILE -force

try {
    import-module psget
} catch {
    (new-object Net.WebClient).DownloadString("http://psget.net/GetPsGet.ps1") | iex
}

install-module posh-git -Destination "$($home)/.psmods"
install-module posh-hg -Destination "$($home)/.psmods"
install-module pscx -Destination "$($home)/.psmods"
install-module poshcode -Destination "$($home)/.psmods"
install-module find-string -Destination "$($home)/.psmods"
install-module psurl -Destination "$($home)/.psmods"

# get-poshcode cmatrix -Destination "$($home)/.psmods"
# error proxycred https://getsatisfaction.com/poshcode/topics/error_with_get_poshcode

$psprof = $profile.currentuserallhosts
if (test-path $psprof) { move $psprof "$($psprof).bak" }
junction $psprof "$($pwd)\_psprofile.ps1"

$psmods = $env:psmodulepath
$psmods += ";$($home)/.psmods"
[Environment]::SetEnvironmentVariable("PSModulePath", $psmods)

#$xlpath = "$env:appdata\Microsoft\Excel\XLSTART"


if if (!(test-path "$($home)\vimfiles")) { mkdir "$($home)\vimfiles" }

function LinkFile ([string]$tolink) {

    $source="$($pwd)\$(tolink)"

    If ($tolink -eq "_vim") { 
        $target="$($home)\vimfiles"  # peculiar win32 gvim inconsistency
    } Else {
        $target="$($home)\$($tolink)"
    }

    If (Test-Path $target) { move $target "$($target).bak" }

    junction $target $source
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
        pip install virtualenvwrapper-powershell
    } catch {
        Throw "ERROR: please pip install virtualenvwrapper-powershell"
    }
}

# make sure pyflakes is up to date
try {
    pip install pyflakes
} catch {
    Throw "ERROR: please pip install pyflakes"
}

# update all the submodules
cd "$($home)\vimfiles"
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

set-location $home -PassThru
