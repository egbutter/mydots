## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)


$psprof = "$env:systemroot\system32\WindowsPowerShell\v1.0"
$psmod = "$env:systemroot\system32\WindowsPowerShell\v1.0\Modules"

$xlpath = "$env:appdata\Microsoft\Excel\XLSTART"

$vimpath = $null 
Switch -regex ($env:path.Split(";")) 
{
    "C:.*vim73" { 
        echo $._
        $vimpath = $_.Substring( 0, $_.Length - $_.LastIndexOf("vim73") + 1 )
        break
    }
} 
If (!$vimpath) { Throw "Please add vim73 to your %PATH%" }
echo "path"
echo $env:path.Split(";")
echo "vim"
echo ($vimpath)

function LinkFile ([string]$tolink) {

    $source="$($pwd)/$(tolink)"
    echo $source

    If ($tolink -eq "_vim") { 
        $target="$($vimpath)\$($tolink)files"  # peculiar win32 gvim inconsistency
    } Else {
        $target="$($vimpath)\$($tolink)"
    }

    If (Test-Path $target) { move $target "$target.bak" }

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

# make sure posh-git is set up
. (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1")

# make sure pyflakes is up to date
try {
    easy_install pyflakes --upgrade
} catch {
    Throw "Please install python 27, setuptools, pyflakes first!"
}

# update all the submodules
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update
