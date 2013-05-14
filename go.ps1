## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)

$vimpath = $null 

($env:path.Split(";")) 

Switch -regex ($env:path.Split(";")) 
    {
        "vim73" { 
            $vimpath = $_.Substring( 0, $_.Length - $_.LastIndexOf("vim73") + 1 )
            break
        }
    } 
If (!$vimpath) { Throw "Please add vim73 to your %PATH%" }

function LinkFile ([string]$tolink) {

    $source="$($PWD)/$(tolink)"
    If ($tolink -eq "_vim") { 
        $target="$($vimpath)\$($tolink)files"  # peculiar win32 gvim inconsistency
    } Else {
        $target="$($vimpath)\$($tolink)"
    }

    If (Test-Path $target ) { move $target $target".bak" }

    junction $target $source
}

If ( $args[0] -eq "vim" ) {
    Foreach ( $i in _vim* ) {
        LinkFile $i
    }
} 

# make sure posh-git is set up
. (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1")
# update all the submodules
git. (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1").cmd submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

# setup command-t
#cd _vim/bundle/command-t
#rake make
