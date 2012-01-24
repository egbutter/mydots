## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)

$vimpath = $null
Switch -regex ( $env:path.Split(";") ) {
    "C:.*vim73" { 
        $vimpath = $_.Substring( 0, $_.Length - $_.LastIndexOf("vim73") + 1 )
        break
    }
} 
If (!$vimpath) { Throw "Please add vim73 to your \$PATH" }

function LinkFile ([string]$tolink) {

    $source="$($PWD)/$(tolink)"
    If ($tolink -eq "_vim") { 
        $target="$($vimpath)\$($tolink)files"  # peculiar win32 gvim inconsistency
    } Else {
        $target="$($vimpath)\$($tolink)"
    }

    If (Test-Path $target ) { move $target $target".bak" }

    mklink /D $target $source
}

If ( ( $args[0] -eq "vim" ) {
    For ( $i in _vim* ) {
        link_file $i
    }
} 

# setup git submodules
git.cmd submodule sync
git.cmd submodule init
git.cmd submodule update
git.cmd submodule foreach git pull origin master
git.cmd submodule foreach git submodule init
git.cmd submodule foreach git submodule update

# setup command-t
cd _vim/bundle/command-t
rake make
