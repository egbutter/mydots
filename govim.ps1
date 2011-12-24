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

    $source="$($PWD)\$($tolink)"
    If ($tolink -eq "_vim") { 
        $target="$($vimpath)\$($tolink)files"  # peculiar win32 gvim inconsistency
    } Else {
        $target="$($vimpath)\$($tolink)"
    }

    If (Test-Path $target ) { move $target $target".bak" }

    If ( (Get-Item ".\_vim") -is [System.IO.DirectoryInfo] ) {
        $symopt = "/D"
    } Else {
        $symopt = ""
    }
    
    echo "mklink $($symopt) $($target) $($source)"
    cmd /c mklink $symopt $target $source
}

If ( $args[0] -eq "vim" ) {
    $targetregex = "^_vim.*"
} Else {
    $targetregex = "^_.*"
}

Switch -regex ( Dir ) {
    "$($targetregex)" { LinkFile $_.name }
}

# setup git submodules
cd _vim
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

# setup command-t
cd bundle/command-t
rake make
cd ../../..
