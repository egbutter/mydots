## this is a Powershell version of eponymous bash script
## just handles vim for now -- this is urgent ... extend later ...
## this is still windows, so needs dos endings (no d2u!)

$sysroot = [environment]::GetEnvironmentVariable("SystemRoot")
if (!$sysroot) { Throw "Please set the %SystemRoot% environment variable" }

$psprof = "$sysroot\system32\WindowsPowerShell\v1.0"
$psmod = "$sysroot\system32\WindowsPowerShell\v1.0\Modules"

$vimpath = $null 

Switch -regex ($env:path.Split(";")) 
    {
        "C:.*vim73" { 
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

    $source="$($PWD)/$(tolink)"
    echo $source
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
} Else {
    LinkFile ${env:windir}"\system32\WindowsPowerShell\v1.0"
}

# make sure posh-git is set up
. (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1")
# update all the submodules
git submodule sync
git submodule init
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

# setup command-t
#cd _vim/bundle/command-t
#rake make
