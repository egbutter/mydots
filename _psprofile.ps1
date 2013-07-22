set-location $env:userprofile -PassThru

import-module psget
import-module poshcode
import-module pscx
import-module posh-hg
import-module posh-git
import-module find-string
import-module psurl
import-module virtualenvwrapper

set-alias cat        get-content
set-alias cd         set-location
set-alias clear      clear-host
set-alias cp         copy-item
set-alias h          get-history
set-alias history    get-history
set-alias kill       stop-process
set-alias lp         out-printer
set-alias ls         get-childitem
set-alias mount      new-mshdrive
set-alias mv         move-item
set-alias popd       pop-location
set-alias ps         get-process
set-alias pushd      push-location
set-alias pwd        get-location
set-alias r          invoke-history
set-alias rm         remove-item
set-alias rmdir      remove-item
set-alias echo       write-output

set-alias cls        clear-host
set-alias chdir      set-location
set-alias copy       copy-item
set-alias del        remove-item
set-alias dir        get-childitem
set-alias erase      remove-item
set-alias move       move-item
set-alias rd         remove-item
set-alias ren        rename-item
set-alias set        set-variable
set-alias type       get-content
set-alias which      get-command

function grep 
{
    get-childitem $args[0] -include $args[1] -rec | select-string -pattern $args[2:]
}

function tail
{
    get-content $args[0] -tail 30 -wait
}

function help
{
    get-help $args[0] | out-host -paging
}

function man
{
    get-help $args[0] | out-host -paging
}

function mkdir
{
    new-item -type directory -path $args
}

function md
{
    new-item -type directory -path $args
}

function prompt
{
    "PS " + $(get-location) + "> "
}

function pro 
{ 
    vim $profile 
}

function prompt 
{
	$path = ""
	$pathbits = ([string]$pwd).split("\", [System.StringSplitOptions]::RemoveEmptyEntries)
	if($pathbits.length -eq 1) {
		$path = $pathbits[0] + "\"
	} else {
		$path = $pathbits[$pathbits.length - 1]
	}
	$userLocation = $env:username + '@' + [System.Environment]::MachineName + ' ' + $path
	$host.UI.RawUi.WindowTitle = $userLocation
    Write-Host($userLocation) -nonewline -foregroundcolor Green 

	Write-Host('>') -nonewline -foregroundcolor Green    
	return " "
}


& {
    for ($i = 0; $i -lt 26; $i++) 
    { 
        $funcname = ([System.Char]($i+65)) + ':'
        $str = "function global:$funcname { set-location $funcname } " 
        invoke-expression $str 
    }
}
