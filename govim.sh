#!/usr/bin/env bash
#kudos sontek https://github.com/sontek/dotfiles.git

# if cygwin, replace $HOME=/home/username use /cygdrive/c/vim/
# _vimrc -> /cygdrive/c/vim/_vimrc  
# _vim -> /cygdrive/c/vim/vimfiles
# if cygwin, do not forget d2u to run scripts


function link_file {
    source="${PWD}/$1"
    target="${HOME}/${1/_/.}"

    if [ -e "${target}" ]; then
        mv $target $target.bak
    fi

    ln -sf ${source} ${target}
}

if [ "$1" = "vim" ]; then
    for i in _vim*
    do
        link_file $i
    done
else
    for i in _*
    do
        link_file $i
    done
fi

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
