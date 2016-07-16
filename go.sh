#!/usr/bin/env bash
#kudos sontek https://github.com/sontek/dotfiles

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
git submodule update
git submodule foreach git pull origin master
git submodule foreach git submodule init
git submodule foreach git submodule update

curl -L https://raw.githubusercontent.com/yyuu/pyenv-installer/master/bin/pyenv-installer | bash
