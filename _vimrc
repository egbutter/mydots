filetype off "must be off to call pathogen bundles
call pathogen#runtime_append_all_bundles()
call pathogen#helptags()
filetype on "add filetype on top
filetype plugin indent on 
syntax enable "always keep syntax highlighting on
set background=dark
let g:solarized_termtrans=1
let g:solarized_termcolors=256
let g:solarized_contrast="high"
let g:solarized_visibility="high"
colorscheme blue
set title "put title on top of vim
set ruler "line nums
set history=1000 "increase buffer history
set number
set wildmode=list:longest "bashlike file switch
setlocal tabstop=4 "tab=4spaces
setlocal shiftwidth=4
set expandtab
set autoindent
set smartindent
set ignorecase "ignore case in search unless capitalized
set smartcase
set hlsearch "highlight search
set incsearch "dynamic search
set backspace=indent,eol,start "backspace how we would expect
set foldmethod=indent "enable code folding
set foldlevel=99
"quicker movement between windows
map <c-j> <c-w>j
map <c-k> <c-w>k
map <c-l> <c-w>l
map <c-h> <c-w>h
let mapleader=","
map <leader>todo <Plug>TaskList 
map <leader>undo :GundoToggle<CR>
map <leader>v :sp .vimrc
map <leader>V :source .vimrc
let g:pyflakes_use_quickfix = 0
let g:pep8_map='<leader>pep8'
let g:slime_target = "screen"
au FileType python set omnifunc=pythoncomplete#Complete "tab completion 
let g:SuperTabDefaultCompletionType="context"
set completeopt=menuone,longest,preview "python documentation on pw
set runtimepath^=.vim/bundle/ctrlp.vim
"include code completion from venvs for python
py << EOF
import os.path
import sys
import vim
if 'VIRTUAL_ENV' in os.environ:
    project_base_dir = os.environ['VIRTUAL_ENV']
    sys.path.insert(0, project_base_dir)
    activate_this = os.path.join(project_base_dir, 'bin/activate_this.py')
    execfile(activate_this, dict(__file__=activate_this))
EOF
set ff=unix
"check if nix or dos then :set ff=unix and :set ff=dos
"what's up with the weird commands like shift+control+v block visual instead of
"control v?  get this fixed FIXME
