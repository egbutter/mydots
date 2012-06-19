#!/bin/sh 
# Find Replace Recursive
find /path -type f -iname "*.txt" -exec sed -i.bak 's/foo/bar/g' "{}" +;
