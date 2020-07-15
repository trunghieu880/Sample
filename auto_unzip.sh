#!/bin/bash -x

path_source="$1"
name_source=`basename ${path_source}`

Tool_7z='/c/Program Files/7-Zip/7z.exe'

type="zip"

echo "${Tool_7z}"

if [ -e "${path_source}" ]
then
  echo "Extract ${path_source}"
else
  exit 1
fi

/c/Program\ Files/7-Zip/7z.exe x -aoa -bd ${path_source}

