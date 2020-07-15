#!/bin/bash

directory='<yoursever>
version=`cat assets/version.json  | grep version | sed -e 's/\s\+//g' -e 's/"//g' -e 's/\./_/g' | awk -F: '{print $NF}'`
name="script_auto_checker_v${version}"

echo "COPY ASSETS"

grep error ams.py | grep logger | sed 's/^\s*\s//g' | sed 's/^logger\.error("//g' | sort -u > CHECK_LIST.txt

list="assets log_json template README.txt CHECK_LIST.txt"
new_path="`find ./dist -mindepth 1 -maxdepth 1`"

for f in `echo $list`
do
	echo "Copy file: $f"
	cp -rf $f ${new_path}
done

if [ -e ${directory}/${name} ]
then
  echo "Remove ${directory}/${name}"
  rm -rf ${directory}/${name}
fi

echo "Copy $name"
cp -rf ${new_path} ${directory}/${name}

