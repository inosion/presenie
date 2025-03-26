#!/bin/bash


original_pptx=$1
new_pptx=$2

temp_dir=$(mktemp -d)

mkdir -p $temp_dir/orig/../new/../old-new-compare
unzip -d $temp_dir/orig ${original_pptx}
unzip -d $temp_dir/new ${new_pptx}

find $temp_dir/new -type f | xargs -Ixx sh -c 'cp xx '$temp_dir'/old-new-compare/$(echo xx | sed "s/\//_/g")'
find $temp_dir/orig -type f | xargs -Ixx sh -c 'cp xx '$temp_dir'/old-new-compare/$(echo xx | sed "s/\//_/g")'
find $temp_dir/old-new-compare -type f -name '*.xml' | xargs -Ixx sh -c 'mv xx xx.bak && xmllint --format xx.bak -o xx'

echo "Generated the comparison files in $temp_dir/old-new-compare"
 
