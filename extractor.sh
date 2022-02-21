#!/bin/bash 

#TODO: Have the option to extract specific file types such as word doc with user args
#TODO: get multiple files working at once rather then running script over and over
#TODO: have nonconflicting file names when extracting multiple files in outout folder
	#Have Arg name plus output to show difference? 

#Select file

Fileone=$1 

rm -r OutPut
mkdir OutPut 

#extracting files 
 
cp ${Fileone} ${Fileone}.zip 
unzip -q ${Fileone}.zip 
rm ${Fileone}.zip
cp ./xl/embeddings/* ./OutPut/ 

#Rename files

#cleaning up extra files
rm \[Content_Types\].xml 
rm -r docProps/  
rm -r _rels/   
rm -r xl/ 

echo "Successfully extracted all files!"
