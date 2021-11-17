#!/bin/bash
#set -x

cwd="`pwd`"
srcdir="${cwd}/../com/fourjs/spreadsheet"
javafile="${srcdir}/CopySheets.java"

if [ ! -d "$srcdir" ]; then
   echo "You should be in the script directory when running this script"
   exit 1
fi

if [ ! -f "$javafile" ]; then
   echo "The java source file $javafile does not exist"
   exit 1
fi

javac "$javafile"

if [ $? -ne 0 ]; then
   echo "Compilation error occurred"
   exit 1
fi

jarroot="${cwd}/.."
if [ ! -d "$jarroot" ]; then
   echo "You should be in the script directory when running this script"
   exit 1
fi

cd "$jarroot"
classfiles=`ls -1 com/fourjs/spreadsheet/*.class com/fourjs/spreadsheet/*.java`

jar cMf spreadsheet-copy-fourjs.jar META-INF/MANIFEST.MF $classfiles
if [ $? -ne 0 ]; then
   echo "An error occurred attempting to create the jar file"
   exit 1
fi

echo "Jar file has been created successfully"
