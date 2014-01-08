#!/bin/bash



#get file name components
SCRIPTNAME=$(basename "$0")

if [[ $# -eq 0 ]] ; then
    printf "\nOpenSesame will reset the VBA password on Microsoft Office macro-enabled files (.docm,.xlsm,.pptm).\nThe new password is 'opensesame' - quite useful if you've forgotten the original password.\nOutput is written to a new file, with -opensesame appended to the filename.\n"
    printf "\nUsage:\n$SCRIPTNAME file.xlsm\n\n"
    exit 0
fi


SOURCE=$(python -c "import os,sys; print os.path.realpath(sys.argv[1])" "$1")
SOURCEDIR=$(dirname "$SOURCE")
EXTENSION=$(echo "$SOURCE" | awk -F . '{ print tolower($NF) }')

VBAPATH="xl/vbaProject.bin"

#check extension is in xlsm, docm, set subdir for the vbaProject.bin file (xl/word)
case "$EXTENSION" in
	xlsm)
		VBAPATH="xl/vbaProject.bin";;
	docm)
		VBAPATH="word/vbaProject.bin";;
	pptm)
		VBAPATH="ppt/vbaProject.bin";;
	*)
		printf "\n$SCRIPTNAME\n$SCRIPTNAMEdoesn't know what to do with $EXTENSION files, only .docm, .xlsm, .pptm macro enabled Office files.\nSorry.\n"
		exit 1
		;;
esac		

SOURCEFN=$(basename "$SOURCE" .$EXTENSION)
OUTNAME="$SOURCEDIR/$SOURCEFN-opensesame.$EXTENSION"

WORKINGDIR=$(mktemp -d)
cd "$WORKINGDIR"

printf "\n$SCRIPTNAME\nProcessing file: $SOURCE\nOutput to:       $OUTNAME\nUsing temp dir:  $WORKINGDIR\n\n"

#unpack the office file
unzip -qq "$SOURCE"

#edit vbaProject.bin
#replace the password hashes with those for 'opensesame'
sed -i 's/CMG="[A-Z0-9]*"/CMG="0A08A6BEB4C2B4C2B0C6B0C6"/m' $VBAPATH
sed -i 's/DPB="[A-Z0-9]*"/DPB="A4A60889387955795586AB7A55A1A7BA5A1B737F64F2CC6B88E52ED7FC63FA2D016799E4E937"/m' $VBAPATH
sed -i 's/GC="[A-Z0-9]*"/GC="7F7DD3F275165117511751"/m' $VBAPATH

zip -rqq "$OUTNAME" * 

printf "Done - saved file $OUTNAME has VBA password 'opensesame'\n\n"

exit $?
