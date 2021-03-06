#!/usr/bin/env bash

if [[ "$#" != "3" ]]; then
    echo "usage: $0 $DOC1"".dox $DOC2.dox output_diff.diff"
    exit 1
fi

DOC1=$1
DOC2=$2
DIFF_FILE_NAME=$3

export TMPDIR=`pwd`
TEMP_DIR=`mktemp -d "${TMPDIR:-/tmp}/tmp.XXXXXXXXX"`
echo "temp dir: $TEMP_DIR";

if [[ "$TEMP_DIR" == "" ]]; then
	echo "temp dir creation / failed??"
	exit 1
fi

function unpack_doc () {
    DOC_NAME=$1
    TMP_FOLDER_NAME=$2
    OUTPUT_FILE_NAME=$3
    unzip $DOC_NAME -d $TMP_FOLDER_NAME
    DOCUMENT_XML_IN="$TMP_FOLDER_NAME/word/document.xml"
    xmllint --format $DOCUMENT_XML_IN -o $OUTPUT_FILE_NAME
}

DOC1_CONTENT_FOLDER="$TEMP_DIR/$DOC1""_content"
DOC2_CONTENT_FOLDER="$TEMP_DIR/$DOC2_content"
DOC1_FORMATTED="$TEMP_DIR/$DOC1""_docxml_formatted.xml"
DOC2_FORMATTED="$TEMP_DIR/$DOC2""_docxml_formatted.xml"

mkdir -p $DOC1_CONTENT_FOLDER
mkdir -p $DOC2_CONTENT_FOLDER
.
unpack_doc $DOC1 "$TEMP_DIR/$DOC1""_content" $DOC1_FORMATTED
unpack_doc $DOC2 "$TEMP_DIR/$DOC2""_content" $DOC2_FORMATTED

diff -C 12 $DOC1_FORMATTED $DOC2_FORMATTED > $DIFF_FILE_NAME

