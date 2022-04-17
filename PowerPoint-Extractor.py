#!/usr/bin/env python3

import argparse
import shutil
import os
from zipfile import ZipFile

parser = argparse.ArgumentParser(
    add_help=False,
    description="Simple Python script to extract embedded powerpoints within Excel documents",
)

parser.add_argument(
    "--file",
    "-f",
    type=argparse.FileType("r"),
    default=None,
    help="Select file for extracting",
    action="store",
)

parser.add_argument(
    "--output",
    "-o",
    type=str,
    help='The output directory for the extracted documents, By defualt the documents will extract to "Output"',
    default="Output",
)

parser.add_argument("--help", "-h", action="store_true")


args = parser.parse_args()

# Custom help message

if args.help:
    print(
        "\nUsage: Powerpoint-extractor.py [-h] [--file FILE] [--output OUTPUT]\n\nSimple Python script to extract embedded powerpoints within Excel documents\n\nOptional Arguments:\n --file, -f Select file for extracting.\n --output, -o The directory where the extracted documents will be placed, By defualt the documents will extract to 'Output'.\n --help, -h Display this help message."
    )
    exit()


# Stops script if no file is provided

if args.file is None:
    print("No file was selected, To select a file please use -f")
    exit()


# Get Current Directory

outputdir = os.getcwd() + "/" + args.output + "/"


# Make directory

try:
    os.mkdir(outputdir)
except:
    print(f"A folder with that name already exist, Outputing to {args.output} \n")


outputfile = rf"{outputdir}{args.file.name}"


# Copy User file to output directory

if args.file:
    shutil.copyfile(args.file.name, outputfile)


# Rename file to zip.zip

newzip = f"{outputdir}zip.zip"
os.rename(outputfile, newzip)

# Extracts files but keep the full directory

with ZipFile(newzip, "r") as zipObj:
    listOfFileNames = zipObj.namelist()
    for fileName in listOfFileNames:
        if fileName.endswith(".pptx"):
            zipObj.extract(fileName, outputdir)

finaldir = outputdir + "xl/embeddings/"

# Moves file to output folder

extractedfiles = os.listdir(finaldir)

for f in extractedfiles:
    os.rename(finaldir + f, outputdir + f)

# Clean up

shutil.rmtree(outputdir + "xl", ignore_errors=True)
os.remove(outputdir + "zip.zip")

print("Files have have extracted")
