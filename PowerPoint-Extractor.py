#!/usr/bin/env python3

import argparse
import shutil
import random
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


# Get Current Directory and Output Directory

currentdir = os.getcwd() + os.sep
outputdir = os.getcwd() + os.sep + args.output + os.sep

# Make Output directory

doesoutputdirexist = False

if not os.path.exists(args.output):
    os.mkdir(args.output)

outputfile = rf"{outputdir}{args.file.name}"


# Copy User file to output directory

if args.file:
    shutil.copyfile(args.file.name, outputfile)


# Rename file to zip.zip

newzip = f"{currentdir}{random.randint(0,100000)}.zip"
os.rename(outputfile, newzip)

# Extracts files but keep the full directory

with ZipFile(newzip, "r") as zipObj:
    listOfFileNames = zipObj.namelist()
    for fileName in listOfFileNames:
        if fileName.endswith(".pptx"):
            zipObj.extract(fileName, outputdir)

finaldir = outputdir + f"xl{os.sep}embeddings{os.sep}"

# Moves file to output folder

extractedfiles = os.listdir(finaldir)

for f in extractedfiles:
    try:
        os.rename(finaldir + f, outputdir + f)
    except:
        print(
            f"There is already a file with the name '{f}', You have either already ran this script or need to change the name of the file\n"
        )

# Clean up

outputdir = os.getcwd() + os.sep + args.output + os.sep
shutil.rmtree(outputdir + "xl", ignore_errors=True)
os.remove(newzip)

print("Files have have extracted!\n")
