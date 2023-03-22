#!/usr/bin/env python3

import argparse
import shutil
import random
import os
from zipfile import ZipFile


def parse_args():
    parser = argparse.ArgumentParser(
        add_help=False,
        description="Simple Python script to extract embedded powerpoints within Excel documents",
    )

    parser.add_argument(
        "--file",
        "-f",
        type=str,
        help="Select file for extracting",
        required=True,
    )

    parser.add_argument(
        "--output",
        "-o",
        type=str,
        help='The output directory for the extracted documents, By default the documents will extract to "Output"',
        default="Output",
    )

    parser.add_argument("--help", "-h", action="store_true")

    args = parser.parse_args()

    if not os.path.isabs(args.file):
        # If the file path is relative, make it absolute using the current working directory
        args.file = os.path.join(os.getcwd(), args.file)

    return args


def create_output_directory(output_dir):
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)


def copy_file_to_output_directory(input_file, output_dir):
    input_file_name = input_file
    output_file_name = os.path.join(output_dir, os.path.basename(input_file_name))
    shutil.copyfile(input_file_name, output_file_name)
    return output_file_name


def rename_file_with_random_integer(file_path):
    new_name = os.path.join(os.getcwd(), str(random.randint(0, 100000)) + ".zip")
    os.rename(file_path, new_name)
    return new_name


def extract_powerpoint_files_from_zip(input_zip, output_dir):
    with ZipFile(input_zip, "r") as zipObj:
        for fileName in zipObj.namelist():
            if fileName.endswith(".pptx"):
                zipObj.extract(fileName, output_dir)


def rename_extracted_files(output_dir, final_dir):
    extracted_files = os.listdir(final_dir)
    for file_name in extracted_files:
        try:
            os.rename(
                os.path.join(final_dir, file_name), os.path.join(output_dir, file_name)
            )
        except:
            print(
                f"There is already a file with the name '{file_name}', You have either already ran this script or need to change the name of the file\n"
            )


def delete_zip_and_xl_directory(zip_file_path, output_dir):
    shutil.rmtree(os.path.join(output_dir, "xl"), ignore_errors=True)
    os.remove(zip_file_path)


def main():
    args = parse_args()

    if args.help:
        print(
            "\nUsage: Powerpoint-extractor.py [-h] [--file FILE] [--output OUTPUT]\n\nSimple Python script to extract embedded powerpoints within Excel documents\n\nOptional Arguments:\n --file, -f Select file for extracting.\n --output, -o The directory where the extracted documents will be placed, By default the documents will extract to 'Output'.\n --help, -h Display this help message."
        )
        return

    output_dir = os.path.join(os.getcwd(), args.output)
    create_output_directory(output_dir)

    output_file_path = copy_file_to_output_directory(args.file, output_dir)

    zip_file_path = rename_file_with_random_integer(output_file_path)

    extract_powerpoint_files_from_zip(zip_file_path, output_dir)

    final_dir = os.path.join(output_dir, "xl", "embeddings")
    rename_extracted_files(output_dir, final_dir)

    delete_zip_and_xl_directory(zip_file_path, output_dir)

    print("Files have been extracted!\n")


if __name__ == "__main__":
    main()
