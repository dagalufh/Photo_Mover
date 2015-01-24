Photo_Mover
===========

Photo Mover for moving/copying photos from a directory structure to multi level based on date taken.
The script now offers to ways of usage, either via VBS file or via HTA file. The HTA file offers a graphical interface.

There are optional keywords (Year, Month, Day) that can be used to organize photos based on the date taken property.

### Current Version:
- V1.0.0.9-HTA
 - Added control to make sure that the source folder exists before continuing.
 - Fixed default path for the folder browser to allow any directory.

### Requirements:
This script requires you to run a Windows operating system.

### Installation:
Download the file from https://github.com/dagalufh/Photo_Mover/releases and place it anywhere you like and run it.

### .HTA file usage (Move/Copy photos):
When starting the HTA file you are presented with an interface for supplying the source and target folder.
You are also given the option to move the photos or just copy them to the target.

As with the VBS file, script will create the files needed.

### .VBS file Usage (Move photos only):
Upon starting the VBS file, you as a user are asked to specify the source where photos are located. It will search through all the subdirectories for photos.

After that, you are asked about the target directory. There are keywords available that is year, month and day. These will be replaced with the dates from the photo in question. (based on Date Taken property)

Then the script starts creating directories as needed and moves photos. A log file is created in the source directory showing what has been done.

### Disclamer:
As this moves photos and reorganizes them, use this at your own risk.
