I will try to create a change-log file.
The idea is to post here more detailed description of what has been done in current(last) update.

Author: Mike

**Date** 2019-09-28.

**Commit note:** Fixed path to movie files, unexpectedly I created a PDF viewer

**Detailed note:**

I started process of splitting all-in-one ```DEFAULT.ASP``` file into modules.


It will allow us to work on "your-own" module without interfacing with anybody else

All submodules will be stored in the /ASP/ folder
Currently I created following sub-modules

- fnFormatDosPath.asp 
- fnLW.asp 
- fnMovieBox.asp 
- fnPdfBox.asp 
- fnSubfolder.asp 
- fnUpperFolder.asp

Currently these modules are **NOT** fully independent they are sharing common objects and variables with ```DEFAULT.ASP``` file, but the next few steps I will make them totally independent, so it will be possible to re-use them and test them independently.

------------
