# Plex-Recorded-Video-Converter
The HDHomeRun Prime records .TS video files. This project finds those files and converts them to .MP4. My "Plex Recorded TV Shows" library that stores these .TS files is a different library than my "TV Shows" library that I share. 

This project uses Handbrake and the Handbrake CLI command line interface. Edit the VBS source and set the paths where the script will find your files.

I set this script to run as a scheduled task on a few computers around the house and keep the video conversion work away from the Plex server. Schedule the task when the computers are on but not used. Conversion can eat up all the CPU of the server. Handbrake does have an NVIDIA GPU option that makes quick work of video conversion. I haven't tried that yet with this setup.

Folder to search for recorded shows:
CONST SOURCE_MEDIA_FOLDER = "\\homepvr\d$\Plex Recorded TV\"

Log file:
CONST LOGFILE = "\\homepvr\d$\Plex Recorded TV\_executionlog.log"

Command line for converting. I have replaceable parameters here inside % %:
CONST CONVERTER_COMMAND_LINE = """c:\Program Files (x86)\Plex Video Converter\Handbrake\HandBrakeCLI.exe"" -i ""%INPUT_FILE%"" -t 1 --angle 1 -c 1-11 -o ""%OUTPUT_FILE%""  -f mp4 --width %OUTPUT_WIDTH% --height %OUTPUT_HEIGHT% --crop 0:0:6:4 --loose-anamorphic  --modulus 2 -e x264 -q 20 --vfr -a 1 -E av_aac -6 dpl2 -R Auto -B 160 -D 0 --gain 0 --audio-fallback ac3  --encoder-preset=veryfast  --encoder-level=""4.0""  --encoder-profile=main  --verbose=1"

Parameter strings from the command line:
CONST INPUT_FILE_STRING = "%INPUT_FILE%"
CONST OUTPUT_FILE_STRING = "%OUTPUT_FILE%"
CONST OUTPUT_WIDTH_STRING = "%OUTPUT_WIDTH%"
CONST OUTPUT_HEIGHT_STRING = "%OUTPUT_HEIGHT%"

TV Library folder to save the output:
CONST OUTPUT_FOLDER = "\\homepvr\data\video\television\"

If FFMPEG encounters an error move the error file to this folder to examine:
CONST OUTPUT_ERROR_FOLDER = "\\homepvr\d$\Plex Recorded TV Conversion Errors\"
CONST OUTPUT_FILE_TYPE = ".MP4"
CONST INPUT_FILE_TYPE = ".TS"

Set the desired width and height for the output video. Make sure it is supported by the encoder level, profile, and file type:
CONST OUTPUT_WIDTH = "1280"
CONST OUTPUT_HEIGHT = "720"
CONST LOGGING = True
CONST PARSE_SUBFOLDERS = True
CONST IGNORE_HIDDEN_SUBFOLDERS = True
CONST DEBUGGING=FALSE

Other notes:
If you don't want a recorded video to overwrite files in a folder you can make the files read-only, or in the destination folder drop a text file named "_Do not overwrite.txt". The video converter will honor and not overwrite. I use this with a ripped show DVD output folder so that I don't overwrite the DVD content with recorded content.

