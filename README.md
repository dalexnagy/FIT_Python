# FIT_Python

FITProc - FIT File Processing Program

BASIC PROCESSING:

The FITProc Python program opens a FIT file uploaded from a bicycle computer, extracts data from it, and:
  A.	Saves some of this data into a MySQL database
  B.	Updates a spreadsheet containing data for the year by adding a line for the ride represented by the FIT data.
  C.	Sends a text message to a specific phone number (or numbers – it can be a list) with basic ride information and as a notification the rider finished the ride and is presumably home, 
  D.	If the program detects the reported battery level of the bike computer or electronic shifting system (DI2) is below a threshold, sends an email message to a user (or list of users) as a reminder to recharge the device or devices.

This program was written for the FIT file created by a Wahoo bicycle computer (Elemnt, Roam, Bolt).  The phone app can be set to upload this data file to services such as Strava, RidewithGPS, and others, and to DropBox.

When the file is uploaded to DropBox, this service will send a copy to connected computers. 

My primary desktop runs Ubuntu Linux and uses the ‘incron’ daemon to detect this file has been created and it starts the FITProc program automatically.  

The control line in ‘incrontab’ to automatically start the program is:
/home/dave/Dropbox/apps/WahooFitness/   IN_MOVED_TO     /home/dave/Scripts/FITProc.sh

The Bash shell script in the incrontab control will run the Python program and move the FIT file to an archive.

I cannot offer any suggestions to run this program and process a FIT file on a Windows or MacOS system.

REQUIREMENTS:

This program requires the following Python system libraries (besides those used for browsing folders and handling dates):
  1.	‘fitdecode’ to open and parse the FIT file structure.
  2.	‘openpyxl’ to open, read, and update an Excel spreadsheet (XLSX) file
  3.	‘pymysql’ to open, read, and update a MySQL database
  4.	‘smtplib’ and ‘ssl’ to send an email message
  5.	‘requests’ to connect with a service to send a text message
     
This program also uses three ‘private’ libraries to load variables used to login to MySQL, send an email message through GMail, and authorize a connection with a text messaging service provider.  These libraries are simple modules with variable assignments and are in a special folder.  Skeletons of these are provided in the GitHub library (after updating, save to ‘_Configs’ or your chosen folder, with an appropriate name) and should be adequately commented to provide guidance how to update these for your needs.
