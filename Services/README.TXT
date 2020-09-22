STEP 1: Open the project in the folder called 00_StartHere
	This project contains the source and compiled OCX.
	Run the project within the IDE and click the button
	to install the package onto x:\cwService where x:
	is your boot drive.

STEP 2: Open the next project 01_CPUTEST and COMPILE the 
	project into an EXE.  Use the INSTALL and UNINSTALL
	batch scripts to install or uninstall the service.
	DO NOT IN/UNINSTALL WHILE IN THE VB IDE.
	Once the service is installed, use the service
	manager to start/pause/stop the service.  This
	project is an INTERACTIVE service; thus it is a
	visible service seen on your desktop.

STEP 3: In the same fashion as you did in STEP 2, open and
	compile the project 02_WebService.  This project is
	a NON-INTERACTIVE web server that runs completely
	on its own weather your machine is logged into or
	not.  (Like Anti-Virus applications are run.)
	The server can be tested by pointing your web browser
	to http://localhost - if for some reason you already
	have a web server running on your machine, the demo
	will automatically try to use port 8080 instead.  If
	so, point your browser to http://localhost:8080 to 
	see the demonstration.

BOTH DEMONSTRATION PROJECTS ARE OTHER PEOPLE'S PROJECTS FROM
PLANET-SOURCE-CODE.  ALL CREDITS FOR THE FUNCTIONALITY GO TO
THE ASSOCIATED AUTHORS.  We chose to use EXISTING projects
from PSC simply to show you how you can take an existing app
and convert it into a Windows Background Service.

NOTE: Some changes had to be made to the 3rd party software
in order to make it usable as a service.  In many cases, you
will need for first add the cwSvc.ocx reference to the project
and set up the software to where it can be controlled by the
service instead of the loading form.  It may also be wise to
create a separate shut-down routine so that your app does not
actually quit when it's told to PAUSE.

CPU PROJECT credited to:
http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=27505&lngWId=1

WEB PROJECT Credited to:
http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=49779&lngWId=1

No liability is offered for the functionality of either the 3rd
party resources nor the enclosed control project.  Raw sources
for the control are offered under the General Public License and
may be used for non-commercial projects only.  Commercial and/or
"for-sale" projects require written permission where you will be
asked for your name, intended application and minimal contact 
information to assist in preventing malicious use of our products.