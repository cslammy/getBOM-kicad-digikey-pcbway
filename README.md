# getBOM-kicad-digikey-pcbway
Python script taht creates PCBWAY assembly BOMs from Kicad9 BOM exports using Digikey API

....This is the python script I use to migrate Kicad 9 BOM’s to PCBWAY assembly BOM’s.
Instead of looking up every part by hand it tries to use the Digikey API to find matching parts.

To make it work.

--Make sure your python3 is working.

--Create new directory and put getbom.py into it.

--using PIP or uv, Add these python modules  

pandas 
openpyxl 
requests 
beautifulsoup4

--In Kicad 9 schematic editor tools > export BOM, save KICAD BOM as CSV file, put the CSV file into same folder as script. call it kicadbom.csv

--Put the Sample BOM PCBWAY file in the same folder as the script.

--Get a (free) Digikey API user and key:  
https://developer.digikey.com
Sign in there with your regular DigiKey account (the one you use to order parts), then go to My Apps → Create App, 
create a company, 
create an app called getbom, 
For Oauth callback use http://localhost
put getbom app into production
Edit getbom app

get your Client ID and Client Secret, visible from EDIT screen, and record them somewhere.

--on the PC used to run the script--Set environment variables (windows: “set”,  Mac/Linux: “export”) for the keys or (what I do) use Windows UI “system variables"

--run it:  from windows cmd or linux terminal:

python getbom.py kicadbom.csv

HOW IT WORKS:
It tries to match anything in Kicad BOM to Digikey part list.  
For most hardware (Switches, knobs, etc) lookups it will ask for a URL (I added this override because Digikey almost never gets the exact part right for hardware lookups)  
It will also give you a list of what it couldn’t find at the end.
--Creates an XLSX file that conforms to the PCBWAY BOM template when it’s done.
--Script will try to scrape the digikey site if it can, but I have found in most cases Digikey will block scrapes so you have to get the API working.

Final steps:
Edit the output XLSX file.
Remove things that obviously have no matching parts, and will not be used in assembly, like wirepads, panel mounted pots, etc., from the final list.

How well does it work?
I’ve used this for the last 3 projects and—works OK.

It may not find some parts—you have to manually fix that, but it’s still much faster than doing the whole thing by hand. 

Hope someone other than me finds this useful.

