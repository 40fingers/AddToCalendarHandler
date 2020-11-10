# 40FINGERS AddToCalendarHandler for DNN

## What is it?
We created a handler you can use in several ways in your DNN website to create a button or link to allow users to add an item to their calendar in Outlook, iCloud, Google, or any other calendar accepting ICAL files.

## How to install it?


## How to use it?
You will just need to add a link to this handler, with the parameters you need. You could construct that link in your templateable module, like OpenContent.
The location of the handler would be:
/DesktopModules/40Fingers/AddToCalendarHandler/AddToCalendar.ashx

Currently, the following parameters are supported:
* type: mandatory. Accepted values: [google|outlook|icloud|ical]
* subject: optional (but strongly advised as this is the subject of the event)
* description: optional
* location: optional
* start: mandatory: Start date/time in format yyyyMMddTHHmmss
* end: optional: End date/time in format yyyyMMddTHHmmss
* duration: optional: number of minutes

Note: if both "end" and "duration" are missing, the event defaults to 60 minutes.
Note 2: currently types icloud, outlook and ical result in identical ical files.

This would result in the url looking something like this:
/DesktopModules/40Fingers/AddToCalendarHandler/AddToCalendar.ashx?type=ical&subject=testsubject&description=testdescription&start=20201201T090000&duration=120&location=online
