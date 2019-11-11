# booking-form

### Introduction

This is a script Guillaume created to book the fablab resources and update a resource calendar, send notifications etc... all based on Google tools. Some bits of scripts have been copied here and there so if you know of the owner, please let us know so we can attribute them correctly.

*Here's a screenshot of the form :*

![form overview](https://github.com/Pangloss-Labs/booking-form/blob/master/bookingform1.PNG)

*Here's a screenshot of the resource details*

![resource details](https://github.com/Pangloss-Labs/booking-form/blob/master/bookingform2.PNG)

### How does it work ?

- A form (scripts at the root of the repo) is created.
- Users fill it in to book a room, machine or event.
- It checks for the resource availability and if everything is available (checking against google sheet + calendars), it updates a google sheet with a new row for a booking. Otherwise the user gets notified to book at a different time
- It sends an email to the user to confirm the booking.
- If it's an event, the conciergerie needs to approve it so they receive an email to update the google sheet
- Macros run on the google sheet to check for items, and if they are approved and not published, it publishes them on the corresponding calendars (scripts in the event manager folder)

### How do I install it ?

You'll need to create a google webapp and copy the scripts in the right files.
Copy this google sheet that will capture all the events created by the form : https://docs.google.com/spreadsheets/d/1sbn8Jqxz-Fk-rxmNp1zpur3_IVpnJyeJXy2R_dpzL-c/edit?usp=sharing
Populate the parameter.gs file with your parameters.
Look for the word "sample" in the files and change them to the necessary value : calendar ID, email, etc...
Copy the macros,etc... from the repo subfolders to the google sheet
Publish the whole thing
And it should work :-)

### Disclaimer

Guillaume and I (yannick) are not developpers so the code might be a bit messy.
A lot can be improved, I'd almost say it's not in a ready to share state yet but since we're all limited on time, we might as well start for here :-)
I'm not even sure if this is the right way to share it/ deploy, any input about that is welcome.
Please look at the issues if you want to help improve it so it's easier to deploy/share , your help is appreciated !


Created by Guillaume Cabrié (http://lemantek.com/fr ), updated by Yannick Laignel for the Pangloss Labs association (https://panglosslabs.org/ )
