# Strava2GoogleSheet
Analyse your Strava Data in Google Sheets. 

The code uses google script to authenticate to the Strava API and imports your activities to a sheet. From here you can create your dashboards.

## Authenticate to Strava and its API

Everything you need to know is described here: https://www.benlcollins.com/spreadsheets/strava-api-with-google-sheets/ || https://github.com/benlcollins/strava-sheets-integration
Thanks to Ben Collins. All the kudos belongs to him alone!

## Google Spreadsheets Stuff

My dashboards looks like this: https://github.com/rotti/Strava2GoogleSheet/blob/main/dashboard.png

I don't provide the pivot tables and actual dashboard elements. You will find many guides and tutorials on the web. You can even give money to Ben Collins and ask him. If you want to have a look how the result looks for me, you can have a look here: https://docs.google.com/spreadsheets/d/1ucdShZGE7XIRqKcaoEUEqMHPvVrvlvjwycSRedxBTfQ/edit?usp=sharing 
Note that one of the dashboards elements (the gauge one) needs edit rights to work. Unlikely I will provide those rights to you :)

I provide a sample CSV which can be used for the initial data-table. The last column "translates" the ISO Date from the activity to a format, which can be used for a time series dashboard, if wanted. Just add the following formular to this column inside the sheet: =ARRAYFORMULA(IF(LEN(A2:A),--REGEXREPLACE(REGEXREPLACE(E2:E,"T"," "),"[^0-9-:\s]","")))

My general tip is that you create pivot tables in seperated sheets for all your needs. Use "slicers", they are cool. For example I am using pivot for stuff like:
* daily, monthly, year statistic for all types of activities correlated to distance, elevation, time spend, number of activities
* activity (all types) correlated to distance, elevation, time spend, number of activities
* activity type "ride" correlated to max distance, max elevation, max time spend

For the rest just make charts from the sheet with the activity data. The gauge element uses "VLOOKUP" and gets its green/yellow/red ranges based on my personal "goals". 

You simple make charts in all your sheets and copy and paste them to your dashboard sheet. Than hide all your pivot sheets and the data sheet itself.

Just play around. Google made awesome stuff here.



