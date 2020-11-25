# Strava2GoogleSheet
Analyse your Strava Data in Google Sheets. 

The code uses google script to authenticate to the Strava API and imports your activities to a sheet. From here you can create your dashboards.

## Authenticate to Strava and its API

Everything you need to know is described here: https://www.benlcollins.com/spreadsheets/strava-api-with-google-sheets/ || https://github.com/benlcollins/strava-sheets-integration
Thanks to Ben Collins. All the kudos belongs to him alone!

## Google Spreadsheets Stuff

My dashboards looks like this: https://github.com/rotti/Strava2GoogleSheet/blob/main/dashboard.png

I dont provide the pivot tables and actual dashboard elements here. You will find many guides and tutorials on the web. If you want to have a look how the result looks for me, you can have a look here: https://docs.google.com/spreadsheets/d/1ucdShZGE7XIRqKcaoEUEqMHPvVrvlvjwycSRedxBTfQ/edit?usp=sharing 
Note the one of the dashboards elements (the gauge one) needs edit rights to work. Unlikely I will provide those rights to you :)

I provide a sample CSV which can be used for the initial data-table. The last column "translates" the ISO Date from the activity to a format, which can be used for a time series dashboard. Just add the following formular to this column inside the sheet: =ARRAYFORMULA(IF(LEN(A2:A),--REGEXREPLACE(REGEXREPLACE(E2:E,"T"," "),"[^0-9-:\s]","")))

My general tip is that you create pivot tables for all your needs. For example I am using stuff like:
* daily, monthly, year statistic for all types of activities correlated to distance, elevation, time spend, number of activities
* activity (all types) correlated to distance, elevation, time spend, number of activities
* activity type "ride" correlated to max distance, elevation, time spend



