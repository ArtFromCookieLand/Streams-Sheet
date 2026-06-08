 # TStreams

Project for storing and managing Taylor Swift's streams on Spotify in a google spreadsheet.

## Description

In order to keep track of Taylor Swift's streams on Spotify, this project was created. Google Spreadsheets is used to store the data, Apify App "Spotify Play Count Scraper" is used to scrape the data, and Google Apps Script is used to automate the process of updating the spreadsheet with the latest stream counts. Clasp is used to connect Google Apps Script with a coding environment such as VSCode. 

## Usage

### Pushing and pulling

In order to push and pull the code from the Google Apps Script, you can use the following commands:
```
clasp push
clasp pull
```

## Authors

Haunted_Spotify (also known as Haunted_Jade or Jade) [Jade is not the actual name]

## Version History

* 1.1 (13/5/26)
    * Added generation of summaries for albums and the discrography.
* 1.0
    * The first tracked version of the project from May 2026 (initially the project was created in August 2025). Includes such features as:
        * Saving the data in a Google Spreadsheet and managing it
        * Scraping Taylor Swift's stream counts on Spotify from Apify (added late 2025)
        * Tracking milestones surpassed and to be surpassed 
        * Automating the process of updating the spreadsheet with the latest stream counts using Google Apps Script.
        * Running a webpage with daily streams update, which pulls the data from the spreadsheet and displays it in a visually appealing way (added early 2026 and deprecated shortly after due to lack of time and release of alternatives)

## Links

* [The Spreadsheet](https://docs.google.com/spreadsheets/d/1ANVfUER8MInJ_y-oRp7Nguwk0yDI9LtBt36PWBMmZbU)
* [The Apify App](https://console.apify.com/actors/YZhD6hYc8daYSWXKs/input?addFromActorId=YZhD6hYc8daYSWXKs)
* [Twitter Page](https://x.com/Haunted_Spotify)