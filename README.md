# ServiceNow-IEAutomation

This repository is a collection of PowerShell scripts I made to be able to automate the logging of tickets in Service-Now.

## Use
Store in a central location ideally using a DNS Name for the server its on so if the location changes you will not need to update multiple scripts or processes that call them

## Process of script
The Process of the scripts are fairly simple:
  - Open IE as a COMObject
  - Reteieve password for account to login to Service-Now
  - Browse to Service-Now Website
  - Retrieve login fields
  - Enter login information
  - Using Service Now Favourite bar click on the button required using HTML CLASS attribute
  - Fill out the information in the ticket
  - Click submit

## Pre Setup Instructions

### Securely reteieve password back to clear text
As we are entering credentials into the website we need a password to be retrieved in clear text. This can be done any way you want but I use the PowerShell credential encryption. This is setup as follows (for a service account that doesn't have login rights):
  1. Login to server
  2. Setup a scheduled task in the following way:
    - Account -> Service Account Script will run as
    - Action -> Powershell.exe -executionpolicy bypass -file <Path to script snippet below>
    - Script file
      #Enter the password
      $Password = "PASSWORD"
      #Enter Location of the output file
      $PasswordFile = "C:\Password.txt"
      $Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File $PasswordFile

### Setting up correct web tags
When first setting these scripts up the HTML Tags may need to be changed to work with your companies Serivce-Now environment as I assume they are not the same everywhere.

The main one that will definitely need to be changed is the FavouriteBar ClassIDs.
This tag is the elements tag and you can look for it by searching the scripts for GetElementsByClassName

