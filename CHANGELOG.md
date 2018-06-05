# Changelog for LIP

## vNEXT
**Released:** 2018-xx-xx

**Authors:** Jonny Springare

**Comments:** 

* New version of the Progress bar VBA form that will solve issue in some environments.


## v1.3.0
**Released:** 2018-04-xx

**Authors:** Fredrik Eriksson, Lars Jensen, Jonny Springare

**Comments:** 

* Now supports the structure for add-ons.
* Possible to install Actionpads.
* Progress bar title now states "Simulating" when simulations are run.
* User is now notified if there are things that must be installed manually after a lip.InstallFromZip has been run (SQL or LISA things).
* Installing LIP no longer asks if you want to install the Package Builder (since this wasn't done fully anyway and LIP maybe wasn't ready to install other packages without a restart in some cases).
* Unexpected local date formats could result in an error with the log file now solved.
* Improved README.


## v1.2.0
**Released:** 2017-08-04

**Authors:** Lars Jensen, Jonny Springare, Rasmus Alestig Thunborg

**Comments:** 

#### Features

* After a package/app is installed an explorer window and the README.md file is opened.
* Helper method for LIP version added.

#### Bug fixes

* LIP would throw error when trying to create text fields (and decimal fields) with options.
* Naming difference would caused LIP to be installed again when installing packages.
* LIP assumed that there was a packages folder in the webfolder which caused error when installing an app or package for the first time.
* Localization posts weren't installed correctly.
* Tables and fields with special characters resulted in an error.


## v1.1.0
**Released:** 2017-05-03

**Authors:** Lars Jensen, Jonny Springare

**Comments:**

#### Features

* When you install LIP it asks if you want to install LIPPackageBuilder automatically.
* Added extended warnings when simulating installation.
* Helper method for LIP version added.

#### Bug fixes

* HTML tabs where installed as HTML fields.
* Intepretaion of decimal sometimes broke LIP installation.
* Mandatory option fields threw an error.
* Long default values on textfields broke installation of LIP packages.
* Some minor bugs regarding writing to log file fixed.


## v1.0.0
**Released:** 2017-01-18

**Authors:** Lars Jensen, Jonny Springare, Pawel Demczuk, Fredrik Eriksson, Filip Arenbo

**Comments:** This is the first official release of app and package installer for Lime CRM.
