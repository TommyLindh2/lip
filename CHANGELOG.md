# Changelog for LIP

## v1.2.0
**Released:** 2017-08-04

**Authors:** Lars Jensen, Jonny Springare, Rasmus Alestig Thunborg, Lundalogik AB

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

**Authors:** Lars Jensen, Jonny Springare, Lundalogik AB

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

**Authors:** Lars Jensen, Jonny Springare, Pawel Demczuk, Fredrik Eriksson, Filip Arenbo, Lundalogik AB

**Comments:** This is the first official release of app and package installer for Lime CRM.
