# LIP - Package Management for Lime CRM

LIP is a package management tool for Lime CRM add-ons. A package can contain for example tables and fields, VBA modules and localizations. LIP can install add-ons from a zip file and in some cases also download and install packages from the App Store. The latter is under reconstruction and will work for all Community or Core add-ons in the future.

LIP is inspired by Pythons PIP and Nodes NPM.

## Using LIP
The current implementation is written in VBA and is used through the Immediate window of the VBA editor.

### Install a package 
To install a package, simply type your command in the Immediate window of the VBA. There are three different installation methods:

`lip.Install("Packagename")`
*Currently not working for all add-ons, we need a new version of the App Store first.*

Searches for the package on the stores configured in your `packages.json` file in the Actionpads folder.

`lip.InstallFromZip`
*Currently the recommended way of using LIP.*

Installs a deploy package or add-on from a zip file. It will open a file dialog where you can point out the zip file you wish to install.

`lip.InstallFromPackageFile`
*Currently not in use!*

All installed packages are kept tracked of inside the file `packages.json` in the Actionpads folder. If you transfer this file to a new Lime CRM database you can use this file to make LIP install all add-ons listed in the `packages.json` file by typing the command above.

### Upgrade a package
*Currently not working for all add-ons, we need a new version of the App Store first.*

If a package already exist and should be updated or reinstalled you must explicitly use the update command:

`lip.UpgradePackage("ExamplePackage")`


## Features

### Automatic installation
These are the things LIP can automatically install:

* VBA modules, class modules and forms
* Tables
* Fields
* Localizations (in the localize table)
* Actionpads (manual registration needed afterwards)

### Manual installation required
The following things can be included in a LIP package but requires manual installation after the automatic installation has finished. In which folder in the LIP package they should be stored is stated within parenthesis. More details on how to store them can be found in the Community Add-ons requirement list (soon on Lime Prime).

* Lime Bootstrap apps (apps)
* VBA code snippets to insert in existing VBA modules, class modules or forms (vba)
* SQL procedures and functions (sql)
* Table icons (lisa\icons)
* Option queries (lisa\optionqueries)
* Table descriptive expressions (lisa\descriptives)
* SQL expressions (lisa\sql_expressions)
* SQL for new (lisa\sql_for_new)
* SQL on update (lisa\sql_on_update)


### Unsupported
Below is a list of unsupported things. Some of them may be added in the future.

* Policies
* Groups
* Users
* Record access
* Virtual tables
* Filters
* Views
* Templates
* Reports
* Set lists for fields


## Installation of LIP
LIP is included in all Lime Core database versions from v6.0.0. If your application however does not include LIP you can always install it yourself by following the below steps.

1. Download the zip file for [the latest LIP release](https://github.com/Lundalogik/lip/releases/).
2. Add the SQL procedures to your database by running the SQL-scripts (Important! Make sure you run the scripts on your database and NOT the master-database by selecting the correct database in the upper left corner).
3. Restart the LDC.
4. Restart the Lime CRM client.
5. Import the `vba\lip.bas` file to your VBA.
6. Import the `vba\FormProgress.frm` to your VBA.
7. Run `lip.InstallLIP` in the Immediate window of the VBA editor. This will install all other necessary things.
8. Compile and save the VBA.
9. Publish actionpads.
10. *Recommended:* Also install the LIP Package Builder by following the steps described [here](https://github.com/Lundalogik/addon-lip-package-builder).


## Upgrade
*More information coming*


## Definitions

### A Package
A package is a zip file containing all required resources to install a package. Packages could be created by the LIP Package Builder or manually if you create a package.json file. If you create a manual file, make sure to save it with the encoding "Western (Windows 1252)", for example using the editor Sublime Text, to avoid trouble with localization records.

### A Package Store
*This is to be revised.*

A Package Store is any valid source which serves correct JSON-files and package.zip files. You can configure which stores to use in the packages.json-file in your Actionpads folder. A Package Store could either be file based or web based. A store has a fixed URL (example "http://limebootstrap.lundalogik.se/api/apps"). The URL has subdirectories for each app (example "./checklist"). If the source is a file-based a `/app.json` should automatically be append.

#### Specifying Own Package Stores
*This is to be revised.*

Open the `packages.json`-file in the Actionpad folder and add your own stores in the "stores"-node. Give the store a name and provide the address. Example: "AppStore":"http://limebootstrap.lundalogik.com/api/apps/"
When installing or upgrading a package, the stores will be searched from the top to the bottom, meaning your first store will be searched for the package first, then the second store and so on.


## Versioning
### Package Versioning
Packages should adhere to semantic versioning, example `1.0.0` or `MAJOR.MINOR.PATCH`. Please read [this](http://semver.org). 

Simplified:
`MAJOR`: Breaks backwards compatibility
`MINOR`: Adds new features but backward compatible
`PATCH`: Bugfixes


### Dependency versioning
*This is currently not in use in LIP but the concepts thought of are described below*

Stating dependency verisons should adhere to [NPMs versioning](https://github.com/npm/node-semver)

Minor and Patchs should always be upgraded to automatically if a dependency requires it.

Major versions can only be upgraded to if explicit Upgrade command is used

A `version range` is a set of `comparators` which specify versions
that satisfy the range.

A `comparator` is composed of an `operator` and a `version`.  The set
of primitive `operators` is:

* `<` Less than
* `<=` Less than or equal to
* `>` Greater than
* `>=` Greater than or equal to
* `=` Equal.  If no operator is specified, then equality is assumed,
  so this operator is optional, but MAY be included.

For example, the comparator `>=1.2.7` would match the versions
`1.2.7`, `1.2.8`, `2.5.3`, and `1.3.9`, but not the versions `1.2.6`
or `1.1.0`.


## Reference

### package.json

An example of what the file package.json could look like:

```JSON
{
    "uniqueName": "[A UNIQUE NAME OF PACKAGE ONLY CONTAINING a-z]",
    "dependencies": {
        "vba_json":"1.0",
        "addon-lime-core":"6.0.0"
    },
    "install": {
        "localize": [
            {
                "owner": "checklist",
                "context": "title",
                "sv": "Test",
                "en_us": "Test",
                "no": "Test",
                "fi": "Test"
            },
            {
                "owner": "checklist",
                "context": "description",
                "sv": "A short description",
                "en_us": "A short description",
                "no": "A short description",
                "fi": "A short description"
            }
        ],
        "vba": [
            {
                "relPath": "vba\\Test.bas",
                "name": "Test"
            }
        ],
        "tables": [
            {
                "name": "test",
                "localname_singular": 
                {
                    "sv": "Test",
                    "en_us": "Test"
                },
                "localname_plural": 
                {
                    "sv": "Test",
                    "en_us": "Test"
                },
                "attributes": {
                    "tableorder": "3",
                    "invisible": "2",
                    "descriptive":"[test].[title]",
                    "syscomment": "My comment",
                    "label": "15",
                    "log": "1",
                    "actionpad": "lbs.html"
                },
                "fields": [
                    {
                        "name": "title",
                        "localname": {
                            "sv": "Titel",
                            "en_us": "Title"
                        }, 
                        "attributes": {
                            "fieldtype": "text",
                            "limereadonly": "1",
                            "invisible": "0",
                            "required": "1",
                            "width": "3",
                            "height": "1",
                            "length": "256",
                            "defaultvalue": "Lund",
                            "limedefaultvalue": "Application.ActiveUser.Record.Id",
                            "limerequiredforedit": "0",
                            "newline": "2",
                            "fieldorder": "4",
                            "isnullable": "0",
                            "type": "1",
                            "relationtab": "1",
                            "syscomment": "My private comment",
                            "formatsql": "0",
                            "limevalidationrule": "My validation rule",
                            "label": "18",
                            "adlabel": "31"
                        },
                        "separator": {
                            "sv": "Testseparator",
                            "en_us": "Test separator"
                        },
                        "limevalidationtext": {
                        	"sv": "Min valideringstext",
                        	"en_us": "My validation text"
                        },
                        "comment": {
                        	"sv": "Min kommentar",
                        	"en_us": "My comment"
                        },
                        "description": {
                        	"sv": "Min beskrivning",
                        	"en_us": "My tooltip"
                        },
                        "options": [
                        	{
                        		"sv": "Alt 1",
                        		"en_us": "Alt 1",
                        		"key": "alt1",
                        		"color": "65535",
                        		"default": "true"
                        	},
                        	{
                        		"sv": "Alt 2",
                        		"en_us": "Alt 2",
                        		"key": "alt2",
                        		"color": "255"
                			}
        				]
                    }
                ]
            }
        ],
        "relations": [
            {
                "table1": "company",
                "field1": "person",
                "table2": "person",
                "field2": "company"
            },
            {
                "table1": "business",
                "field1": "responsible",
                "table2": "coworker",
                "field2": "business"
            }
        ]
    }
}
```

#### localize
Here you can specify records to be added to the localize table in Lime CRM. An example:

```JSON
"localize": [
    {
        "owner": "checklist",
        "context": "title",
        "sv": "Test",
        "en_us": "Test",
        "da": "Test",
        "no": "Test",
        "fi": "Test"
    },
    {
        "owner": "checklist",
        "context": "description",
        "sv": "A short description",
        "en_us": "A short description",
        "da": "A short description",
        "no": "A short description",
        "fi": "A short description"
    }
]
```
#### vba
Here you can specify VBA-modules (Forms and Class Modules are also supported) that should be installed. Please note that the VBA-file MUST be included in the zip file of your package under the subfolder `vba`. Please specify the relative path to the VBA file and the name of the VBA module. When adding forms, please include both form-files (.frm and .frx) and specify the .frm-file in "relPath". Example:

```JSON
"vba": [
    {
        "relPath": "vba\\MyForm.frm",
        "name": "MyForm"
    },
    {
        "relPath": "vba\\MyClassModule.cls",
        "name": "MyClassModule"
    }
]
```

#### tables

##### name (mandatory)
Database name of the table. Example:

```JSON
"name": "goaltable"
```
##### localname_singular (mandatory)
Localnames in singular. Each line in this node should represent one language. Valid languages are all languages Lime CRM supports. Example:

```JSON
"localname_singular": {
"sv": "Måltabell",
"en_us": "Goal table"
}
```

##### localname_plural (mandatory)
Localnames in plural. Each line in this node represent a language. Valid languages are all languages Lime CRM supports. Example:

```JSON
"localname_plural": {
"sv": "Måltabeller",
"en_us": "Goal tables"
}
```

##### attributes
Sets attributes for the table. Each line in this node represent an attribute. Example:

```JSON
"attributes": {
                    "tableorder": "3",
                    "invisible": "2",
                    "descriptive": "[test].[title]",
                    "syscomment": "My comment",
                    "label": "15",
                    "log": "1",
                    "actionpad": "lbs.html"
                }
```

Valid attributes:

Attribute | Mandatory |  Possible values | Value if not provided
-----|------|-----|-----
tableorder|No|Integer|Placed last
descriptive|No|text|Record ID
invisible|No|1/2 ("Yes"/"Yes, for everyone but administrators")|"No"
syscomment ("Comment")|No|text|<empty>
label|No|Integer|No label
log ("Log all changes")|No|0/1 (No/Yes)|Default
actionpad|No|text|<empty>

##### fields
    
###### name (mandatory)
The database name of the field. Example:

```JSON
"name": "customernbr"
```

###### localname (mandatory)
Localnames for the field. Each line in this node represent a language. Valid languages are all languages Lime CRM supports. Example:

```JSON
"localname": {
    "sv": "Kundnummer",
    "en_us": "Customer number"
}
```

###### separator
Adds a separator to the field. The separator is placed BEFORE the field. Specify the localnames for the separator inside this node. Example:

```JSON
"separator": {
    "sv": "Administrativ information",
    "en_us": "Administrative information"
}
```
###### limevalidationtext
Adds validation text to the field. Specify the localnames for the validation text inside this node.
Example:

```JSON
"limevalidationtext": {
    "sv": "Min valideringstext",
    "en_us": "My validation text"
}
```
###### comment
Adds a comment to the field. Specify the localnames for the comment inside this node.
Example:

```JSON
"comment": {
    "sv": "Min kommentar",
    "en_us": "My comment"
}
```
###### description
Adds tooltip to the field. Specify the localnames for the tooltip inside this node.
Example:

```JSON
"description": {
    "sv": "Min beskrivning",
    "en_us": "My tooltip"
}
```
###### options
Adds options to an option field, set field or textfield. Every option must be specified as an own node and localnames and other attributes are placed inside this node.

**Important!** Localize rows must be placed first in the node. If you place the color or default attribute first, these attributes will not be set. The integer representing the color in Lime CRM is derived by taking the RGB hex code for the desired color, reordering it as BGR, and then transform from hexadecimal to decimal.

Some color-examples:

Color | Integer
-----|-----
red|255
yellow|65535
blue|16711680
green|32768

Example:

```JSON
"options": [
	{
		"sv": "Alt 1",
		"en_us": "Alt 1",
		"key": "alt1",
		"color": "65535",
		"default": "true"
	},
	{
		"sv": "Alt 2",
		"en_us": "Alt 2",
		"key": "alt2",
		"color": "255"
	}
]
```

###### attributes
Sets attributes for the field. Each line in this node represent an attribute.

```JSON
"attributes": {
    "fieldtype": "text",
    "limereadonly": "1",
    "invisible": "0",
    "required": "1",
    "width": "3",
    "height": "1",
    "length": "256",
    "defaultvalue": "Lund",
    "limedefaultvalue": "Application.ActiveUser.Record.Id",
    "limerequiredforedit": "0",
    "newline": "2",
    "fieldorder": "4",
    "isnullable": "0",
    "type": "1",
    "relationtab": "1",
    "syscomment": "My private comment",
    "formatsql": "0",
    "limevalidationrule": "My validation rule",
    "label": "18",
    "adlabel": "31"
}
```

Valid attributes:

Attribute | Mandatory |  Possible values | Value if not provided
------- | -------- | ------- | --------
fieldtype|Yes|string/integer/decimal/time/html xml/link/yesno/file/relation/geography set/option/formatedstring/color/sql|-
invisible|No|0/1/2/65535 (No/On forms/In lists/Everywhere)|Default
length|No|integer (can only be set for textfields)|nvarchar(max)
required|No|0/1|0
fieldorder|No|Integer|Put last
height|No|Integer|Default
width|No|Integer|Default
newline (width properties)|No|0/1/2/3 ("Variable width"/ "Variable width on New line"/ "Fixed width"/ "Fixed width on new line")|2 (Fixed Width)
defaultvalue|No|text|Default
limedefaultvalue|No|text|Default
isnullable|No|0/1|0
limereadonly|No|0/1|Default
limerequiredforedit|No|0/1|Default
type|No|**Timefields:** 0/1/2/3/4/5/6/7/8/9 ("Date" / "Date and Time" / "Time" / "Year" / "Half a Year" / "Four Months" / "Quarter" / "Month" / "Week" / "Date and Time (with Seconds)" **Optionlists:** 0/1 ("Color and Text"/"Only Color")|0
relationtab|No|0/1 (actually corresponds to relationmaxcount, the name is misleading. Only valid when creating a relation)|0
syscomment (private comment)|No|text|<empty>
formatsql|No|0/1 (False/True)|Default
limevalidationrule|No|text|<empty>
label|No|Integer|No label
adlabel|No|Integer|No AD-label

#### relations
Here you specify which relations to create. This section only contains information about which fields/tabs to create a relation between, the rest of the information about each field you specify in the field-section. There you also specify whether the field should be an actual field or a tab (attribute relationtab, which actually corresponds to attribute relationmaxcount). Example:

```JSON
"relations": [
	{
		"table1": "company",
		"field1": "person",
		"table2": "person",
		"field2": "company"
	},
	{
		"table1": "business",
		"field1": "responsible",
		"table2": "coworker",
		"field2": "business"
	}
]
```
