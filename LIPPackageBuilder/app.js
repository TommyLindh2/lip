lbs.apploader.register('LIPPackageBuilder', function () {
    var self = this;

    /*Config (version 2.0)
        This is the setup of your app. Specify which data and resources that should loaded to set the enviroment of your app.
        App specific setup for your app to the config section here, i.e self.config.yourPropertiy:'foo'
        The variabels specified in "config:{}", when you initalize your app are available in in the object "appConfig".
    */
    self.config =  function(appConfig){
            this.yourPropertyDefinedWhenTheAppIsUsed = appConfig.yourProperty;
            this.dataSources = [];
            this.resources = {
                scripts: ['model.js'], // <= External libs for your apps. Must be a file
                styles: ['app.css'], // <= Load styling for the app.
                libs: ['json2xml.js'] // <= Allready included libs, put not loaded per default. Example json2xml.js
            };
    };

    //initialize
    /*Initialize
        Initialize happens after the data and recources are loaded but before the view is rendered.
        Here it is your job to implement the logic of your app, by attaching data and functions to 'viewModel' and then returning it
        The data you requested along with localization are delivered in the variable viewModel.
        You may make any modifications you please to it or replace is with a entirely new one before returning it.
        The returned viewModel will be used to build your app.
        
        Node is a reference to the HTML-node where the app is being initalized form. Frankly we do not know when you'll ever need it,
        but, well, here you have it.
    */
    self.initialize = function (node, vm) {

        $('title').html('LIP Package builder');


        vm.fieldTypes = {
            "1" : "string",
            "2" : "geography",
            "3" : "integer",
            "4" : "decimal",
            "7" : "time",
            "8" : "text",
            "9" : "script",
            "10" : "html",
            "11" : "xml",
            "12" : "link",
            "13" : "yesno",
            "14" : "multirelation",
            "15" : "file",
            "16" : "relation",
            "17" : "user",
            "18" : "security",
            "19" : "calendar",
            "20" : "set",
            "21" : "option",
            "22" : "image",
            "23" : "formatedstring",
            "25" : "automatic",
            "26" : "color",
            "27" : "sql",
            "255" : "system"
        };
        
        // Attributes for tables
        vm.tableAttributes = [
            "tableorder",
            "invisible",
            "descriptive",
            "syscomment",
            "label",
            "log",
            "actionpad"
        ];

        // Attributes for fields
        vm.fieldAttributes = [
            "fieldtype",
            "limereadonly",
            "invisible",
            "required",
            "width",
            "height",
            "length",
            "defaultvalue",
            "limedefaultvalue",
            "limerequiredforedit",
            "newline",
            "sql",
            "onsqlupdate",
            "onsqlinsert",
            "fieldorder",
            "isnullable",
            "type",
            "relationtab",
            "syscomment",
            "formatsql",
            "limevalidationrule",
            "label",
            "adlabel",
            "idrelation"
        ];
        
        // Checkbox to select all tables
        vm.selectTables = ko.observable(false);
        
        vm.selectTables.subscribe(function(newValue){
            ko.utils.arrayForEach(vm.filteredTables(),function(item){
                item.selected(newValue);
            });
        });
        vm.getVbaComponents = function(){
            try{
                var components = lbs.common.executeVba('LIPPackageBuilder.GetVBAComponents');
                components = $.parseJSON(components);
            
                vm.vbaComponents(ko.utils.arrayMap(components,function(c){
                    if (c.type !== '100'){
                        return new VbaComponent(c);
                    }
                }));
            
            vm.vbaComponents.sort(function(left,right){
                return left.type == right.type ? 0 : (left.type < right.type ? -1 : 1);
            });
            }catch(e){alert(e);}
            vm.componentFilter("");
            vm.filteredComponents(vm.vbaComponents());
            vm.showComponents(true);
        }
        

        // Navbar function to change tab
        vm.showTab = function(t){
            try{
                if (t == 'vba'){
                    vm.getVbaComponents();
                }
                vm.activeTab(t);
                
            }
            catch(e){alert(e);}
        }
        
        // Set default tab to details
        vm.activeTab = ko.observable("details");
        
        // Array with VBA components
        vm.vbaComponents = ko.observableArray();
        vm.showComponents = ko.observable(false);
        
        //Relation container
        vm.relations = ko.observableArray();
        
        // Serialize selected tables and fields and combine with localization data
        vm.serializePackage = function(){
            var data = {};
            var packageTables = [];
            var tables = [];
            var packageRelations = [];
            var relations = {};
        
            if (vm.name() == ""){
                alert("Package name is required");
                return;
            }
            try{
                // For each selected table
                $.each(vm.selectedTables(),function(i,table){
                    var packageTable = {};
                    // Fetch local names from table with same name
                    var localNameTable  = vm.localNames.Tables.filter(function(t){
                        return t.name == table.name;
                    })[0];

                    // Set singular and plural local names for table
                    packageTable.localname_singular = localNameTable.localname_singular;
                    packageTable.localname_plural = localNameTable.localname_plural;
                    
                    // For each selected field in current table
                    var fields = [];
                    var packageFields = [];
                    $.each(table.selectedFields(),function(j,field){
                        // Fetch local names from field with same name
                        var localNameField = localNameTable.Fields.filter(function(f){
                            return f.name == field.name;
                        })[0];
                        //Clone the field
                        var packageField = jQuery.extend(true,{},field);
                        // Set local names for current field
                        packageField.localname = localNameField;
                        
                        //create relations
                        try{
                            if(field.attributes.fieldtype == "relation"){
                                //Lookup if relation already added
                                var existingRelation = relations[field.attributes.idrelation];
                                
                                if(existingRelation == null || existingRelation == undefined){
                                    var packageRelation = new Relation(field.attributes.idrelation,table.name, field.name);
                                    relations[field.attributes.idrelation] = packageRelation;
                                    
                                    
                                }
                                else{
                                    existingRelation.table2 = table.name;
                                    existingRelation.field2 = field.name;
                                }
                            }
                        }
                        catch(e){
                            alert(e);
                        }
                        
                        if(packageField.localname && packageField.localname.name){
                            delete packageField.localname.name;
                        }

                        if(packageField.localname && packageField.localname.order){
                            delete packageField.localname.order;
                        }
                        
                        //The separator is added correctly as a property on a field, instead of localname.
                        if(packageField.localname && packageField.localname.separator){
                            packageField.separator = packageField.localname.separator;
                            
                            delete packageField.localname.separator;
                                
                        }
                        
                        if(packageField.separator && packageField.separator.order)
                            delete packageField.separator.order;   

                        if(packageField.localname && packageField.localname.option)
                            delete packageField.localname.option;

                        // Push field to fields
                        fields.push(packageField);
                        
                        
                    });
                    
                    
                    
                    //Add relations as the package expects
                    
                    for(idrelation in relations){
                        if(relations[idrelation].table2 != ""){
                            packageRelations.push({"table1": relations[idrelation].table1,
                                                    "field1": relations[idrelation].field1,
                                                    "table2": relations[idrelation].table2,
                                                    "field2": relations[idrelation].field2
                                                    })
                        }
                        
                    }
                    
                    // Set fields to the current table
                    table.fields = fields;
                    
                    // Push table to tables
                    packageTables.push(table);
                });
                
                var packageRelationFields = [];
                //Fetch all relationfields in package
                var index;
                for(index = 0;index < packageTables.length; ++index){
                    var j;
                    for (j = 0;j <  packageTables[index].fields.length; j++){
                      var f = packageTables[index].fields[j];
                      if (f.attributes.fieldtype == "relation"){
                        packageRelationFields.push({ "name":packageTables[index].name + '.' + f.name, "remove": 1});   
                      }
                    }
                }
                
                //Check if field is existing in an relation (ugliest code)
                for (index = 0;index < packageRelationFields.length; index++){
                    var rf = packageRelationFields[index];
                    var j;
                    for (j = 0; j < packageRelations.length;j++){
                        var rel = packageRelations[j];
                        if (rel.table1 + '.' + rel.field1 == rf.name || rel.table2 + '.' + rel.field2 == rf.name){
                            rf.remove = 0;
                        }
                    }
                }
                
                //remove unpaired relationfields (This code might be cleaner...)
                $.each(packageRelationFields,function(i,relField){
                    if(relField.remove == 1){
                        $.each(packageTables, function(j,packageTable){
                            if(packageTable.name == relField.name.substring(0, relField.name.indexOf("."))){
                                var indexOfObjectToRemove;
                                //find the field to remove
                                $.each(packageTable.fields, function(k, packageField){
                                    if (packageField.name == relField.name.substring(relField.name.indexOf(".") + 1)){
                                        indexOfObjectToRemove = k;
                                    }
                                });
                                //remove field from package
                                if(indexOfObjectToRemove){
                                    alert(indexOfObjectToRemove);
                                    packageTable.fields.splice(indexOfObjectToRemove,1);
                                }
                            }
                        
                        
                    });
                }
                });
                
                
                
            }
            catch(e){
                alert(e);
            }
            
            try {
                arrComponents = [];
                $.each(vm.selectedVbaComponents(), function(i, component){
                    arrComponents.push({"name": component.name, "relPath": "Install\\" + component.name + component.extension() })
                });
                
                // Build package json from details and database structure
                data = {
                    "name": vm.name(),
                    "author": vm.author(),
                    "status": vm.status(),
                    "shortDesc": vm.description(),
                    "versions":[
                        {
                        "version": vm.versionNumber(),
                        "date": moment().format("YYYY-MM-DD"),
                        "comments": vm.comment()
                    }],
                    "install" : {
                        "tables" : packageTables,
                        "vba" : arrComponents
                        //"relations": packageRelations
                    }
                }
                //lbs.log.debug(JSON.stringify(data));
            }catch(e) {alert("Error serializing LIP Package:\n\n" + e);}
            
            // Save using VBA Method
            try{
                //Base64 encode the entire string, commas don't do well in VBA calls.
                lbs.common.executeVba('LIPPackageBuilder.CreatePackage', window.btoa(JSON.stringify(data)));
            }catch(e){alert(e);}
            
            // Save to file using microsofts weird ass self developed file saving stuff
            //var blobObject = new Blob([JSON.stringify(data)]); 
            //window.navigator.msSaveBlob(blobObject, 'package.json')
        }
            
        
        vm.filterComponents = function(){
            if(vm.componentFilter() != ""){
                vm.filteredComponents.removeAll(); 

                // Filter on the three visible columns (name, localname, timestamp)
                vm.filteredComponents(ko.utils.arrayFilter(vm.vbaComponents(), function(item) {
                    if(item.name.toLowerCase().indexOf(vm.componentFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.type.toLowerCase().indexOf(vm.componentFilter().toLowerCase()) != -1){
                        return true;
                    }
                    return false;
                }));
            }else{  
                vm.filteredComponents(vm.vbaComponents().slice());
            }
        }
    
        
    
        // Function to filter tables
        vm.filterTables = function(){
            if(vm.tableFilter() != ""){
                vm.filteredTables.removeAll(); 

                // Filter on the three visible columns (name, localname, timestamp)
                vm.filteredTables(ko.utils.arrayFilter(vm.tables(), function(item) {
                    if(item.name.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.localname.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.timestamp().toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    return false;
                }));
            }else{  
                vm.filteredTables(vm.tables().slice());
            }
        }

        // Filter observables
        vm.tableFilter = ko.observable("");
        vm.fieldFilter = ko.observable("");
        vm.componentFilter = ko.observable("");
        
        function b64_to_utf8(str) {
            return unescape(window.atob(str));
        }
    
        
        
        // Load databas structure
        try{
            var db = {};
            //lbs.loader.loadDataSource(db, { type: 'storedProcedure', source: 'csp_lip_getxmldatabase_wrapper', alias: 'structure' }, false);
            db = window.external.run('LIPPackageBuilder.LoadDataStructure', 'csp_lip_getxmldatabase_wrapper');
            db = db.replace(/\r?\n|\r/g,"");
            db = b64_to_utf8(db);
            
            var json = xml2json($.parseXML(db),''); 
            
            json = $.parseJSON(json);
            
            vm.datastructure = json.data;
        }
        catch(err){
            alert(err)
        }
        // Data from details
        vm.author = ko.observable("");
        vm.comment = ko.observable("");
        vm.description = ko.observable("");
        vm.versionNumber = ko.observable("");
        vm.name = ko.observable("");
        // Set default status to development
        vm.status = ko.observable("Development");

        // Set status options 
        vm.statusOptions = ko.observableArray([
            new StatusOption('Development'), new StatusOption('Beta'), new StatusOption('Release')
        ]);
        
        // Load localization data
        try{
            /*var localData = {};
            lbs.loader.loadDataSource(localData, { type: 'storedProcedure', source: 'csp_lip_getlocalnames', alias: 'localNames' }, false);
            vm.localNames = localData.localNames.data;*/
            var localData = "";
            localData = lbs.common.executeVba('LIPPackageBuilder.LoadDataStructure, csp_lip_getlocalnames');
            localData = localData.replace(/\r?\n|\r/g,"");
            localData = b64_to_utf8(localData);
            
            var json = xml2json($.parseXML(localData),''); 
            json = json.replace(/\\/g,"\\\\");
            
            json = $.parseJSON(json);
            vm.localNames = json.data;
        }
        catch(err){
            alert(err)
        }
        // Table for which fields are shown
        vm.shownTable = ko.observable();
        // All tables loaded
        vm.tables = ko.observableArray();
        // Filtered tables. These are the ones loaded into the view
        vm.filteredTables = ko.observableArray();
        
        // Filtered Components
        vm.filteredComponents = ko.observableArray();
        
        // Load model objects
        initModel(vm);

        // Populate table objects
        vm.tables(ko.utils.arrayMap(vm.datastructure.table,function(t){
            return new Table(t);
        }));
        
        // Computed with all selected vba components
        vm.selectedVbaComponents = ko.computed(function(){
            if(vm.vbaComponents()){
                return ko.utils.arrayFilter(vm.vbaComponents(),function(c){
                    return c.selected() | false;
                });
            }
        });
        
        // Computed with all selected tables
        vm.selectedTables = ko.computed(function(){
            return ko.utils.arrayFilter(vm.tables(), function(t){
                return t.selected();
            });
        });

        // Subscribe to changes in filters
        vm.fieldFilter.subscribe(function(newValue){ 
            vm.shownTable().filterFields();
        })
        vm.tableFilter.subscribe(function(newValue){
            vm.filterTables();
        });
        
        vm.componentFilter.subscribe(function(newValue){
            vm.filterComponents();
        });
        
        vm.filterComponents();
        // Set default filter
        vm.filterTables();

        return vm;
    };


    
});


ko.bindingHandlers.stopBubble = {
  init: function(element) {
    ko.utils.registerEventHandler(element, "click", function(event) {
         event.cancelBubble = true;
         if (event.stopPropagation) {
            event.stopPropagation(); 
         }
    });
  }
};