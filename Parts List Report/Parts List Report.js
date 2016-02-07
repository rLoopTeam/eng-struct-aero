//Author - Chris Smith
//Description-Creates a parts report based on active model. Can run in 3 different modes, 1:) all parts and assemblies are included, 
//2:) Just the first level parts and assemblies are listed, and 3:) Only the parts are listed (no assemblies).

/*globals adsk*/
var filename;
var parts = new Array(0);
var report_type;

function run(context) {
    "use strict";
    if (adsk.debug === true) {
		/*jslint debug: true*/
		debugger;
		/*jslint debug: false*/
	}
    
	
    try {
		var app = adsk.core.Application.get();
		var ui = app.userInterface;

		var product = app.activeProduct;
		var design = adsk.fusion.Design(product);
		if (!design) {
			ui.messageBox('No active Fusion design', 'No Design');
			return;
		}
        
        var objIsCancelled = [];
        report_type = ui.inputBox("Select Report Type: 1: All Parts and Assemblies, 2: Top Parts and Assemblies, 3:Parts Only", objIsCancelled);
        // Exit the program if the dialog was cancelled.
            if (objIsCancelled.value || (report_type != 1 && report_type != 2 && report_type != 3)) {
                adsk.terminate();    
                return;
            }
        
		// Get the root component of the active design.
		var rootComp = design.rootComponent;
        
        // Create the title for the output.
		var resultString = 'Assembly structure of ' + design.parentDocument.name + '\r\n';
        var report_title = design.parentDocument.name

        // Call the recursive function to traverse the assembly and build the output string.
		resultString = traverseAssembly(rootComp.occurrences.asList, 1, resultString);

        var fileDialog = ui.createFileDialog();
        fileDialog.isMultiSelectEnabled = false;
        fileDialog.title = "Save results as...";
        fileDialog.filter = "*.html";
        var dialogResult = fileDialog.showSave();
        
        if (dialogResult == adsk.core.DialogResults.DialogOK) {
            filename = fileDialog.filename;
            
        }
        
        var report_type_title = "All Parts and Assemblies";
        if ( report_type == 2 ) { report_type_title = "Top Parts and Assemblies"; }
        if ( report_type == 3 ) { report_type_title = "Parts Only"; }

        resultString = "<html><head><title>Part Report</title></head><body style='width:800px;'>";
        resultString += "<img src='https://scontent-sea1-1.xx.fbcdn.net/hphotos-xat1/v/t1.0-9/12032134_1041698022515115_3902714344652686388_n.png?oh=ad166bf71afa3d2120577c6d354208c6&oe=5739E323' style='height:75px; float:right;'>";
        resultString += "<div><h1 style='margin:0px;'>Parts Report" + "</h1><h3 style='margin:0px;'>" + report_title + "</h3></div>";
        resultString += "<div style='clear:both; font-weight:bold; margin-bottom:15px; font-size:20px;'>" + report_type_title  + "</div>";
        resultString += "<table style='width:800px;'>";
        resultString += "<tr><td style='font-weight:bold; width:200px;'>Part Number</td>";
        resultString += "<td style='font-weight:bold;'>Description</td>";
        resultString += "<td style='font-weight:bold; width:100px;'>Mass</td>";
        resultString += "<td style='font-weight:bold; width:50px;'>Qty.</td>";
        
        parts.sort(compare)
        for (var i=0; i<parts.length; i++) {
            
            var part_mass = parts[i].mass;
            var mass_unit = "kg";
            
           if ( part_mass < 1 ) {
                part_mass = part_mass * 1000;
                mass_unit = "g";
            }
            
            resultString += "<tr>";
            resultString += "<td>" + parts[i].part_number + "</td>";
            resultString += "<td>" + parts[i].desc + "</td>";
            resultString += "<td>" + part_mass.toFixed(3) + " " + mass_unit + "</td>";
            resultString += "<td>" + parts[i].qty + "</td>";
            resultString += "</tr>";
        }
        resultString += "</table>";
        resultString += "<div style='font-size:12px; margin-top:15px;'>Report generated on " + formatAMPM() + " UTC by " + app.currentUser.displayName + "</div>";
        resultString += "</body></html>";
        
        adsk.writeFile(filename, resultString); 

	} catch (err) {
        if (ui) {
            ui.messageBox('Unexpected error: ' + err.message);
        }
    }

    adsk.terminate();
}


// Performs a recursive traversal of an entire assembly structure.
function traverseAssembly(occurrences, currentLevel, inputString) {
    for (var i = 0; i < occurrences.count; ++i) {
        var occ = occurrences.item(i);

        //Check to see if part is already in the array
        var partFound = false;
        var indexVal;
        for (var partlooper=0; partlooper<parts.length; partlooper++) {
            if (parts[partlooper].part_number == occ.component.partNumber) { 
                partFound = true; 
                indexVal = partlooper;
            } 
        }

        if (report_type == 1 || report_type == 2 || (report_type == 3 && occ.childOccurrences.count == 0) ) {
        if ( partFound ) {
            parts[indexVal].qty++;
        } else {
            var part = {part_number:occ.component.partNumber, desc:occ.component.description, qty:1, mass:occ.component.physicalProperties.mass}
            parts.push(part);
        }
        }
        if (occ.childOccurrences && report_type !=2 ) {
            inputString = traverseAssembly(occ.childOccurrences, currentLevel + 1, inputString);
        }
    }

    return inputString;
}

function formatAMPM() {
    var d = new Date(),
    minutes = d.getUTCMinutes().toString().length == 1 ? '0'+d.getUTCMinutes() : d.getUTCMinutes(),
    hours = d.getUTCHours().toString().length == 1 ? '0'+d.getUTCHours() : d.getUTCHours(),
    ampm = d.getUTCHours() >= 12 ? 'pm' : 'am',
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
    days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    
    return d.getFullYear() + '-' + d.getUTCMonth() + '-' + d.getUTCDate() + ' ' + hours + ':' + minutes;
}


function compare(a,b) {
  if (a.part_number < b.part_number)
    return -1;
  else if (a.part_number > b.part_number)
    return 1;
  else 
    return 0;
}