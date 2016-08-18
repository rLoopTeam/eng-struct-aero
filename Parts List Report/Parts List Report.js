//Author - Chris Smith
//Description-Creates a parts report based on active model. Can run in 3 different modes, 1:) all parts and assemblies are included, 
//2:) Just the first level parts and assemblies are listed, and 3:) Only the parts are listed (no assemblies).

/*globals adsk*/
var filename;
var parts = new Array(0);
var report_type;
const CM_TO_IN = 0.393701;


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
        report_type = ui.inputBox("Select Report Type: 1: All Parts and Assemblies, 2: Top Parts and Assemblies, 3:Parts Only, 4: All Parts w/ COM", objIsCancelled);
        // Exit the program if the dialog was cancelled.
            if (objIsCancelled.value || (report_type != 1 && report_type != 2 && report_type != 3 && report_type != 4)) {
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
        if ( report_type == 4 ) { report_type_title = "Parts with Center of Mass"; }

        resultString = "<html><head><title>Part Report</title></head><body style='width:800px;'>";
        resultString += "<div style='background-color:rgba(0, 176, 224, 0.6);'>";
        resultString += "<img src='http://rloop.org/assets/images/logo-white.png' style='height:75px; float:right;'>";
        resultString += "<div><h1 style='margin:0px;'>Parts Report" + "</h1><h3 style='margin:0px;'>" + report_title + "</h3></div>";
        resultString += "<div style='clear:both; font-weight:bold; margin-bottom:15px; font-size:20px;'>" + report_type_title  + "</div>";
         resultString += "</div>";
        resultString += "<table style='width:800px;'>";
        resultString += "<tr><td style='font-weight:bold; width:200px;'>Part Number</td>";
        resultString += "<td style='font-weight:bold; width:250px;'>Description</td>";
        resultString += "<td style='font-weight:bold; width:100px;'>Mass</td>";
        if (report_type == 4 ) {
            resultString += "<td style='font-weight:bold; width:100px;'>X (in)</td>";
            resultString += "<td style='font-weight:bold; width:100px;'>Y (in)</td>";
            resultString += "<td style='font-weight:bold; width:100px;'>Z (in)</td>";
        } else {
            resultString += "<td style='font-weight:bold; width:50px;'>Qty.</td>";
        }
        
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
            if (report_type == 4 ) {
                resultString += "<td>" + Math.round(parts[i].com_x * 1000) / 1000 + "</td>";
                resultString += "<td>" + Math.round(parts[i].com_y * 1000) / 1000 + "</td>";
                resultString += "<td>" + Math.round(parts[i].com_z * 1000) / 1000 + "</td>";
            } else {
                resultString += "<td>" + parts[i].qty + "</td>";
            }
            resultString += "</tr>";
        }
        resultString += "</table>";
        resultString += "<div style='font-size:12px; margin-top:15px;'>Report generated on " + formatAMPM() + " UTC by " + app.currentUser.displayName + "</div>";
        resultString += "</body></html>";
        
        adsk.writeFile(filename, resultString); 
        adsk.readFile(filename);

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
        
        if (report_type != 4 ){ // Don't do this for center of mass report
            for (var partlooper=0; partlooper<parts.length; partlooper++) {
                if (parts[partlooper].part_number == occ.component.partNumber) { 
                    partFound = true; 
                    indexVal = partlooper;
                } 
            }
        }

        if (report_type == 1 || report_type == 2 || ((report_type == 3 || report_type == 4) && occ.childOccurrences.count == 0) ) {
            if ( partFound ) {
                parts[indexVal].qty++;
            } else {
                // Center of Mass data is not in world coordinates
                var part = {part_number:occ.component.partNumber, desc:occ.component.description, qty:1, mass:occ.physicalProperties.mass, com_x:occ.physicalProperties.centerOfMass.x * CM_TO_IN, com_y:occ.physicalProperties.centerOfMass.y * CM_TO_IN, com_z:occ.physicalProperties.centerOfMass.z * CM_TO_IN}
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