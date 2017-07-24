var popup = null;

function getToday () {
// get today's date as a string
var tmpdate = new Date();
return ( tmpdate.getMonth() +1 ) + "/" + tmpdate.getDate() + "/" + tmpdate.getFullYear();
}

function weekday ( xdate ) {
// get weekday number ( Monday = 1 )
var date = new Date();
return date.getDay();
}

function getTodayDateNum(){
	var xtmpdate = new Date();
	return xtmpdate.getDate();
}

// changes color of Calendar cells
var rowPicked = new String (getTodayDateNum());  // needs to be today's date onLoad.....
if(rowPicked.length == 1){rowPicked = "0" + rowPicked;}
function setColor(control){
	var cell = document.getElementById(rowPicked);
	cell.className = "normal";
	if(control == null){
		control=document.getElementById(rowPicked);
	}
	if(control != null){  // thinking of cases where you might not have the 31st in the next month chosen....
		control.className="tag";
		rowPicked=control.id;
	}
}

function helpPop ( xkey ) {
	var xURL = getPath() + "/luhelp/" + xkey + "?OpenDocument";
	openPopup( xURL , "HelpPop" , "500" , "300" )
}

function getCookieStr (xname) {
var xcookie = document.cookie;
if ( xcookie.indexOf ( xname ) < 0 ) { return ""; } 
if ( xcookie.indexOf ( ";" ) > 0 ) { // get end of cookie
	var xend = xcookie.indexOf ( ";" ); }
	else { var xend = xcookie.length - 1; }
var xstart = xcookie.indexOf ( xname ) + xname.length + 1;
return unescape ( xcookie.substring ( xstart , xend ) );
}

function delCookie(xcookie) {
	var lastyear = new Date();
	lastyear.setFullYear(lastyear.getFullYear() - 20);
	document.cookie = xcookie + "=0; path=/ expires=" + lastyear.toGMTString();
}

function jsDatePicker( szField, szDate, szAction){
 var form = document.forms[0];
 var field = form.elements[szField];
 if(szAction == "1"){
  field.value=szDate;
 }
 return field.value;
}

function checkRefresh() {
// Checks 'DateTime' Notes field on form (24 hr time format) and if the page is X seconds old or more, reloads the view!
	var thisform = window.document.forms[0];
	var xthen = thisform.DateTime.value;
	xthen = ( xthen.substring (0,2) *3600 ) + (xthen.substring(3,5) * 60 ) + ( xthen.substring(6,8) * 1);
	var xnow = new Date();
	var xnowmin = ( xnow.getHours() * 3600 ) + ( xnow.getMinutes() * 60 ) + ( xnow.getSeconds() *1);
	var secs = 5;
	if ( xnowmin - xthen > secs ){
		return window.location.reload();}
}

function deleteDocument(navURL) {
if ( confirm("Are you sure you want to remove this document?") ) {
	var thisform = window.document.forms[0];
	window.location.href= getPath() + "/(DeleteDoc)?OpenAgent&UNID=" +thisform.docunid.value + "&VIEW=" + navURL;
	}
}

function DatePicker(whichfield) {
//for the calendar picker buttons
var pathname = (window.location.pathname);
window.document.forms[0].DatePickerFieldName.value = whichfield;   //let the calendar know which field to fill out
window.focus();
newWindow = window.open(pathname.substring(0,(pathname.lastIndexOf('.nsf')+4)) + '/Calendar?OpenPage','Calendar','status=yes,scrollbars=no,resizable=yes,top=120,left=200,width=200,height=225');
newWindow.focus()
}

function doLoad() {
if ( parent.frames.length == 0) {  // no frames found
	window.setTimeout( 'setFrames()',500); // need to pause the code, or else we got in a onLoad - no frames loop
	}
}

function setFrames() {
var xFS = "MainFS";
var xframe = "MainR";
var xpath = window.location.pathname;
var xnum = xpath.length;
var xUNID = xpath.substring(xnum-32,xnum); // xpath does not have any "?" commands, etc. so get the end -- the UNID -- only
var xpathname = xpath.substring(0,(xpath.lastIndexOf('.nsf')+4));
window.location.href = xpathname + "/"+ xFS + "?OpenFrameSet&Frame="+xframe + "&Src=" + xpathname +"/0/"+ xUNID + "?OpenDocument";
}

function checkDate(field) {  // returns true or false if date not formatted correctly
var xval = field.value;
var pattern = /\d\d\W\d\d\W\d\d\d\d\b/;
var pattern2 = /\d\W\d\d\W\d\d\d\d\b/;
var pattern3 = /\d\d\W\d\W\d\d\d\d\b/;
var pattern4 = /\d\W\d\W\d\d\d\d\b/;
var result = pattern.test ( xval ) ;
var result2 = pattern2.test ( xval ) ;
var result3 = pattern3.test ( xval ) ;
var result4 = pattern4.test ( xval ) ;
if ( result == true || result2 == true || result3 == true || result4 == true ) {
	return true; }
	else { return "Check Format: " + field.title + " Field"; }
}

function checkVal(){
var thisform = window.document.forms[0];
var xval = new String();
var xreturn = new String();
var xdfields = new String();
for (i = 0; i < thisform.elements.length; i++) {
	var field = thisform.elements[i];
	var xfname = field.name.toLowerCase();
// check any non-hidden date fields	
	if ( xfname.indexOf ( "date" ) > -1 && field.type != "hidden" && field.value != "" ) { 
		if ( checkDate(field) != true ) {
			xdfields = xdfields + "\r\n" + checkDate(field); }
	}
// check all required fields	
	if ( field.REQUIRED=="TRUE"){
		if (field.type=="select-one" || field.type=="select-multiple"){
			xval =getSelectValue (thisform.elements[i]);}
			else {xval = field.value;
		}
		if (xval == "" || xval == null || xval.indexOf("<<") > -1 ){
			xreturn = xreturn + "\r\n" + field.title;
		}
	}
}
xreturn = xreturn + xdfields;
if ( xreturn != "" ){
	alert ( "Please fill out the following fields before Saving: " + "\r\n" + xreturn );
	return false; }
	else { return true; }
}

function popImage(image){
	var baseRef = image.href;
	var url = baseRef.substring(0,baseRef.length - 4) + "LARGE.jpg";
	openPopup(url,"ImageWin","650","850");
}

function openPopup( url , name , height , width ) {
	closePopup();	//ONLY ONE POPUP, and the reference is stored in the variable popup (hardcoded!)
	var xh = height ; 
	var xw = width;
	var dialogparam = ",scrollbars=yes,status=no,location=no,menubar=no,resizeable=yes";
	var screenparam = "height="+xh+",width="+xw ;
	params = screenparam+dialogparam;
	popup = window.open( url , name , params );
}

function closePopup() {
if (popup != null && popup.open) popup.close();
popup = null;
}

function getSelectValue (select) {
if (select.value == "") return "";
if (select.value == null ) return "";
if (select.type == "select-one") {
   return select.options[select.selectedIndex].value};  
// Type is "select-multiple"
   var answer = "";
   for (var i = 0; i < select.options.length; i++) {
      var option = select.options[i];
      if (option.selected) answer += "," + option.value;
   }
   return answer.substring(1);
}   

function getPath(){
var xtmp = window.location.pathname.toLowerCase();
xpath = window.location.pathname;
var xindex=xtmp.indexOf(".nsf") +4;
return xpath.slice(0,xindex);
}

function repSubstring ( xval , xfrom , xto ) {
// Replace subsrtrings
	while ( xval.indexOf ( xfrom ) > 0 ) { 
		xval = xval.slice ( 0 , xval.indexOf ( xfrom ) ) + xto + xval.slice ( xval.indexOf ( xfrom ) + xfrom.length , xval.length );
	}
return xval;
}

function launchNAB(xfieldname,xtype) {
// two parameters required: 1. Name of the field that needs to be set with names, 2. Single or multiple select
// Note that we're passing in URL strings that the form grabs
var xpathname = getPath();
return openPopup( xpathname + "/nabpopup?OpenForm&" + xfieldname + "@@" + xtype, "NABWin" , "200" , "750" ) ;
}