function clearTextbox(obj) {
	defaultTextboxValue = obj.value;
	obj.value = '';
}

function refillTextbox(obj) {
	if (obj.value.length == 0) {
		obj.value = defaultTextboxValue;
	}
}
var open2='';
function newsExpImp(id) {
if (document.getElementById('n_'+id).style.display=='block'){
document.getElementById('n_'+id).style.display='none';
if(open2==id)open2='';
}
else{
if(open2!='')document.getElementById('n_'+open2).style.display='none';
document.getElementById('n_'+id).style.display='block';
open2=id;
}
}
function checkSearchField() {
	if (document.searchform.bbp_search.value.length < 3) {
		alert('The keyword must at least be 3 characters in length.');
		return false;
	}
	return true;
}
