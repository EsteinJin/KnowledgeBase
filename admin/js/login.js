function check() {
	if (document.login.adminname.value == "") {		
		alert ("����������½���û���!");
		document.login.adminname.focus();		
		return false;
	}
	if (document.login.adminname.value.length < 2) {
		alert("�û�����������2λ!");
		document.login.adminname.focus();
		return false;
	}
	if (document.login.adminpass.value == ""){
		alert ("���������ĵ�½���룡");	
		document.login.adminpass.focus();
		return false;
	}
	if (document.login.adminpass.value.length < 6){
		alert ("���벻������6λ��");	
		document.login.adminpass.focus();
		return false;
	}
	if (document.login.yzm.value.length != 4){
		alert ("��֤�������4λ��");	
		document.login.yzm.focus();
		return false;
	}
	//is�ǣ�NaN  N��ʾNot,a,Number
	if (isNaN(document.login.yzm.value)) {
		alert ("��֤����������֣�");	
		document.login.yzm.focus();
		return false;
	}
	return true;
}
