function check() {
	if (document.add.title.value.length < 4 || document.add.title.value.length > 100) {		
		alert ("���ⲻС��4λ�����ߴ���100λ");
		document.add.title.focus();		
		return false;
	}
	
	if (document.add.name.value == "") {
		alert("�����˲���Ϊ��")	;
		document.add.name.focus();
		return false;
	}
	
	if (document.add.info.value.length < 10 || document.add.info.value.length > 200) {		
		alert ("���ݼ�鲻��С��10λ������200λ");
		document.add.info.focus();		
		return false;
	}
	

	
	return true;
}