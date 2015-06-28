<form>
<select name="lan">
<option value="en|de">英語 翻譯成 德語</option>
<option value="en|es">英語 翻譯成 西班牙語</option>
<option value="en|fr">英語 翻譯成 法語</option>
<option value="en|it">英語 翻譯成 意大利語</option>
<option value="en|pt">英語 翻譯成 葡萄牙語</option>
<option value="en|ja">英語 翻譯成 日語 BETA</option>
<option value="en|ko">英語 翻譯成 朝鮮語 BETA</option>
<option value="en|zh-CN" >英語 翻譯成 中文(簡體) BETA</option>
<option value="de|en">德語 翻譯成 英語</option>
<option value="de|fr">德語 翻譯成 法語</option>
<option value="es|en">西班牙語 翻譯成 英語</option>
<option value="fr|en">法語 翻譯成 英語</option>
<option value="fr|de">法語 翻譯成 德語</option>
<option value="it|en">意大利語 翻譯成 英語</option>
<option value="pt|en">葡萄牙語 翻譯成 英語</option>
<option value="ja|en">日語 翻譯成 英語 BETA</option>
<option value="ko|en">朝鮮語 翻譯成 英語 BETA</option>
<option value="zh-CN|en">中文(簡體) 翻譯成 英語 BETA</option>
<input style="FONT-SIZE: 12px" type="button" value="Go->" name="Button1" onClick="javascript:window.open('translate.asp?urls=' document.location '&lan=' lan.value,'_self','')">
</select>
</form>