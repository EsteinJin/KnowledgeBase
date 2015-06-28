<style>
*{font-size: 12px}
</style>

<Object style="visibility:hidden" id="MSAgent" ClassID="CLSID:D45FD31B-5C6E-11D1-9EC1-00C04FD7081F" CodeBase="http://activex.microsoft.com/activex/controls/agent2/MSagent.exe#VERSION=2,0,0,0"></Object>
<Object style="visibility:hidden" id="L&HTruVoice" ClassID="CLSID:B8F2846E-CE36-11D0-AC83-00C04FD97575" CodeBase="http://activex.microsoft.com/activex/controls/agent2/tv_enua.exe#VERSION=6,0,0,0"></Object>

<Script Language="JavaScript" For="MSAgent" Event="RequestStart(RequestObject)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

switch (RequestObject) {
	case AgentLoadRequest :
		window.status = "Loading MSAgent File From Internet For " + AgentID + " ...";
		break;
	case AgentStateRequest :
		window.status = "Loading MSAgent State From Internet For " + AgentID + " ...";
		break;
	case AgentAnimationRequest :
		window.status = "Loading MSAgent Animation From Internet For " + AgentID + " ...";
		break;
	default:
		break;
}
</Script>

<Script Language="JavaScript" For="MSAgent" Event="RequestComplete(RequestObject)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

switch (RequestObject) {
	case AgentLoadRequest :
		if(RequestObject.Status == 0) {
			window.status = "MSAgent File For " + AgentID + " Has Been Loaded Successfully !";
			if(confirm("Cannot find the MSAgent charactor file on your hard disk! \nWould you like to download the MSAgent charactor file for the next show?"))
				window.open("http://www.msagentring.org/download.asp?char="+NewAgent.toLowerCase(),"_blank","top=2000px");
		} else {
			window.status = "Cannot Load MSAgent File For " + AgentID + " From " + AgentACS + " !";
			alert("Cannot find MSAgent file from local disk or internet!");
			AgentLoad = false;
		}
		break;
	case AgentStateRequest :
		if(RequestObject.Status == 0) {
			window.status = "MSAgent State For " + AgentID + " Has Been Loaded Successfully !";
		} else {
			window.status = "Cannot Load MSAgent State For " + AgentID + " From " + AgentACS + " !";
		}
		break;
		break;
	case AgentAnimationRequest :
		if(RequestObject.Status == 0) {
			window.status = "MSAgent Animation For " + AgentID + " Has Been Loaded Successfully !";
		} else {
			window.status = "Cannot Load MSAgent Animation For " + AgentID + " From " + AgentACS + " !";
		}
		break;
		break;
	default:
		window.status = "";
		break;
}
</Script>

<Script Language="JavaScript" For="MSAgent" Event="Click(CharacterID, Button, Shift, X, Y)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

if(Button==1 && Agent.Visible) {
	Agent.Stop();
	Agent_Show("Acknowledge", "Yes sir! " + CharacterID + " is right here!");
	Agent_Show("Pleased", "What can I do for you?");
} else if(Button==4097) {
	Agent.Visible?Agent.Hide():Agent.show();
}
</Script>

<Script Language="JavaScript" For="MSAgent" Event="DblClick(CharacterID, Button, Shift, X, Y)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

if(Button==1 || Button==4097) {
	Agent.StopAll();
	if (!Agent.HasOtherClients) {
		MSAgent.Characters.Unload(AgentID);
		MSAgent.Connected = false;
		Agent = null;
		AgentLoad = false;
	}
}
</Script>

<Script Language="JavaScript" For="MSAgent" Event="Move(CharacterID, X, Y, Cause)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

var rnd_words = new Array();
rnd_words.push("Ha, I am the king of screen!");
rnd_words.push("It's good day to fly!");
rnd_words.push("Ah, your boss is on your back!");
rnd_words.push("Oh, it's fruit time, do you want an apple?");
rnd_words.push("How do you think about me?");
rnd_words.push("Hey guy, have a rest!");
rnd_words.push("Pretty girl everywhere, single me over there...");
rnd_words.push("Hi, don't you think I like neo in Matrix?");
rnd_words.push("If you think it, you will make it!");
rnd_words.push("I am so lonely, together with me, come on!");

if(Cause==2) {
	Agent_Show("random", rnd_words[GetRandomNum(0, rnd_words.length-1)]);
}
</Script>

<Script Language="JavaScript" For="MSAgent" Event="DragStart(CharacterID, Button, Shift, X, Y)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

cur_x = X - Agent.width/2;
cur_y = Y - Agent.height/2;
</Script>

<Script Language="JavaScript" For="MSAgent" Event="DragComplete(CharacterID, Button, Shift, X, Y)">
//Coded by Windy_sk <windy_sk@126.com> 20040214

Agent.Stop();
Agent_Show("Confused", "Don't move me OK?", "RestPose");
Agent.MoveTo(cur_x, cur_y);
</Script>

<Script Language="JavaScript" For="MSAgent" Event="Command(UserInput)">
var BadConfidence = 10;
if (UserInput.Confidence <= -40){
	alert("Bad Recognition!");
} else if ((UserInput.Alt1Name != "") && (Math.abs(Math.abs(UserInput.Alt1Confidence) - Math.abs(UserInput.Confidence)) < BadConfidence)) {
	alert("Bad Confidence - too close to another command !");
} else if ((UserInput.Alt2Name != "") && (Math.abs(Math.abs(UserInput.Alt1Confidence) - Math.abs(UserInput.Confidence)) < BadConfidence)) {
	alert("Bad Confidence - too close to another command !");
} else {
	switch(UserInput.Name) {
		case "ACO" :
			MSAgent.PropertySheet.Visible = true;
			break;
		case "READ":
			Agent_Show("Read");
			Agent.Speak(show.value);
			break;
		case "SAYTIME":
			Agent_Show("Suggest");
			Agent.Speak("It is now " + (new Date()) + "!");
			break;
		case "INTRO" :
			Agent_Show("Explain");
			Agent.Speak("My name is " + AgentID + ", I think I'm the best one!");
			break;
		case "AUTHOR":
			Agent_Show("Announce");
			Agent.Speak("Windy_sk <windy_sk@126.com> wrote the program, I think he's great! (^o^)");
			break;
		case "FLY":
			Agent.MoveTo(Math.round(Math.random() * screen.width - Agent.width), Math.round(Math.random() * screen.height - Agent.height));
			break;
		case "STOP":
			Agent.StopAll();
			Agent_Show("RestPose");
			break;
		default:
			break;
	}
}
</Script>

<Script language="JavaScript">
//Coded by Windy_sk <windy_sk@126.com> 20040214

function reportError(msg,url,line) {
	var str = "You have found an error as below: \n\n";
	str += "Err: " + msg + " on line: " + line;
	alert(str);
	return true;
}

window.onerror = reportError;

var Agent = null;
var AgentID, AgentACS;
var AgentLoad = false;
var AgentLoadRequest, AgentStateRequest, AgentAnimationRequest;
var AgentStates = "GesturingDown, GesturingLeft, GesturingRight, GesturingUp, Hearing, Hiding, IdlingLevel1, IdlingLevel2, IdlingLevel3, Listening, MovingDown, MovingLeft, MovingRight, MovingUp, Showing, Speaking";
var AgentAnimations = ["Acknowledge", "Alert", "Announce", "Blink", "Confused", "Congratulate", "Congratulate_2", "Decline", "DoMagic1", "DoMagic2", "DontRecognize", "Explain", "GestureDown", "GestureLeft", "GestureRight", "GestureUp", "GetAttention", "GetAttentionContinued", "GetAttentionReturn", "Greet", "Hearing_1", "Hearing_2", "Hearing_3", "Hearing_4", "Hide", "Idle1_1", "Idle1_2", "Idle1_3", "Idle1_4", "Idle2_1", "Idle2_2", "Idle3_1", "Idle3_2", "LookDown", "LookLeft", "LookRight", "LookUp", "MoveDown", "MoveLeft", "MoveRight", "MoveUp", "Pleased", "Process", "Processing", "Read", "ReadContinued", "ReadReturn", "Reading", "RestPose", "Sad", "Search", "Searching", "Show", "StartListening", "StopListening", "Suggest", "Surprised", "Think", "Uncertain", "Wave", "Write", "WriteContinued", "WriteReturn", "Writing"];
var remote = false;
var cur_x = 400, cur_y = 300;
var MoveTimer = null;

function LoadAgent(NewAgent) {
	if(AgentLoad) {
		MSAgent.Characters.Unload(AgentID);
		MSAgent.Connected = false;
		Agent = null;
	}
	AgentID = NewAgent;
	AgentACS = NewAgent + ".acs";
	MSAgent.Connected = true;
	try {
		MSAgent.Characters.Load(AgentID, AgentACS);
	} catch(e) {
		AgentACS = "http://agent.microsoft.com/agent2/chars/" + NewAgent + "/" + NewAgent + ".acf";
		remote = true;
		AgentLoadRequest = MSAgent.Characters.Load(AgentID, AgentACS);
	}
	try {
		AgentLoad = true;
		Agent = MSAgent.Characters.Character(AgentID);
		Agent.LanguageID = 0x0409;
		Agent.Balloon.Style = 0x330000F;
		
		Agent.Commands.RemoveAll();
		Agent.Commands.Visible = true;
		Agent.Commands.Caption = "MSAgent's Menu - by windy_sk";
		Agent.Commands.Add("ACO", "Advanced Character Options", "Advanced Character Options");
		Agent.Commands.Add("READ", "Read Text In Textarea", "Read Text In Textarea");
		Agent.Commands.Add("INTRO", "Introduce Yourself", "Introduce yourself");
		Agent.Commands.Add("AUTHOR", "Who Write The Program", "Who Write The Program");
		Agent.Commands.Add("SAYTIME", "What Time Is It Now", "What Time Is It Now");
		Agent.Commands.Add("FLY", "Can You Fly", "Can You Fly");
		Agent.Commands.Add("STOP", "Stop All Actions", "Stop All Actions");

		if(remote) {
			AgentStateRequest = Agent.get("state", "Showing, Thinking");
			AgentAnimationRequest = Agent.get("animation", "GetAttention, RestPose");
		}
		Agent.MoveTo(cur_x, cur_y);
		Agent.Show();
		try {
			Agent.Play("GetAttention");
		} catch(e) {
			Agent.Play("RestPose");
		}
		Agent.speak("Hi, I am " + NewAgent + ", can I help you, sir?");
		//Agent.think("Oh so bad, I just wanna take a nap...");
		if(remote) AgentStateRequest = Agent.get("state", AgentStates);
	} catch(e) {
		for(var x in e) alert(x + " - " + e[x]);
		AgentLoad = false;
	}
	return;
}

function GetRandomNum(Min,Max){
	var Range = Max - Min;
	var Rand = Math.random();
	return(Min + Math.round(Rand * Range));
}

function Agent_Show() {
	if(!AgentLoad) return;
	var argv = Agent_Show.arguments;
	var argc = argv.length;
	if(!Agent.Visible) Agent.Show();
	for(var i=0; i<argc; i+=2) {
		if(argv[i] == "random") argv[i] = AgentAnimations[GetRandomNum(0, AgentAnimations.length-1)];
		try {
			if(remote) Agent.get("animation", argv[i]);
			Agent.Play(argv[i]);
		} catch(e) {
			Agent.Play("RestPose");
		}
		if(typeof(argv[i+1]) != "undefined" && argv[i+1] != "") Agent.speak(argv[i+1]);
	}
	return;
}

function Agent_Show_All(mode) {
	if(Agent==null || !AgentLoad) return;
	if(!Agent.Visible) Agent.Show();
	show.value = "Animation for " + AgentID;
	for(var i=0; i<AgentAnimations.length; i++){
		show.value += "\ntesting '" + AgentAnimations[i] + "' - 	";
		try {
			if(remote) Agent.get("animation", AgentAnimations[i]);
			Agent.Play(AgentAnimations[i]);
			Agent.speak(AgentID + " can play '" + AgentAnimations[i] + "'!");
			show.value += "OK!";
		} catch(e) {
			Agent.Play("RestPose");
			Agent.speak(AgentID + " play '" + AgentAnimations[i] + "' failed !");
			show.value += "Failed!";
		}
		if(!mode) Agent.Stop();
	}
	return;
}

function Agent_Move(){
	if(!AgentLoad) return;
	if(GetRandomNum(1, 10) > 6){
		var Scr_width = window.screen.width - 100;
		var Scr_Height = window.screen.height - 100;
		Agent.MoveTo(GetRandomNum(0,Scr_width),GetRandomNum(0,Scr_Height));
	}
	MoveTimer = setTimeout("Agent_Move()",5000);
	return;
}

window.onhelp = function() {
	if(!AgentLoad) return;
	if(Agent == null) {
		LoadAgent(AgentID);
	} else {
		if(!Agent.Visible) Agent.Show();
		Agent.speak("Can I help you, sir?");
	}
	return false;
}

LoadAgent("Merlin");
</Script>

Charactor Select : 
<SELECT name="Agent_select" onchange="LoadAgent(this[this.selectedIndex].text)">
	<Optgroup label="Offical Charactors">
		<OPTION>Merlin</OPTION>
		<OPTION>Peedy</OPTION>
		<OPTION>Genie</OPTION>
		<OPTION>Robby</OPTION>
	</Optgroup>
	<Optgroup label="Charactors from Office">
		<OPTION>CLIPPIT</OPTION>
		<OPTION>courtney</OPTION>
		<OPTION>DOLPHIN</OPTION>
		<OPTION>DOT</OPTION>
		<OPTION>earl</OPTION>
		<OPTION>F1</OPTION>
		<OPTION>LOGO</OPTION>
		<OPTION>MNATURE</OPTION>
		<OPTION>MNKYKING</OPTION>
		<OPTION>OFFCAT</OPTION>
		<OPTION>qmark</OPTION>
		<OPTION>ROCKY</OPTION>
		<OPTION>rover</OPTION>
		<OPTION>SAEKO</OPTION>
	</Optgroup>
</SELECT>
<input type="checkbox" id="auto_move" onclick="this.checked?Agent_Move():clearTimeout(MoveTimer)"><label for="auto_move">Auto Move</label>
<br /><br />
<textarea id="show" style="width: 400px; height: 200px">
Hello everybody, I am Office Agent, I hope I am useful to you !
</textarea>
<br />
<input type="button" value="Test Animation" onclick="Agent_Show_All(show_animation.checked)">
<input type="checkbox" id="show_animation"><label for="show_animation">Show Animation</label>