//helppane.js Version 1.7
var H_URL_BASE='',H_TOPIC='',H_KEY='',L_H_TEXT='',H_FILTER='',H_BRAND='',bSearch=false;
var H_CONFIG='',L_H_APP='',L_CONTACTUS_URL='';
var H_BURL='/cgi-bin/dasp/HelpPane.asp',H_TARG='',H_VER='1.7';
var h_win,H_OTHER='',bResize=true;
function DoHelp(iNm) {
var sQP='?',W,H,sWD,sc=screen.width,bIE4PC;
var agent = navigator.userAgent.toLowerCase();
var app = navigator.appName.toLowerCase();		
sQP+='H_VER='+H_VER;
if (H_BRAND!='') sQP+='&BrandID='+H_BRAND
if (H_FILTER!='') sQP+='&Filter='+H_FILTER 
sQP+=(bSearch) ? '&SEARCHTERM='+escape(H_KEY)+'&S_TEXT='+escape(L_H_TEXT):'&TOPIC='+H_TOPIC
if (typeof(v1)!="undefined") sQP+='&v1='+escape(v1)
else sQP+='&v1='+escape(document.location.protocol + "//" + document.location.hostname)
sQP+='&v2='+escape(document.location.search);
if (typeof(H_CONFIG) != "undefined" && (self.name == null || self.name == "" || self.name == "msnMain")) self.name = H_CONFIG.substring(0,H_CONFIG.indexOf("."));
sQP+='&tmt='+escape(window.name);
if (sc<=800) sQP+="&sp=1";
W=(sc<= 800 && agent.indexOf("mac")==-1)?180:230;
H=(agent.indexOf("windows")>0 && agent.indexOf("aol")>0) ? screen.availHeight-window.screenTop-22:screen.availHeight//*AOL
var agent_isMSN = false, vi = agent.indexOf('msn ');
if (vi > -1) {
agent_isMSN = agent.substring(vi+4);
agent_isMSN = parseFloat(agent_isMSN.substring(0,agent_isMSN.indexOf(";")));
agent_isMSN = (agent_isMSN != NaN && agent_isMSN >= 6)
}
if (agent_isMSN){
window.external.showHelpPane(H_URL_BASE+'/frameset.asp'+sQP+'&H_APP='+escape(L_H_APP)+'&INI='+H_CONFIG,W)
}
else if (agent.indexOf('webtv')>0 || agent.indexOf('msn companion')>0){
top.location.replace(H_URL_BASE+'/frameset.asp'+sQP+'&H_APP='+escape(L_H_APP)+'&INI='+H_CONFIG)
}
else {
sWD="toolbar=0,status=0,menubar=0,width="+W+",height="+H+",left="+(sc-W)+",top=0,resizable=1";
bResize=false;
bIE4PC = agent.indexOf("msie 4")>0 && agent.indexOf("aol")<0 && agent.indexOf("mac")<0
if (H_TARG=='') H_TARG = (bIE4PC)?'_help17':'_help';
if (iNm != null) H_TARG+=iNm;
if (bIE4PC && h_win!=null && !h_win.closed) h_win.location.replace(H_BURL+sQP)
else h_win=window.open(H_BURL+sQP,H_TARG,sWD);
if (h_win && agent.indexOf("mac")<0 && app.indexOf("netscape")<0) h_win.opener=self//*IE5+PC
}
}
L_H_APP = "MSN+Hotmail";
H_URL_BASE = "http://help.msn.com/EN_AU";
H_CONFIG = "hotmailv6.ini";
bSearch = false;var L_SignInAB_Text = "Sign in";
var L_SignInABcont_Text = " to MSN Messenger to see who is online.";
var L_IsOnline_Text = "Online";
var L_IsOffline_Text = "Offline";
var L_IsBusy_Text =  "Busy";
var L_IsAway_Text = "Away";
var L_IsBRB_Text = "Be Right Back";
var L_IsOnThePhone_Text = "On The Phone";
var L_IsOutToLunch_Text = "Out To Lunch";
var L_Add_Text = "Add ";
var L_ContactList_Text = " to My Messenger Contacts.";
var L_MyMessList_Text = " to <b>My Messenger Contacts</b>";
var L_UseWithMSNmessenger_Text = "Use with MSN Messenger";
var WatchCount = "";
var L_OnlineStatus_Text = "Online Status";
var L_DloadMess_Text = "Download MSN Messenger";
var L_DloadMessCont_Text = " to see who is online, exchange instant messages, and more!";
var LocalUserEmail = "";
function MsngrCreateObj() {
MsngrObj = new ActiveXObject("MSNMessenger.HotmailControl");
LocalUserEmail = HMname.innerText.replace(/\s+/g,"");
}
function MsngrIsStateOnline(state)
{
var ret;
switch (state)
{
case 2:
//online
case 6:	
//invisible
case 10:
//busy
case 14:
//be right back
case 18:
//idle
case 34:
//away
case 50:
//on the phone
case 66:
//out to lunch
ret = true;
break;
default:
ret = false;
break;
}
return ret;
}
function MsngrGetContact(eMail,location)
{
var ret;
var img;
var msg;
var ContactState = MsngrObj.GetUserStatus(eMail);
switch (ContactState)
{
case 1:
// Offline
img = 'src="http://64.4.14.24/icon_messenger1.gif"';
msg = L_IsOffline_Text;
break;
case 2:
// Online
img = 'src="http://64.4.14.24/icon_messenger0.gif"';
msg = L_IsOnline_Text;
break;
case 10:
// Busy
img = 'src="http://64.4.14.24/icon_messenger3.gif"';
msg = L_IsBusy_Text;
break;		
case 14:
// Be Right Back
img = 'src="http://64.4.14.24/icon_messenger2.gif"';
msg = L_IsBRB_Text;
break;	
case 18:
// Away
img = 'src="http://64.4.14.24/icon_messenger2.gif"';
msg = L_IsAway_Text;
break;
case 34:
// Away
img = 'src="http://64.4.14.24/icon_messenger2.gif"';
msg = L_IsAway_Text;
break;
case 50:
// On the Phone
img = 'src="http://64.4.14.24/icon_messenger3.gif"';
msg = L_IsOnThePhone_Text;
break;		
case 66:
// Out To Lunch
img = 'src="http://64.4.14.24/icon_messenger2.gif"';
msg = L_IsOutToLunch_Text;
break;
}
if (location=="InBox")
{
ret = '<span style="position:relative;width:21px;top:2px;"><A HREF="JavaScript:MsngrIM(%22' + eMail + '%22)"><IMG alt="" '+img+' border=0></a></span>';
}
else if (location=="AddressBookList")
{
if (ContactState==2)
ret = '<span style="position:relative;"><A HREF="JavaScript:MsngrIM(%22' + eMail + '%22)"><IMG alt="" '+img+' border=0></a>&nbsp;&nbsp;<nobr>('+msg+')</nobr></span>';
else
ret = '<span style="position:relative;"><A HREF="JavaScript:DoCompose(%22ADDR'+eMail+'%22)"><IMG alt="" '+img+' border=0></a>&nbsp;&nbsp;<nobr>('+msg+')</nobr></span>';
}
else if (location=="ReadMessage")
{
ret = '<table border=0 cellspacing=0 cellpadding=0><tr valign="middle"><td><A HREF="JavaScript:MsngrIM(%22' + eMail + '%22)"><IMG alt="" '+ img +' border=0></a></td><td>&nbsp;<font class="s"><A HREF="JavaScript:MsngrIM(%22' + eMail + '%22)">' + document.all.msgFromName.value +'</a> (' + msg +')</font></td></tr></table>';
}
return ret;
}
function MsngrIM(eMail)
{
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()) && MsngrIsUser())
{
if ( MsngrIsStateOnline(MsngrObj.GetUserStatus(eMail)) && (LocalUserEmail != eMail))
MsngrObj.InstantMessage(eMail);
else
MsngrObj.ShowContactList();
}
else
MsngrObj.Signin(LocalUserEmail);
}
function MsngrContacts() {
CA(1);  //Check Box Check
var DataTable = document.all.ListTable;
var HdrCell = DataTable.rows[1].insertCell();
HdrCell.style.border='none';
HdrCell.innerHTML = "<font class='sw'><b>&nbsp;<nobr>"+L_OnlineStatus_Text+"</nobr></b></font>"
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()) && MsngrIsUser())
{
for (i=2; i<DataTable.rows.length; i++)
{
var ThisCell = DataTable.rows[i].insertCell();
if (DataTable.rows[i].name)
{
var E = DataTable.rows[i].name;
if (MsngrObj.GetUserStatus(E)!=0)
ThisCell.innerHTML = MsngrGetContact(E,"AddressBookList");
else
ThisCell.innerHTML="&nbsp;";
}
else
{
if (DataTable.rows[i].cells[0].children[0].tagName=="INPUT")
ThisCell.innerHTML="&nbsp;"
else
ThisCell.style.borderBottom="1px";
}
}
}
else
{
var PromoCell = DataTable.rows[2].insertCell();
PromoCell.style.border='none';
PromoCell.style.paddingLeft='10px';
PromoCell.style.width=130;
PromoCell.bgColor="#FFFFFF";
PromoCell.rowSpan=DataTable.rows.length;
PromoCell.vAlign="top";
var eMail = LocalUserEmail;
PromoCell.innerHTML="<br><a href='JavaScript:MsngrObj.Signin(\""+eMail+"\");'>"+L_SignInAB_Text+"</a>"+L_SignInABcont_Text;
}
}
function MsngrContactsMR()
{
var DataTable = document.all.ListTable;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
for (i=1; i<DataTable.rows.length; i++)
{
var E = DataTable.rows[i].cells[0].children[0].children[0].innerHTML;
if (MsngrIsStateOnline(MsngrObj.GetUserStatus(E)))
DataTable.rows[i].cells[1].innerHTML = MsngrGetContact(E,"AddressBookList");
}
}
}
function MsngrContactsPROMO()
{
CA(1);  //Check Box Check
var DataTable = document.all.ListTable;
var HdrCell = DataTable.rows[1].insertCell();
HdrCell.style.border='none';
HdrCell.innerHTML = "<font class='sw'><b>&nbsp;<nobr>"+L_OnlineStatus_Text+"</nobr></b></font>"
var PromoCell = DataTable.rows[2].insertCell();
PromoCell.style.border='none';
PromoCell.style.paddingLeft='10px';
PromoCell.style.width=130;
PromoCell.bgColor="#FFFFFF";
PromoCell.rowSpan=DataTable.rows.length;
PromoCell.vAlign="top";
PromoCell.innerHTML="<br><font class='s'><a href='http://g.msn.com/1HM7ENAU/144??PS="+HMPS+"'>"+L_DloadMess_Text+"</a>"+L_DloadMessCont_Text+"</font>";
}
function MsngrInBox()
{
CA(1); //Check Box Check
var DataTable;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
DataTable = document.all.MsgTable;	
for (i=1; i<DataTable.rows.length; i++)
{
if (DataTable.rows[i].cells.length>=6)
{
var E = DataTable.rows[i].cells[0].name;
E = E.replace(/\s+/g,"");
if (MsngrIsStateOnline(MsngrObj.GetUserStatus(E)))
{
DataTable.rows[i].cells[2].innerHTML = '<table cellpadding=0 cellspacing=0 style="position:relative;"><tr><td nowrap style="border:none">'+DataTable.rows[i].cells[2].innerHTML+'</td><td width="100%"></td><td style="border:none">'+MsngrGetContact(E,"InBox");+'</td></tr></table>'
}
}
}
}
}
function MsngrSMCPROMO()
{
if (eval(WatchCount)==1)
{
document.all.IMsngrTitle.innerHTML=content[4];
document.all.IMsngrContent.innerHTML=L_WCS_Text+content[5];
}
else if (eval(WatchCount) >= 1)
{
document.all.IMsngrTitle.innerHTML=content[4];
document.all.IMsngrContent.innerHTML=WCP+content[5];
}
else
{
document.all.IMsngrTitle.innerHTML=content[0];
document.all.IMsngrContent.innerHTML=BullImg+content[1]+"<br>"+BullImg+content[2]+"<br>"+BullImg+content[3]+"<br>"+MsgrLink+"<br>";
}
}
function MsngrQuickList()
{
var qlt = document.all.quicklist;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
for (i=1; i<qlt.rows.length; i++)
{
Data = qlt.rows[i].cells[5];
if ("undefined" != typeof(Data.name))
{
if ( (Data.name.indexOf(",") == -1) && (Data.name.indexOf("@") != -1) && (Data.name !="" ))
{
var Address = Data.name.match(/(\w+@.+\.\w+)/);
var E = Address[1];
if (MsngrIsStateOnline(MsngrObj.GetUserStatus(E)))
{
Data.innerHTML = '<table cellpadding=0 cellspacing=0 style="position:relative;"><tr><td style="border-bottom:none;" nowrap>'+Data.innerHTML+'</td><td width="100%"></td><td style="border-bottom:none;">'+MsngrGetContact(E,"InBox")+'</td></tr></table>'
}
}
}
}
}
}
function MsngrReadMessage()
{
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()) && MsngrIsUser())
{
var strFromText = document.all.FromText.value.toLowerCase();
if (document.all.MsgHeaders.rows[1].cells[0].id == "From" && ValidateEmail(strFromText))
{
if ( (MsngrIsStateOnline(MsngrObj.GetUserStatus(strFromText))) && (strFromText != null))
{
var TheHtml = MsngrGetContact(strFromText,"ReadMessage");
MyTable = document.all.MsgHeaders;
MyRow = MyTable.insertRow();
MyRow.insertCell();
MyRow.cells[0].innerHTML=TheHtml;
}
else
{
MsngrAllowedDomain();
MsngrExemptList();
MsngrExemptListObj[LocalUserEmail.toLowerCase()] = 1
var results = strFromText.match(/(.+)@(.+)/);
if ((MsngrAllowedDomainObj[results[2]]) && (MsngrExemptListObj[strFromText]==null) && (MsngrObj.GetUserStatus(strFromText)==0))
{
var TheHtml = '<font class="s"><A HREF="'+ SaveAddLink +'">' + L_Add_Text + strFromText + '</a>' + L_ContactList_Text;
MyTable = document.all.MsgHeaders;
MyRow = MyTable.insertRow();
MyRow.insertCell();
MyRow.cells[0].innerHTML=TheHtml;
}
}
}
}
}
function MsngrAllowedDomain()
{
MsngrAllowedDomainObj = new Object();
MsngrAllowedDomainObj["hotmail.com"] = "1";
MsngrAllowedDomainObj["msn.com"] = "1";
}
function MsngrSaveAddresses()
{
var DataTAble;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
DataTable = document.all.msngrdata;
for (i=1;i<DataTable.rows.length;i++)
{
var E = DataTable.rows[i].cells[0].children[0].rows[0].cells[2].id;
var from = DataTable.rows[i].cells[0].children[0].rows[0].cells[0].id;
if ( (MsngrObj.GetUserStatus(E) == 0) && (ValidateEmail(E)) && MsngrIsUser() )
{
E = MsngrHTMLencode(E);
MyRow = DataTable.rows[i].cells[0].children[0].insertRow(4);
MyRow.insertCell();
MyRow.cells[0].colSpan=4;
MyRow.cells[0].style.backgroundColor="#DBEAF5";
MyRow.cells[0].align = "center";
MyRow.cells[0].innerHTML = '<font class="s"><input type="checkbox" name="msngr'+E+'" value="'+E+'"'+from+' onClick="CheckCheckAll();"> '+L_Add_Text+E+L_MyMessList_Text+'</font>'+'';
}
}
}
}
function MsngrSaveAddressesSubmit()
{
if ("undefined" != typeof(MsngrObj) && document.domsgaddresses)
{	
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
for (var i=0;i<document.domsgaddresses.elements.length;i++)
{
var e = document.domsgaddresses.elements[i];
if ( (e.name != 'allbox') && (e.name.match(/msngr/)) && (e.checked) )
MsngrObj.AddContact(LocalUserEmail,e.value);
}
}
}
}
function MsngrEditContacts()
{
document.addr.alias.focus();
if ( (IsABmigrationComplete!="0") && (ContactType!="messenger") )
{
var DataTable = document.all.msngrdata;
var DataCell = DataTable.rows[5].cells[1];
DataCell.style.padding="0px";
DataCell.innerHTML="<input type='checkbox' name='msngr' value='' checked disabled> <font class='s' color='#c0c0c0'>"+L_UseWithMSNmessenger_Text+"</font>";
if (adfrm.addrim.value!="")
{
if (ValidateEmail(adfrm.addrim.value))
{
adfrm.msngr.disabled=false;
document.all.msngrdata.rows[5].cells[1].children[1].color="#000000";
}
}
document.addr.alias.focus();
}
}
function MsngrEditContactsSubmit()
{
if ( (IsABmigrationComplete!="0")  && (ContactType!="messenger") )
{
if ("undefined" != typeof(MsngrObj))
{
var DataTable = document.all.msngrdata;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
if ( (document.addr.msngr.checked) && (document.addr.msngr.disabled==false) && (MsngrObj.GetUserStatus(document.addr.addrim.value) == 0) )
MsngrObj.AddContact(LocalUserEmail,document.addr.addrim.value);
}
}
}
}
function MsngrNotJunkMail()
{
var DataTAble;
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()) && MsngrIsUser() )
{
DataTable = document.all.msngrdata;
var E = DataTable.rows[0].cells[1].id;
if (MsngrObj.GetUserStatus(E) == 0)
{
MyRow = DataTable.insertRow(4);
MyRow.insertCell();
MyRow.cells[0].colSpan=4;
MyRow.cells[0].align = "left";
MyRow.cells[0].innerHTML = '<font class="s"><input type="checkbox" name="msngr'+E+'" value="'+E+'" tabindex=10 checked> '+L_Add_Text+E+L_MyMessList_Text+'</font>';
}
}
}
function MsngrNotJunkMailSubmit()
{
if ("undefined" != typeof(MsngrObj))
{	
if (MsngrIsStateOnline(MsngrObj.GetLocalUserStatus()))
{
for (var i=0;i<document.addtoaddressbook.elements.length;i++)
{
var e = document.addtoaddressbook.elements[i];
if ( (e.name != 'allbox') && (e.name.match(/msngr/)) && (e.checked) )
MsngrObj.AddContact(LocalUserEmail,e.value);
}
}
}
}
function MsngrHMHomePROMO()
{
if (eval(WatchCount) == 1)
content = '<font class="s">' + L_WCS_Text + MsgrLink +'</font>';
else if (eval(WatchCount)>=1)
content = '<font class="s">' + WCP + MsgrLink +'</font>';
else
content = '<font class="s">' + content[Math.round(Math.random()*1)] + MsgrLink +'</font>' + '&nbsp;';
document.all.IMsngrContent.innerHTML= content;
}
function MsngrHTMLencode(strToCode)
{
strToCode = strToCode.replace(/</g,"&lt;");
strToCode = strToCode.replace(/>/g,"&gt;");
strToCode = strToCode.replace(/"/g,"&quot;");
return strToCode;
}
function MsngrIsUser()
{
return MsngrObj.IsUser(LocalUserEmail);
}var alphaChars = "abcdefghijklmnopqrstuvwxyz";
var digitChars = "0123456789";
var asciiChars = alphaChars + digitChars + "!\"#$%&'()*+,-./:;<=>?@[\]^_`{}~";
var folderID = "";
ie = document.all?1:0
ns4 = document.layers?1:0
dodiv=0;
function CallPaneHelp(topic_id,topic_displaystr) {
if (topic_id.indexOf(".htm")<0) {
bSearch=true;
H_KEY=topic_id;
L_H_TEXT=topic_displaystr;
} else { 
bSearch=false;
H_TOPIC=topic_id;
}
DoHelp();
}
function isASCII(str){
var v_len = str.length;
var i;
for (i = 0; i < v_len; i++)
{
if (asciiChars.indexOf(str.charAt(i)) == -1)
return false;
}
return true;
}
function ValidateEmail(str)
{
var ret = false;
if (typeof(str) != "undefined")
{
if (/^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$/.test(str))
{
ret = str;
}
}
return ret;
}
function ValidateLooseEmail(str){
var resultStr = str.replace(/ /gi, "");
var atIndex   = resultStr.indexOf("@");
var dotIndex  = resultStr.lastIndexOf(".");
if( resultStr == "" || !isASCII(resultStr) || dotIndex == -1)
return "";
if ( resultStr.lastIndexOf("@") != atIndex || resultStr.charAt(atIndex+1) == ".")
return "";
if ( atIndex <= 0 || dotIndex < atIndex ||  dotIndex >= resultStr.length-1)
return "";
return resultStr;
}
function ValidateDomain(str){
var resultStr = str.replace(/ /gi, "");
var atIndex   = resultStr.indexOf("@");
var dotIndex  = resultStr.lastIndexOf(".");
if( resultStr=="" || !isASCII(resultStr) || dotIndex == -1)
return "";
if ( atIndex > 0 || resultStr.charAt(atIndex+1) == "." || dotIndex >= resultStr.length-1 )
return "";
return resultStr.replace(/@/i, "");
}
function isEmail(str) {
var pass = 0;
if (window.RegExp) {
var tempStr = "a";
var tempReg = new RegExp(tempStr);
if (tempReg.test(tempStr)) pass = 1;
}
if (!pass)
return (str.indexOf(".") > 2) && (str.indexOf("@") > 0);
var r1 = new RegExp("(@.*@)|(\\.\\.)|(@\\.)|(^\\.)");
var r2 = new RegExp("^[a-zA-Z0-9\\.\\!\\#\\$\\%\\&\\'\\*\\+\\-\\/\\=\\?\\^\\_\\`\\{\\}\\~]*[a-zA-Z0-9\\!\\#\\$\\%\\&\\'\\*\\+\\-\\/\\=\\?\\^\\_\\`\\{\\}\\~]\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$");
return (!r1.test(str) && r2.test(str));
}
function CA(isOnload){
var trk=0;
for (var i=0;i<frm.elements.length;i++)
{
var e = frm.elements[i];
if ((e.name != 'allbox') && (e.type=='checkbox'))
{
if (isOnload != 1)
{
trk++;
e.checked = frm.allbox.checked;
if (frm.allbox.checked)
{
hL(e);
if ((folderID == "F000000005") && (ie) && (trk > 1))
frm.notbulkmail.disabled = true;
}
else
{
dL(e);
if ((folderID == "F000000005") && (ie))
frm.notbulkmail.disabled = false;
}
if (frm.nullbulkmail)
frm.nullbulkmail.disabled = frm.notbulkmail.disabled;
}
else
{
e.tabIndex = i;
if (folderID != "")
e.parentElement.parentElement.children[2].children[0].tabIndex = i;
if (e.checked)
{
hL(e);
}
else
{
dL(e);
}
}
}
}
}
function CCA(CB){
if (CB.checked)
hL(CB);
else
dL(CB);
var TB=TO=0;
for (var i=0;i<frm.elements.length;i++)
{
var e = frm.elements[i];
if ((e.name != 'allbox') && (e.type=='checkbox'))
{
TB++;
if (e.checked)
TO++;
}
}
if ((folderID == "F000000005") && (ie))
{
if (TO > 1)
document.all.notbulkmail.disabled = true;
else
document.all.notbulkmail.disabled = false;
if (document.all.nullbulkmail)
document.all.nullbulkmail.disabled = document.all.notbulkmail.disabled;
}
if (TO==TB)
frm.allbox.checked=true;
else
frm.allbox.checked=false;
}
function hL(E){
if (ie)
{
while (E.tagName!="TR")
{E=E.parentElement;}
}
else
{
while (E.tagName!="TR")
{E=E.parentNode;}
}
E.className = "H";
}
function dL(E){
if (ie)
{
while (E.tagName!="TR")
{E=E.parentElement;}
}
else
{
while (E.tagName!="TR")
{E=E.parentNode;}
}
E.className = "";
}
function doTabIndex(tbleColl)
{
if (tbleColl != null)
{
for (var z=0;z<tbleColl.length;z++)
{
if ((tbleColl.item(z).tagName=='A') || ((tbleColl.item(z).tagName=='INPUT') && (tbleColl.item(z).type!='hidden')) || (tbleColl.item(z).tagName=='SELECT'))
tbleColl.item(z).tabIndex=5;
}
}
}
function HMError(strEType,strError,strOther,strEN)
{
strError = unescape(strError).replace(/\+/g," ");
strError = strError.replace(/\\n/g,"\n");
switch(strEType)
{
case "A":
alert(strError);
break;
case "M":
if (ie)
DoModal(strOther,strEN);
else
DoFakeModal(strOther,strEN);
break;
case "C":
return(confirm(strError));
break;
}
}
function DoModal(strOther,strEN)
{
rv = window.showModalDialog("/cgi-bin/dasp/error_modalshell.asp?strEN="+strEN+"&r="+Math.round(Math.random()*1000000),"","dialogWidth:360px;dialogHeight:217px;help:0;scroll:0;status:0;");
if (rv.help)
CallPaneHelp(rv.help);
if (rv.url)
{
if (strOther=="attach")
DoSaveMSG();
else
location.href=rv.url;
}
}
function DoFakeModal(strOther,strEN)
{
ErrOther = strOther;
window.open("/cgi-bin/dasp/error_modalshell.asp?strEN="+strEN+"&r="+Math.round(Math.random()*1000000), "newwin", "resizable=no,width=360,height=217");
}
function MsngrExemptList()
{
var NotAddList = "crano|alert|msndirect|member_services|specialoffers_help|msn_newsletters|offershelp|viod|web_communities|staff|_esc_costarica|_esc_manila|_esc_mla_sykesredmond|_esc_policy_sunnyvale|_esc_tech_sunnyvale|abuse|Addressbook_Abuse|Addressbook_esc_Manila|Addressbook_Privacy|Addressbook_ts|bugreporter|datfix|DT_Hotmail|DT_MSN_Addressbook|HM_InternalEsc|hotmail_training|Hotmailprivacy|Hotmailprivacy_esc|microsoftcom_contactus|mmssupport|msncom_ca_en|msncom_us|passport|postmaster|premium_x|service_x|support_x|Hotmailprivacy|oebeta|sales|forgotpass|ycache|support|whatispp|invalidpass|changepass|createacct|formatadd|changeacct|yperson|prefopt|closeacct|whatismms|helpsend|howcheck|howfolder|saveprint|set_remind|addressbook|memdir|attachhelp|notifi|deletehelp|popmail|sighelp|mwdict|classify_2000|e_greetings|wchelp|hmoex|bulk_mail|hotmailalert|intruslog|whomd|acctsize|blockdom|speedissue|misdirect|mailprob|multsend|webtvhelp|maillisthelp|dnsprob|poperr|othererrors|abouthm|anti_viruses|adtag|teltext|helplink|frame_nframe|timestamp|homepagehelp|welchome|contactpartn|howsecure|commentquest";
var NotAdd = NotAddList.split("|");
MsngrExemptListObj = new Object();
for (var No=0; No<NotAdd.length; No++)
{
var Ad = NotAdd[No] + "@hotmail.com";
MsngrExemptListObj[Ad] = No;
}
}