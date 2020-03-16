<%APPLICATION.LOCK 
sSessID = Session.SessionID <!--Returns a unique id for each visitor-->
sSessURL= Request.ServerVariables ("URL") <!--Returns page user is currently viewing-->
sQString= Request.ServerVariables ("QUERY_STRING") <!--If there is a querystring, include that as well-->
if sQString <>"" then
sSessURL=sSessURL&"?"&sQString
end if
Application("yoyo"&sSessID) = "<a href='" & sSessURL & "'>" & sSessURL &"</a></td><td>" & Session("Start")<!--Throw it all together with a "yoyo" so that you can later call all variables that contain "yoyo" and you would get the right ones-->
APPLICATION.UNLOCK
%>