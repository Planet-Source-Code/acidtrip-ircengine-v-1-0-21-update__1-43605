|--------------------------------|
	Project Details
|--------------------------------|
Version: 1.0.22
Date Released: 02/20/2003
Author: aCiDtRip (Raffy Ibasco)
VB Version: 6
Includes: Csocket Class, Sample Project
You can contact me at: email: acidtrip@most-wanted.com
		       website: www33.brinkster.com/constantanxiety/
----------------
Description: 
----------------
	Just another IRC parser, but compiled as a class module...ISN'T THAT NEAT!?!?!? It has a "lot" of events for you to use, there are 34 events actually...9 methods and 3 properties... anyway, you are VERY free to mess around with the code, change it by all means! Haven't tried it on any server online like DALnet or Undernet, but don't worry this is based from the RFC document 1459 and 2812...Its very "well" commented...(I think)

--------------
How to use:
--------------

	* Remember to include the "WithEvents" when you are delcaring an object variable.
	
		For example: Dim WithEvents irc as IRCengine
	
	* When recieving data from a socket, use the ProcessData Method, or else the whole class module would 		be totally useless ;) 

Properties:

	>> BlockCmdEvents (Boolean) - If set to true, it will ignore all of the Command Processes of the Parser
				      only the OnCommand event would work.
	>> BlockNumEvents (Boolean) - Also the same as the BlockCmdEvents but this time, its the Numeric events
				      that are being ignored. Only the OnServerNumeric event would work.
	>> unixtime (Long) - this property returns the current number of seconds that has elapsed since 
			     January 01, 1970 (Unix Format)
	
Methods:
	
	>> CTCP (String) - Returns the specified CTCP Format
	>> GetLocalTZ (Long) - Returns the Local TimeZone
	>> GetToken (String) - A tokenizer (specially made by me ;>) Returns the delimited string
	>> IP2Long (Long) - Converts an IP Address to a Long integer (This is not entierly accurate...)
	>> Long2IP (String) - Converts from a Long Integer to an IP Address
	>> Notice (String) - Returns the specified Notice Command format
	>> Privmsg (String) - Ummm...same as how the Notice method works
	>> ProcessData - This is the root of everything...if not Implemented, the whole damn thing wont work ;)
	>> sUnixDate (String) - Returns the formatted date based from the specified Value

------------------
Notes:
------------------
	Ok..this is actually my first time ever to release a vb code on the net, I mean I have tons of codes dumped here in my harddrive but never let them out (not even once), so I guess its about time for me to do some sharing huh? Well...I'm not guaranteeing you a HUNDRED percent of this code's functionality, well at least it worked for me...Anyway I need your comments and suggestions, if there are any bugs please report it a.s.a.p, if you think there is a better process then tell me, don't be afraid cause I don't bite!!, I keep an open mind to any ideas...
