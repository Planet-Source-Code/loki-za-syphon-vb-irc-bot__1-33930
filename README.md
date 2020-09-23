<div align="center">

## Syphon \- VB IRC Bot


</div>

### Description

This is a fully functional IRC Bot written in VB6. There are lots of other applications I'm building into it, but I thought I'd release this stable BETA so long. The applications to come are: 1) Statistical logging to a SQL database from which Analytical reports can be drawn. 2) Dictionary driven events. 3) Owners/admins list for stopping normal users abusing the bot (this section is actually almost complete in the version i'm working on now)

<edit>

I've included these instructions for operation:

Inside the /media/ subfolder, there's a file called "settings.conf" .. inside there, change the bot's connection details to the server you want it to connect to. Run the program (debug or compiled) and the bot will connect to that server. You can view the bot's progress in the /media/logs/ folder, under the log file that matches today's date. Once the bot has connected successfully, you can then control it through IRC. Commands are sent like this: /msg <bot> .<command> <parameters>

example:

/msg MyBot .join #bots

When doing commands, if you get the parameters wrong, the bot will give you a syntax breakdown..

Here's all the commands so far:

.join <channel>

.part <channel>

.say <nick/channel> <message>

.notice <nick/channel> <message>

.action <nick/channel> <action>

.quit <quit message>

.kick <channel> <nick> <message>

.ban <channel> <nick>

.mode <channel> <mode type> <mode>

(eg. .mode #bots +b <nick>)

There should be lots more coming in ver.2 :)

</edit>

So tell me what you think? And remember, vote lots!!! ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-04-18 16:12:06
**By**             |[LoKi\-ZA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/loki-za.md)
**Level**          |Advanced
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Syphon\_\-\_V734574192002\.zip](https://github.com/Planet-Source-Code/loki-za-syphon-vb-irc-bot__1-33930/archive/master.zip)








