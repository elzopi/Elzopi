ReadMe for customized Outlook Today by Felix Reta, Epsitel Corporation, all rights reserved 2007-2021.

Loosely based from the original MSNBC outlook today that has been sunset.
Tested with Outlook 2003, Windows XP.

Files:
TC Outlook Today.reg
TC today.htm
TC today_o_ss.css (based from original msnbc.css)
i/ (directory contains specific custom images, logo and grandfathered images from original msnbc OL today as well)

External functionality and references:
http://www.clocklink.com/ (for clocks in the top portion of the dashboard)
RSSinclude.com (for RSS feeds)
kalsey.com (for syndicated weather)

Use at your own risk, comments are more than welcome but I can't promise I'll address your specific issues.

If you liked what I've done and works for you, consider a donation at my website at www.mexibeetle.net or www.epsitel.com. Or you can write to me at elzopi@mexibeetle.net or elzopi@msn.com

Installation:

Before you install, I strongly suggest for you to backup the following files:

Original Outlook Today HTML code, Close Outlook and then use Res://C:\Program Files\Microsoft Office\Office11\1033\Outlwvw.dll/outlook.htm via IE. You can check an article that describes it in more detail at: http://articles.techrepublic.com.com/5100-6346-5149522.html

Original registry entries:
[HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Outlook\Today]



Registry modification

You will have to modify the registry to show the new outlook today page and customize the folder list you'd like to see in the new page. If you feel uncomfortable on doing this then this customization is not for you (sorry).

Following is a description of the registry keys that will need to be added, please note that the version (in this case 11.0) depends on the Office version you have available in your computer.

[HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Outlook\Today]
Add UserDefinedUrl to your location.

[HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Outlook\Today\Folders]
Add the folders you'd wish to show in the new Outlook today page. Format is as follows
"0"="\\\\Mailbox - LastName, FirstName\\Inbox"
"1"="\\\\Mailbox - LastName, FirstName\\Drafts"
"2"="\\\\Mailbox - LastName, FirstName\\Outbox"
"3"="\\\\Mailbox - LastName, FirstName\\Notes"

Check your Mailbox in Outlook for the right folder names.

If you'd like to add your own company or personal logo, substitute a suitable .gif image in the i/ directory and name it as "My_logo.gif", It will be resized to width=345 height=50 pixels.

In this particular implementation, I used syndication for Dilbert daily cartoon, I had to tweak the html code so it would appear at the end of the page so it would fit correctly in Outlook today's real estate, apparently the cartoon has a fixed width that I couldn't figure out how to dynamically resize.

There are additional warnings/documentation/comments embedded in the html code of TC Today, make sure you browse it to make your customization less painful.

Enjoy!
