**Want to Consolidate your vRealize Messages.**

Add this VBA to your mailbox to filter out the message that have been closed. 

Instructions for installation:

1.  Create a Folder called "vRealize Messages" under your Inbox
2.  Create a Folder called "vRealize Closed" under your Inbox
3.  Create a rule to push xx into "vRealize Messages", goto "Manage Rules and Alerts" and add a rule "apply this rule after the message arrives" from vROps@somedomain.name move it to the vRealizeMessages filder and stop processing more rules.
4.  Assign a Retention Policy (Assign Policy Button) to Remove Files from the closed folder after x number of days.
5.  Click on the Developer Button -> Macros -> Macros -> Edit.
6.  Paste/Upload the VBA code in git folder to the "ThisOutLookSession" Project (select this on the left Project selection box)
7.  File -> Save VBA Project

If you already have VBA messages in your in box you can run the Rule you defined above to move these to the message folder, you do this by clicking on the "Run Rules Now" button in the "Folder" ribbon. Once this is done you can send yourself an email to test this is working.
If you have a large amount of messages (I had over 3000) in can take a minute to run the first time and can lock up while working, this is ok. For all new messages you will not even notice it running.

If you wish to run it manually rather than automatically upon receipt of email then remove the "Application_NewMail" sub from the macro and saved the project.

The macro is configurable, so if the message subject or format of the body changes you can edit the settings. The vRealize messages have a unique AlertID tag which is used to identify the message and the closing message. A message must receive a closing message for it to be moved to the closed folder.


** Signing the Macro **

For the macro to autostart after a reboot you need to sign the macro. You can do this is you have a company wide certifcate that can be used or you can create a self signed cert and use that.

From windows; call
* C:\Program Files (x86)\Microsoft Office\root\Office16\selfcert.exe

Enter the name of your certificate - something meaninful is useful.
* Linux Automation - Office VBA Macros

Within outlook go back to the Macro/Visual Basic editor, on the Tools menu, click Digital Signature. Select a certificate or your self signed certificate and click OK.

For more information on signing visit: https://support.office.com/en-ie/article/digitally-sign-your-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01









