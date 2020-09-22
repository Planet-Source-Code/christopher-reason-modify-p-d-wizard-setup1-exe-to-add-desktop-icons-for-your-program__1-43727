<div align="center">

## Modify P&D Wizard/Setup1\.exe to add Desktop Icons for your Program\!


</div>

### Description

How to modify P&D Wizard/Setup1.exe to add Desktop

Icons during the installation of your Visual Basic Programs.

PLEASE, TAKE AN EXTRA MINUTE TO VOTE IF YOU LIKE THIS ARTICLE.

Thanks, Christopher.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-03-04 08:50:58
**By**             |[Christopher Reason](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/christopher-reason.md)
**Level**          |Intermediate
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Modify\_P&D155415342003\.zip](https://github.com/Planet-Source-Code/christopher-reason-modify-p-d-wizard-setup1-exe-to-add-desktop-icons-for-your-program__1-43727/archive/master.zip)





### Source Code

If you find this article useful, please be sure to come back and vote for it! I would greatly appreciate that, Christopher.
________________________________
<p>    In reading the article titled "Make P&D
Wizard/Setup1.exe create Desktop shortcuts" (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29890&lngWId=1) written by W. Baldwin on
12/17/2001 here at PSC, I found that my version of Setup1.vbp (Visual Studio)
was different that the one mentioned in that article. I found that the lines of
code that were mentioned did not match up with what I was looking at myself. I
found the answers by carefully studying the code and wanted to provide these
answers with the rest of you. Although my article will show different code
lines, the concept is the same and I do recommend that you read the before
mentioned article along with this article to get a full understanding of what
you are doing and why you are doing it. W. Baldwin's article has many links to
other articles on MSDN that support and explain the modification of Setup1.vbp.
I wanted to post these differences for anyone that may have the same version as
I do and could not make the necessary modifications to make this change work.</p>
<p>    As always, <b><u>you should back up your Setup1.vbp files
for safety!</u></b> Once you have backed up the files, follow these steps to get
your applications icons on your users desktop with ease. Note that "<font color="#800080">Red</font>"
text is the original text and "<font color="#008000">Green</font>"
text is the modified text.</p>
<ol>
 <li>Open Setup1.vbp, go to frmSetup1 in the Form_Load Event and look for the
 following lines of code:<br>
 <font color="#800080">If (GetGroup(gsICONGROUP, iLoop) = gsSTARTMENUKEY) Or (GetGroup(gsICONGROUP, iLoop) =
 gsPROGMENUKEY) Then<br>
 <br>
 </font>Change the line to read as follows:<br>
 <font color="#008000">If (GetGroup(gsICONGROUP, iLoop) = gsSTARTMENUKEY) Or (GetGroup(gsICONGROUP, iLoop) = gsPROGMENUKEY) Or (GetGroup(gsICONGROUP, iLoop) = "DESKTOP") Then</font></li>
 <li>Now go to basSetup1 in the sub "CreateIcons" and look for the
 following line of code (near the bottom of the sub):<br>
 <font color="#800080">Loop</font><br>
 <font color="#800080">CreateOSLink frmSetup1, strGroup, strProgramPath, strProgramArgs, strProgramIconTitle, fPrivate, sParent<br>
 <br>
 </font>We are going to ADD code <b><u>between the Loop end and the call to
 CreateOSLink</u></b> to say the following:<br>
 <font color="#800080">Loop</font><br>
 <font color="#008000"> If UCase(strGroup) = "DESKTOP" Then<br>
 strGroup = "..\..\Desktop"<br>
 End If</font><font color="#008080"><br>
 </font><font color="#800080">CreateOSLink frmSetup1, strGroup,
 strProgramPath, strProgramArgs, strProgramIconTitle, fPrivate, sParent</font></li>
 <li>We are now done modifying Setup1.vbp, save your work, and make
 "Setup1.exe". This new exe should reside in the folder
 "Program Files/Microsoft Visual Studio/VB98/Wizards/PDWizard" (of
 course, your program name will vary by the version and type of VB you are
 running. But the sub folders will be the same.</li>
 <li>Now we must modify the Setup1.lst file, this file can be opened with any
 text editor, I use NotePad to work on mine. The modifications from here on
 are the same as the previous article mentioned. After you Package/Deploy
 your program, open the file "Setup1.lst" and go down to the [IconGroups]
 section of the file. You will find the following text for a program that has
 just one icon that will be placed in the "Programs" folder of the
 Start Menu:<br>
 <font color="#800080">[IconGroups]<br>
 Group0=Scheduler<br>
 PrivateGroup0=-1<br>
 Parent0=$(Programs)<br>
 <br>
 [Scheduler]<br>
 Icon1="Scheduler.exe"<br>
 Title1=Scheduler<br>
 StartIn1=$(AppPath)<br>
 <br>
 </font>This text sets an Icon for a program called "Scheduler" in
 the Start Menu under "Programs". Under [IconGroups] there is just
 one group and under [Scheduler] you see the information for that group.<br>
 </li>
 <li>Now we are going to add a second group to describe the Desktop icon,
 modify Setup1.lst to read as follows:<br>
 <br>
 <font color="#800080">[IconGroups]<br>
 Group0=Scheduler<br>
 PrivateGroup0=-1<br>
 Parent0=$(Programs)<br>
 </font><br>
 <font color="#008000">Group1=DESKTOP<br>
 PrivateGroup1=-1<br>
 Parent1=$(Programs)<br>
 </font><br>
 <font color="#800080">[Scheduler]<br>
 Icon1="Scheduler.exe"<br>
 Title1=Scheduler<br>
 StartIn1=$(AppPath)<br>
 </font><br>
 <font color="#008000">[DESKTOP]<br>
 Icon1="Scheduler.exe"<br>
 Title1=Scheduler<br>
 StartIn1=$(AppPath)<br>
 <br>
 </font>As you can see, we added another group, "Group1" and called
 it "DESKTOP", then below that we added a description for the
 "DESKTOP" icon to show that the icon itself is the icon for the
 Program's Exe file. The title is Scheduler and the "StartIn" path.
 <b>Note that each group is numbered, starting with Group0, then Group1</b>,
 if we were to add more icons they would be named Group2, Group3 and so on. A
 very good explanation of this is in the previous article I mentioned at the
 top of this tutorial.</li>
</ol>
<p>    I hope that this article will help those that had
questions and just couldn't find the right answer's out there anywhere.<font color="#008000"><br>
</font></p>

