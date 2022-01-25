=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
BSetup! (Package & Deployment Setup Enhanced Version)
Copyright (C)2002 Behrooz Sangani	<bs20014@yahoo.com>
				http://www.geocities.com/bs20014/
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
Setup program is and will be property of Microsoft(R). This copyright
refers to enhancements made to the application.
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

Absolutely freeware under one condition: Credits must remain intact!



=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	How to use this project?
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

1. Run the project. Compile it to get the SETUP2.exe file. This file uses
    encryption methods that may be simply changed for every application
    in order to avoid key matches in two different programs. You may want to
    modify and use your own formula in this section before compiling the project.
   This is explained in the source code.

2. Copy the SETUP2.exe to
    	%VB Directory%\Wizards\PDWizard

3. Every time you want to use Package and Deployment Wizard with BSetup!
    features just rename the SETUP1.exe file in the directory mentioned in No. 2
    to something else (For example SETUP000.exe) and rename our version named
    SETUP2.exe to SETUP1.exe
    Next time you run Package and Deployment Wizard it uses this version and will
   ask for Registration Key on setup.


=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	Registration Method
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

BSetup! uses a registration method that generates unique IDs for every computer.
This means that the chance of using a certain key on two computers is 
nearly impossible. BSetup! gets the Hard Disk Serial Number and generates
a customizable ID from it. The end user is asked to provide this ID and his name
to get a registration key.
The registration information is saved in the registry under the application title so 
that the installed application would be able to use it. 
The encryption process and seed valueis fully customizable and may be changed 
for each application. However, this is not much necessary if you are not a 
professional coder because the encryption process included in the source 
generates a very long string which will take ages to crack.


=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	Files added to setup1 project
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

The original setup project can be found at: 
	%VB Directory%\Wizards\PDWizard\Setup1

The following files were added to setup project:
	Form(s):
	   frmSerialCheck		Main form that appears before all forms
	Module(s):
	  ModHard		Functions & APIs to get HD Serial
	Classe(s):
	  clsValidator		Key Validation Class
	  CMD5 *			MD5 Class
	  COwnerRegistration *	Registration Key Maker
  	Graphic(s):
	  hand_icon.cur		Hand Icon :)
	
 * Adopted from the web as described in source code



=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	KeyGen Project & Registered Users Database
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

You need to generate registration keys for the end users. Include a registration 
form in your packed application and ask users to provide their PC ID. You can 
then generate their unique reg keys and send it back to them so that they could 
install the software. This is the best method for shareware applications because
it reduces the chance to crack the program. 

KeyGen project can easily save the registered users to an access database.
You have access to all registered users information in this database. 



=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	Important Notes
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

* User database password is: BSetup!

* Whenever you change the encryption algorithm used for encrypting PC ID you
  have to change it in the KeyGen project.

* You cannot run SETUP1 project. You have to compile it and use it in a packed
  application to see the results.

* Please leave the credit line in the SETUP1 project as it appears and don't remove
  the reference to my website. This is the only thing I ask for.



=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=
	Disclaimer
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-==-=-=-=-=-=-=-=

THE PRODUCT IS DISTRIBUTED "AS IS". NO WARRANTY OF ANY KIND 
IS EXPRESSED OR IMPLIED. YOU USE THE PRODUCT AT YOUR OWN RISK. 
THE DEVELOPER OF THE PRODUCT WILL NOT BE LIABLE FOR DATA LOSS, 
DAMAGES, LOSS OF PROFITS OR ANY OTHER KIND OF LOSS WHILE USING 
OR MISUSING THIS SOFTWARE.
