# Activate Microsoft Office 365 ProPlus with this cool trick in JUST THREE steps | MANUAL Method

**Step 1: Remove your current trial license**

This step is optional if your trial license was expired. However, if it is still valid, you need to remove it. Because in some cases, after you activate your Office using KMS license, important features are resumed but the expiration notification still remains. Please follow instructions to uninstall the trial license.

Steps to remove your Office license

**Step 1:** Open command prompt as administrator

**Step 2:** Copy/run this command to determine what is the license key you want to remove.

`cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus`
If you see an error, try this command.

`cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /dstatus`

**Note:** “Office16” is codename of Office 2016. If you are using Office 2010/2013, replace “Office16” with “Office14” or “Office15”.

!![Office](https://i.imgur.com/BzRnpwo.png "Office")

**Step 3:** Copy and run these commands to remove the license. Note: replace “VMFTK” with the last 5 characters of your product key.

`cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /unpkey:VMFTK`

If you see an error, try this command.

`cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /unpkey:VMFTK`

![Done](https://i.imgur.com/mMjFG7z.png "Done")

Done!

**Step 2: Make sure your computer is ready**

You need to check your internet connection again and make sure that the Windows Update service is turned on. KMS license has to be verified by making a connection to my KMS server before it can be used. So you need to check if the KMS server is blocked or not. This is pretty simple. Just open your internet browser and try visitting this site: http://kms8.msguides.com. If it is visible, this means my KMS server is not blocked.

**Step 3: Activating your Office 365 using KMS client key**
Time needed: 1 minute.

Manually activate your Office 365 using KMS client key.

**1. Open command prompt as admin.**

First, you need to open command prompt with admin rights, then follow the instruction below step by step. Just copy/paste the commands and do not forget to hit Enter in order to execute them.

![CMD](https://i.imgur.com/lzqigbI.png "CMD")

**2. Navigate to your Office folder.**

If you install your Office in the ProgramFiles folder, the path will be “%ProgramFiles%\Microsoft Office\Office16” or “%ProgramFiles(x86)%\Microsoft Office\Office16”. It depends on the architecture of the Windows OS you are using. If you are not sure of this issue, don’t worry, just run both of the commands above. One of them will be not executed and an error message will be printed on the screen.

```
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```

![Office16](https://i.imgur.com/Yyhz9xR.png "Office16")

**3. Convert your Office license to volume one if possible.**

If your Office is got from Microsoft, this step is required. On the contrary, if you install Office from a Volume ISO file, this is optional so just skip it if you want.

`for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"`

Use KMS client key to activate your Office
Make sure your PC is connected to the internet, then run the following command.

```
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:kms8.msguides.com
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
```

Here is all the text you will get in the command prompt window.

```
C:\Windows\system32>cd /d %ProgramFiles%\Microsoft Office\Office16
 C:\Program Files\Microsoft Office\Office16>cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
 The system cannot find the path specified.
 C:\Program Files\Microsoft Office\Office16>for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Installing Office license: ..\root\licenses16\proplusvl_kms_client-ppd.xrm-ms
 Office license installed successfully.
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Installing Office license: ..\root\licenses16\proplusvl_kms_client-ul-oob.xrm-ms
 Office license installed successfully.
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Installing Office license: ..\root\licenses16\proplusvl_kms_client-ul.xrm-ms
 Office license installed successfully.
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:BTDRB >nul
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:KHGM9 >nul
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:CPQVG >nul
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /sethst:kms8.msguides.com
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Successfully applied setting.
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /setprt:1688
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Successfully applied setting.
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /act
 Microsoft (R) Windows Script Host Version 5.812
 Copyright (C) Microsoft Corporation. All rights reserved.
 ---Processing--------------------------
 Installed product key detected - attempting to activate the following product:
 SKU ID: d450596f-894d-49e0-966a-fd39ed4c4c64
 LICENSE NAME: Office 16, Office16ProPlusVL_KMS_Client edition
 LICENSE DESCRIPTION: Office 16, VOLUME_KMSCLIENT channel
 Last 5 characters of installed product key: WFG99
 
 
 ---Exiting-----------------------------
 C:\Program Files\Microsoft Office\Office16>
```

![successful](https://i.imgur.com/UEM1rKZ.png "successful")

If you would have any questions or concerns, please leave your comments. I would be glad to explain in more details. Thank you so much for all your feedback and support!





