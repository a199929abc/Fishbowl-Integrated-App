## Overview
FishPro (Fishbowl integrated App) is an application that reduces MTC staff repetitive work processes, it can generate the recovery template excel sheet base on the inventory number in the fishbowl system.

### Input Section：
As figure 1 shown, the only input the application needs is a **column of Inventory #** , be careful the header row format is fixed, DO NOT modify it. Once you fill in all the inventory # , save and close the document. you are ready to go.

Figure 1. iupt template
![alt text](/picture/input.JPG)

### Login Section:
After you click the .exe file. You will see a login page. Please use the username and password at the MTC fishbowl system. 

Figure 2. login
![alt text](/picture/login.JPG)

### File read section:
Once you login to the application, you will see a new window shown in figure 3, click **Explore** button to find the .xlsx file you prepared in the input section. After that, you will see the file path shown on the panel. Click **Enter** to run the application. If you accidentally choose the wrong file, you can simply Explore again, the new file path will be shown.


Figure 3. Choose file

![alt text](/picture/chooseFile.JPG)


### Process section: 
If you see a pop-up window shown in figure 4, that means you successfully create a recovery template. The output excel file will normally at the users' root directory. The output filename format: **out_YOUINPUTFILENAME.xlsx** if you can’t find it for some reason. Please search the format name on your computer. 

Figure 4. process complete
![alt text](/picture/processcomplete.JPG)

### Output section: 
After the process has completed, you will receive a new excel file as shown in figure 5. The program automatically fills the excel sheet based on the inventory provided. The Instrument name, serial number and device number getting from the fishbowl inventory system.

Figure 5. output 

![alt text](/picture/outputsample.JPG)

## Description Area:
Description area is the only place user need to pay extra attention. Because there are some custom fields that needs to be manually filled based on various situations and tasks. Such as **"METHOD","ATTENTION TO","RECEIVED", "NOTES", "LINK TO LOG"** etc. **"RECEIVED DATE"** automatically set to current date.

Figure 6. description

![alt text](/picture/description.JPG)

# QA section:
### 1. Why I can't login ? 
There might be two possible error lead to this problem, 1) your username and password are wrong, please double-check and enter again. 2) Please check your Internet and retry.

### 2. Get pop-up error message "Check row number 8 then retry"
This happens because the description area in the inventory system doesn’t format right. There are some possible way to fix it.
1) Open fishbowl API ->search the part number ->check the description area whether format as :”SN: XXX, DI:XXXXX” You can modify the description in the fishbowl system and retry.

Figure 7. Error

![alt text](/picture/error.JPG)

### Format Issue
Correct format :
1. SN and DI exist. SN is in the front of DI :  "**EtherTek RMS-300 Monitor, SN: 00-50-C2-53-89-AC, DI: 24368**"
2. SN and DI exist. DI is in the front of SN :  "**EtherTek RMS-300 Monitor, DI: 24368 SN: 00-50-C2-53-89-AC**"
3. ONLY DI exist : "**EtherTek RMS-300 Monitor, DI: 24368**"
4. ONLY SN exist : "**EtherTek RMS-300 Monitor, SN: 00-50-C2-53-89-AC**"
5. DI and SN don't exist : "**EtherTek RMS-300 Monitor**"
6. (Update) SN exist DI is pending : "**EtherTek RMS-300 Monitor, SN: 00-50-C2-53-89-AC DI: Pending**"  "pending" is Capital insensitive can be any format


Error might show up to case below: 
1. "EtherTek RMS-300 Monitor, SN 00-50-C2-53-89-AC, DI: 24368" Reason : "SN:" tag missing ":" can be add to fishbowl client
2. "EtherTek RMS-300 Monitor, SN: 00-50-C2-53-89-AC, DI 24368" Reason : "DI:" tag missing ":" can be add to fishbowl client
3. "EtherTek RMS-300 Monitor, SN: 00-50-C2-53-89-AC, DI: " Reason : After "DI: " tag missing value. If you don't know the DI, add "Pending" 
4. More cases can be discovered in the Fishbowl client. If you are not sure where is the mistake, FOLLOW THE CORRECT format and modify the description


