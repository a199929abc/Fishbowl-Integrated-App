## Overview
FishPro (Fishbowl integrated App) is an application that reduces MTC staff repetitive work processes, it can generate the recovery template excel sheet base on the inventory number in the fishbowl system.

## Input Section：
As figure 1 shown, the only input the application needs is a **column of Inventory #** , be careful the header row format is fixed, DO NOT modify it. Once you fill in all the inventory # , save and close the document. you are ready to go.

## Login Section:
After you click the .exe file. You will see a login page. Please use the username and password at the MTC fishbowl system. 

## File read section:
Once you login to the application, you will see a new window shown in figure 3, click **Explore** button to find the .xlsx file you prepared in the input section. After that, you will see the file path shown on the panel. Click **Enter** to run the application. If you accidentally choose the wrong file, you can simply Explore again, the new file path will be shown.

## Process section: 
If you see a pop-up window shown in figure 4, that means you successfully create a recovery template. The output excel file will normally at the users' root directory. The output filename format: **out_YOUINPUTFILENAME.xlsx** if you can’t find it for some reason. Please search the format name on your computer. 

## Output section: 
After the process has completed, you will receive a new excel file as shown in figure 5. The program automatically fills the excel sheet based on the inventory provided. The Instrument name, serial number and device number getting from the fishbowl inventory system.

###Description Area: 
