
My goal with this is you can see many ways in what Autohotkey can help and get some ideas and insights of work you can do

Some of the works I've done for people include the following:



# Categorize types of files into specific folders by metadata and extension
**Goal**\
Move specific types of file to a specific folder when they are added to a folder.\
**Overall Process**\
The script uses the "WatchFolder" library to keep track of the new added files in a folder. In the real-application, the folder targeted was the Downloads folder in the user PC. Note that this library works best then there is a single file added, sometimes it have troubles when multiple files are added at the same time, but not always.\
When a new file is added to the target folder, it start a process to categorize it. It uses a function to retrieve the metadata values and narrowing the result to get the "Kind" metadata value, which is the one used to differentiate the type of file (Pictures, Music, Video). 
~~~
attributes_selection := [11] ; 11, Kind
file_metadata := GetDetailsOf_Targeted_Data(new_item_path, attributes_selection)

if (file_metadata["Kind"] = "Picture") {
    pictures_folder := root_folder "\Pictures\"                
    if (!Check_Folder(root_folder, "Pictures")) {
        FileCreateDir, %pictures_folder%
    }
    FileMove, %new_item_path%,  %pictures_folder%
    Check_Duplicate(ErrorLevel)
}

else if (file_metadata["Kind"] = "Music") {
    music_folder := root_folder "\Music\"
    if (!Check_Folder(root_folder, "Music")) {
        FileCreateDir, %music_folder%
    }
    Musics_folder := root_folder "\Music\"
    FileMove, %new_item_path%,  %music_folder%
    Check_Duplicate(ErrorLevel)
}   
[...]
~~~

But also, the script can categorize the file regarding their extension, like this:
~~~
if (OutExtension = "doc" || OutExtension = "docx") {
    Word_folder := root_folder "\Word\"
    if (!Check_Folder(root_folder, "Word")) {
        FileCreateDir, %Word_folder%
    }
    FileMove, %new_item_path%,  %Word_folder%
    Check_Duplicate(ErrorLevel)
}
else if (OutExtension = "psd") {
    Photoshop_folder := root_folder "\Photoshop\"
    if (!Check_Folder(root_folder, "Photoshop")) {
        FileCreateDir, %Photoshop_folder%
    }
    FileMove, %new_item_path%,  %Photoshop_folder%
    Check_Duplicate(ErrorLevel)
}
else if (OutExtension = "rar") {
    Rar_folder := root_folder "\Rar\"
    if (!Check_Folder(root_folder, "Rar")) {
        FileCreateDir, %Rar_folder%
    }
    FileMove, %new_item_path%,  %Rar_folder%
    Check_Duplicate(ErrorLevel)
}
[...]
~~~

This scripts provides a great folder and file structure, and it can be modified to categorize by other means, like using Regular Expressions to match strings in the filename, size, video lenght, video framerate, image resolution or other metadata propieties



# Starring Google Maps places
**Goal**\
Mark and verify all the links in the CSV as a Starred place

**Overall Process**\
Looping through a CSV which contains a link for specific places in Google Maps.\
By using the FindText library, it is searched for 5 key elements to complete the process to mark the palce on the link as a Starred Place.\
There was included a fail-safe to check if the place was already marked or not, if so, continue with the next link and stop until all the links in the CSV are opened and marked.


# New Outlook email fill with Word Template
**Goal**\
To quickly select a template from different Word documents and put it on a new Outlook email

**Overall Process**\
First there is a verification to run Outlook as Administrator to avoid some possible errors of privileges while using COM. In the user desktop, there will be different shortcuts that each one of them point to a specific Word document.\
Where the user selects one shortcut, it can do it normally or double-clicking it while holding the Shift key to open the document to modify it instead of launching the main process. The main process first connect to Word through COM, open the corresponding document, copy all of the contents and close the document.\
Then it connects to Outlook through COM and by using the command
~~~
Outlook.Application.CreateItem(0)
~~~
And this to launch a new email window
~~~ 
Email.Display
~~~
Then it navigates to focus on the body area and paste the contents of the template there.\
There is an additional script that validates the shortcuts for the templates, by having a predefined folder where the Word documents are, then it verify if the shortcut already exists on the desktop, if not, it creates a new one.\
To differ the target document to be opened, in the shortcuts, in the Target field, there is passed an argument that serves as differentiator where the main script looks for its value and execute the main process.


# Display a message after 7 days as a expiration date measure
**Goal**\
Keep track of the days after the first execution of the script and display a message

**Overall Process**\
This is a more an add-on to a main script, where the function "TriggerTrailStatus()" is called when the user wants to start or check the trail of the software.\
The scripts uses an INI file to keep track of the date values. When this function is called, it will eaither create the default INI if it doesn't exist or check for the trial status.
~~~
TriggerTrailStatus() {
    global

    INITrialTrackingVariables()
    
    if (!FileExist(ini_full_path)) { ; If the INI doesn't exist, it means that it's the first run. So create the INI
        CreateDefaultINI()
    }
    else {
        CheckForTrialStatus()
    }   
}
~~~

If the INI doesn't exist, it will be created one from scratch within the code, like this

~~~
CreateDefaultINI(){
    global

    first_day_run := A_YDay ; Get the file
    target_trail_end_day := first_day_run + 7 ; Calculating the end day. It will be 7 days ahead the first day

    INI_Structure = ; Default structure of the INI
    (Ltrim
        [TrialDayTracking]
        first_day = %first_day_run%
        end_trail_day = %target_trail_end_day%
    )

    Transform, INI_Final_Structure, Deref, %INI_Structure% ; Set the value from "first_day_run" within the string. This to optimize the code and ommit a further INIRead and INIWrite. Same for "target_trail_end_day"

    FileDelete, % ini_full_path ; this is just a safebelt, in case that there is any problem to locate the ini. This will delete the ini file in order to start fresh
    FileAppend, % INI_Structure, % ini_full_path ; Load the structure to the INI file
    FileSetAttrib, ^H, %ini_full_path% ; Turn it off right after creation. "^" means Toggle, and because after creation the file is always visible (No Hidden), toggle the state from "On" to "Off"
}
~~~

If the INI file exists, it will read the values from it and compare if the days of trial is reached.\
~~~
CheckForTrialStatus(){
    global

    IniRead, end_trial_day_value, %ini_full_path%, TrialDayTracking, end_trail_day ; Reading the INI and retrieving the end day of the trial
    current_day := A_YDay ; Get the current day

    if (current_day > end_trial_day_value) { ; Check if the current day is bigger than the end trial date
        ; Action if the script is 7 days older
        MsgBox %  "Trial Expired"  ; Alert to show the user its trial ended

        FileDelete, %ini_full_path% ; This will delete the current INI if the end day is reached

        SplitPath, A_ScriptFullPath,, script_folder ; This retrieve the folder where the script is to the "script_folder" variable
        FileRemoveDir, %script_folder%, 1 ; This will delete the folder where the script is located. The "1" is for deleting subfolders (and even if the folder doesn't contains subfolders, keep it, sometimes without the "1" it doesn't work properly)
    }
    else {
        ; Action if the script is less than 7 days older
        MsgBox %  "You have " end_trial_day_value - current_day " days left before the trial ends."  ; This will display the days left until it reaches the end day
    }    
}
~~~

The objective is to mimic a expiration date security measure of a script in order to avoid its usage after 7 days.


# Find matches in a web table column
**Goal**\
Find values of a table column to know if there is a match or not with values on a text file and create an Excel sheet with the results
**Overall Process**\
There is a text file that contains numbers in each line. Those numbers are the ones we want to keep track on a web table column to see if they are there (this means they are unavailable) or not (if not, they are available), they represent packages that are being transported through specific doors and the values change every minute.\
Here it is used Selenium and Chrome. Like most of web scrapping development, first the values and correct pointers need to be discovered by going into the Console panel on Chrome (I prefer using Xpath over any other method to locate web elements; to me it's much more robust and reliable).\
After the XPaths are located, it is used this command 
~~~
row_total := oChrome.findElementsByXpath("//*[@id='dashboard']/tbody/tr").Count()
~~~
to get the total of rows on table to do a proper Loop with the same value (So we don't over or under loop).\
There is a process to validate the data with RegEx, since sometimes the number is accompanied by letters, so the bare numbers are extracted. After that there is a comparison of that value with the values on the text file, and there is where we determine if the value (which represents the door number in the column) is available or not.\
It finishes the looping through all the rows and we get the exact values that we want, then it's time to create the Excel file.\
First there is deleted all of the previous files created because only the latest is relevant. There is a process to name the file using the date and time of the creation. \
Now, a nested array is filled with the final values, then we connect to Excel through COM and create a new Worksheet with 
~~~
XL_Worksheet := XL_O.Workbooks.Add()
~~~
And give it a name to the sheet with 
~~~
XL_Worksheet.Worksheets(1).name := sheet_name
~~~
After that with the use of a For-Loop we set the values on the Excel sheet and give it a table format with
~~~
XL_O.ActiveSheet.ListObjects.Add(1, XL_O.Range("$A$1:$A$" inset_data_array.Count() + 1),_, 1).Name := "Available_Doors_Table"" 
~~~
And set a style with
~~~
XL_O.ActiveSheet.ListObjects("Available_Doors_Table").TableStyle := "TableStyleDark6"
~~~
Finally, the Excel file is saved and the file is ready to open, now the process will repeat itself each minute, with the use of a Timer



# Restrict the use of apps by time per day and a random maximum limit number
**Goal**\
Keep track of the time an app or website is being used, use an INI file to store the values and when reach a random maximum number, close the app and reset the counters on the next day
**Overall Process**\
For this, it was used a nested array that will contain the data of the target apps and websites (like ID, WinTitle, Counters, Limit).\
The array was designed so it can loop through all the values for each app and monitor the time values in the most reduced way.\
When the script starts it will launch a Timer with 1-second interval. It will then use a For-Loop using the main array, inside it, first it reads and verifies that the stored main date is different from the current date. If they are different, it means the day changed, and then it will reset the counters and reach limit.\
After that it will extract the counter values depending on the current app ID.\
Now it will do the actual calculation of the time used by the program by using the "WinExists" command. If the windows does exist, it will be added an increment of 1, that represents the seconds used by the app.\
After the counter was increased or not (in case the app wasn't being used), finally, it will verify if the current counter is greater or equal to the reach limich.\
Note that reach time is set, for practical purposes, on integers (30, 60...) then converted into seconds by multiplying them by 60 and it is set a range from X to X, separated by a "|" symbol, so later it can be separated with the "StrSplit" function and use the Random command to get a random number between that range.\
If the result is positive, it will then close the app and set the "limit_reach" value to 1 and store back at the INI file


# Use Neutron.ahk to create a UI to send data triggered by hotkeys to a Firebase database
**Goal**\
Login with user credentials in a Firebase database within a Neutron window and depending hotkeys triggered, send data to specific parts of the database by using REST API's.\
**Notes**.\
* There was used a custom version of "localStorage" to be able to pass data between pages
* A Neutron function was modified to use the "localStorage" easily
* The user must be logged in order to be able to send the data with the hotkeys
* This is a desktop complementation for a website called Callouts Evolved. Website: https://www.calloutsevolved.com/. This is a gaming team communication enhacement tool\

**Overall Process**\
Neutron.ahk was used to create a login and a main page (that shows after login). There was done a Firebase setup to be able to conenct with the cloud database and validate the user credentials. After they are validated, the main page is displayed and now the user can use the various hotkeys to send specific data to Firebase.\
The hotkeys are set to send strings like "Attack", "Defend", "Retreat" and others and are sent to the session where the user is currently connected to. To know the session, the UID that provides Firebase was used.\
When the script starts, there is a creation of labeled Hotkeys, by using the "Hotkey" command, this is to enable flexibility to change the hotkeys assignements by the user fast without entering to the code.\
Then it starts different arrays, each one with the commands they send to a specific game, then all of those arrays are merged into a master nested array and after that the API Endpoints are set into variables, so they are used easily. Finally, the Neutron window starts, displaying the login page.\
Regarding the REST API process, each hotkey is assigned to launch a Function with a predefined argument, that will later be used to know what value to send. This is an example of how it looks:
 ~~~
Hotkey_4:
Main_Connection_API("Numpad4")
return

Hotkey_5:
Main_Connection_API("Numpad5")
return

Hotkey_6:
Main_Connection_API("Numpad6")
return
 ~~~ 
When those hotkeys are triggered, first the function verify if the user is logged in or not, by calling another function called "Get_User_Data" that uses the modified Neutron function to get the data passed from the login page to the main page that contains an array of data of the user (Email, UID, Logged Status, etc) and then it verifies if the Logged status is set to 0 (not logged in) or 1 (logged in).\
If the user is logged in, then it will continue. Now it get both the target session and the target game by using the UID that is received when the function before made use of the "localStorage" functionality, like below:
 ~~~
session_id_pointer := [Get_User_Target_Session(user_data_stored.uid,"session")] ; This is done to be able to set a key in the array as a variable instead of a fixed value
target_game := Get_User_Target_Session(user_data_stored.uid, "enc")
 ~~~

Also, it uses the main array to get the target action that is send depending on the hotkey that triggered the function. Like so:
 ~~~
action_sent := main_game_actions[target_game, hotkey_ID]
 ~~~

After that, it creates a JSON-like array that will mimic a JSON that will be sent to the Firebase database, as such:

 ~~~
session_Object := { session_id_pointer[1] : { btnId : action_sent 
		, by: {displayName :user_data_stored.displayName
		, email: user_data_stored.email
		, uid: user_data_stored.uid 
		, date: CurrentDate}}}
 ~~~
 
Finally, that array is send as the "Body" in a REST API call by using the method "PATCH" to update the data, like below:
 ~~~
Body := JSON.Dump(session_Object,, "`t")
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("PATCH", db_btnData ".json", false)
oWhr.SetRequestHeader("Content-Type", "application/json")
oWhr.Send(Body)
 ~~~
The update on the cloud database is almost instantaneity, from 1 to 3 seconds maximum


# Track new circles coming from the right in a graph
**Goal**\
Trigger multiples clicks after a new circle appears from the left.\

**Overall Process**\
In a graph, regarding variable values, there are circles that appear over time and are going few pixels to the right as time passes. They indicate a specific value the user wants to keep track of.\
The script uses the FindText library to target the circle reference as an image.\
First, it is verified if Google Chrome is running, if so, there is a calculation of the area that the script will work with regarding the proportions of the Google Chrome window and then doing some adjustments to fit the work area.\
This are integer (some with fixed adjustments) and flag variables that are used across the script to use in further calculations
~~~
WinGetPos, X, Y, Width, Height, ahk_exe chrome.exe
Y := Y + 100      
total_Area := [X, Y, X + Width, Y + Height - 50]

start_X := 0
start_Y := 0
end_X := 0
end_Y := 0

static_start_X := total_Area[1] + 100
static_start_Y := total_Area[2] + 30
static_end_X := total_Area[3] - 50
static_end_Y := total_Area[4] - 50

absolute_end_X := total_Area[3]
absolute_end_Y := total_Area[4]

up_area_offset_X := 50
up_area_offset_Y := 30

down_area_offset_X := 50
down_area_offset_Y := 30

main_circle_coords := {}
circles_counter := 0

main_circle_ref_X := 0
main_circle_ref_Y := 0
first_circle_found := false

duplicate_there := false

master_stop := false

Test_Existing_Circles_A := 30
Test_Existing_Circles_B := 30

action_counter := 0
~~~

When the main hotkey is triggered, it will recreate those variables, which are inside a funtion called "Create_Variables()".\ 
Next, it will start a Timer with 100 ms interval. The graph can start with or without any circle, the process is different between this two. If the first circle is not found, it will keep looking for the first match and use it as a reference. When the first circle is found, the Time will continue with the "Search_Circle()" function. Like so:
~~~
if (first_circle_found) {  
    Search_Circle()
}
else {
    Search_First_Circle()
}
~~~

The first circle is needed in order to continue, because it is the reference of where the math calculation for the areas are further calculated.\
In the "Search_Circle()" function it is used a For-Loop on the "main_circle_coords" variable. Inside it, it will perform different calculations on 2 different areas to the right of the circle coordinates, up and down. Imagine a rectangle and put the circle on the very left side and vertically on the center, then draw a line, it results in two areas (top and bottom), those are the areas the script will look to.\
After the results, it will then take action depending if a new circle was found or not. Here is an extract:

~~~
for main_key, main_value in main_circle_coords {
    Check_For_Current_Circles()
    Sleep 500 
    if (IsObject(new_circleRU := Check_Right_Up([main_value[1], main_value[2]]))) {                   
     
        if (!IsCircleDuplicated(new_circle_RU)) {  
            Check_For_Current_Circles()       
            found_another := true
            found_another_type := "right_up"   
            circles_counter++
            main_circle_coords.Insert(circles_counter, new_circle_RU) 
            main_circle_ref_X := new_circle_RU[1]
            main_circle_ref_Y := new_circle_RU[2]   
            areas_checker.Insert(circles_counter, {new_circle : 1})    
            TargetAction()
            return
        }
    }
    else{
        areas_checker.Insert(3, {new_circle : 0})  
    }
    Check_For_Current_Circles()

[...]

}
~~~

This is the function that search the Right Up area
~~~
Check_Right_UP(current_circle_coord){
    global
    check_right_value := WaitForFindTextReference_Track([Target_Circle, Target_Circle2, Target_Circle3], DetermineAreRight_Up(current_circle_coord), "Right Up Area")

    if (IsObject(check_right_value)) {
        return check_right_value
    }
    else {
        return 0
    }
}
~~~

There is a helper function "Check_For_Current_Circles" that is called in many stages of the script to validate the curent circles. If one circle is not located anymore, it will decrease by one the "circles_counter" and remove its coordinates on the "main_circle_coords" array to stop looking for it. Inside this function there are other types of calculations that serve to properly conclude the coordinates of the circles.\
The "IsCircleDuplicated()" funcion exist to avoid duplicates, since the circles are being moved to the left overtime. It consist of various math calculations to have a threshold of where a circle can be considered as duplicate or not within the original horizontal line.\
In other words, the scripts looks for new circles that appear on the right side of the first circle. When a new circle is found, it is added to an array with its coordinates, a counter is increased by one and it will perform the clicking actions. The script will omit those circles that was already spotted by doing calculations using their coordinates to not confuse them as new circles, because their current coordinates will change overtime since they are being moved to the left in the graph



# Change Font Type and Opacity quickly in Adobe Illustrator
**Goal**\
Pressing different hotkeys that are assigned for specific Font Types and a hotkey to display the opacity slider quickly

**Overall Process**\
It is used the "FindText" library for key image references on Illustrator to perform clicks and focus elements.
Each hotkey is assigned to trigger a function that will send a specific parameter to differ from the target Font:
~~~
A12_Bold:
Change_Font_Type("Bold")
return

A12_Ultra:
Change_Font_Type("Ultra")
return

A12_Condensed:
Change_Font_Type("Condensed")
return
~~~

Then the main process of the "FindText" to look for a reference is executed to click on the reference with some adjustment and send some keys to focus the edit box and send the corresponding Font Type.
~~~
Change_Font_Type(f_type){
    MouseGetPos, Xo, Yo
    Character_Option:="|<>*108$45.zDzzzzzw1zzztzzaDzzzDzts8E2010D9sT39VVtc30t0DaA8H39tY1c3200AzzzzzzzzzzzzzzzJJJJJJJI"
    if (ok:=FindText(616-150000, 49-150000, 616+150000, 49+150000, 0, 0, Character_Option))
    {
        CoordMode, Mouse
        Step_X_1 := ok.1.x, Step_Y_1:=ok.1.y, Comment:=ok.1.id
        MouseClick, left, % Step_X_1 + 100, % Step_Y_1
        Sleep 100 
        Send {Tab}
        Sleep 100 
        Send %f_type%
        Verify_Font(f_type)
    }
    MouseMove, %Xo%, %Yo%
}
~~~

Finally, there is a "Verify_Font()" function to verify the Font actually have that Font Type option, since some Fonts are not compatible with those. If not, it will promt a message saying "Font Type not found".
Regarding the Opacity slider, it is a similar process but there is 3 key references to look for and click. The last one is the actual slider image reference, and there it is send a "MouseClick" command with the "D" option, so it remains held down and wait for the Left Button key to be clicked.

~~~
if (ok:=FindText(Step_X_1, Step_Y_1 + Y_Offset_1, Step_X_1 + X_Offset_2, Step_Y_1 + Y_Offset_2, 0, 0, Opacity_Slider))
{
    CoordMode, Mouse
    X:=ok.1.x, Y:=ok.1.y, Comment:=ok.1.id
    MouseClick, left, %X%, %Y%,, , D
    KeyWait, LButton, L
}
~~~


# Create new reference in Obsidian app with the highlighted word in a specific header
**Goal**\
Create a reference (named after the highlighted word) and name of the current opened note in a target header

**Overall Process**\
The user highlights a reference with a structure like this "[[List Note|List 2]". The first process is to separate those two parts, the parts separated by the "|" (and sometimes it is "#" instead). For this, it is cleared the "[[" and "]]", any break line and using "StrSplit()" and store it in different variables, like this:

~~~
destination_format_first := StrReplace(destination_format, "[[")
destination_format_first := StrReplace(destination_format_first, "]]")
destination_format_first := StrReplace(destination_format_first, "`n")
destination_format_first := StrReplace(destination_format_first, "`r")

if (InStr(destination_format_first, "|")) {
    destination_format_first_split := StrSplit(destination_format_first, "|")
}
else if (InStr(destination_format_first, "#")) {
    destination_format_first_split := StrSplit(destination_format_first, "#")
}

target_note_file := destination_format_first_split[1]
target_note_to_add := destination_format_first_split[2]
~~~

Now the current name of the note is retrieved by the function "GetActiveNoteTitle()" which sends a sequence of "Send" commands to get it and then others to open the target note.\
After that, the target note is opened by sending some "Send" commands, then a process starts to copy the contents and analyze where to put the desired reference by locating the target header. To make this, all of the contents is copied to the clipboard and then it is added to an array to have more flexibilty regarding where to add the new refrence with precision (by using indexes).\
~~~
ContentToArray(received_content) {
    temp_array := []    
    Loop, Parse, received_content, `n, `r
    {
        temp_array.Push(A_LoopField)
    }
    return temp_array
}
~~~

With the new array, then the scripts uses another function called "FindInContentArray()" and proceed to find the target header with a structure like "# List " (but can be lead by any number of hashtags, for example, "### List 3") by using a For-Loop and testing if the content line matches a Regex Expression
~~~
if (RegExMatch(content_line, "#\s" match "\s*$" , output_regex)) {
    found_line := key
    found := True
}

if (found) {
    return found_line
}
~~~

Then, when the header is found, it is used the index to scan ahead that index to look for the first empty line to add the reference there in order to avoid collision with other references put in place already
~~~
Loop, % received_content.Count() - received_index ; To not over-loop
{
    target_index := A_Index + received_index ; Starting in the found header index to scan ahead
    if (received_content[target_index] = "" || received_content[target_index] = "- ") { ; Look for empty lines or lines with "- " to add the reference there
        found_index := target_index
        found := true
        break
    }
}
~~~

Finally, the content is updated with the new reference injected inside the "UpdateWholeContent_Array()"
~~~
if (target_note_parent_index := FindInContentArray(content_in_array, target_note_to_add_reference_regex, "regex")) {
    if (target_note_to_add_index := FindFirstOccurrenceReference_AfterIndex(content_in_array, target_note_parent_index)) {
        content_in_array[target_note_to_add_index] := "- [[" last_note_title "]]"
        content_in_array.InsertAt(target_note_to_add_index + 1, "") ; Substitute new line
        UpdateWholeContent_Array(content_in_array)        
    } 
    else {
        content_in_array[content_in_array.MaxIndex() + 1] := "- [[" last_note_title "]]"
        UpdateWholeContent_Array(content_in_array)
    }
}
else {
    MsgBox %  "Not Target Header Located" 
    return
}
~~~



# Exit apps safely by imitating a manual project save into a specific folder
**Goal**\
Save the current unsaved projects of apps with a specific filename after an IDLE time to trigger the main process


**Overall Process**\
This script is for multiple work computers at an office that sometimes are left open and need to save the current opened and unsaved projects and then shutdown the PC
First off, there are main arrays that contain specific data for each app that will later use to easily manage the actions on an app
~~~
apps_wintitles := {camtasia: "ahk_exe CamtasiaStudio.exe", notepad: "ahk_exe notepad.exe", xmind: "ahk_exe XMind.exe", davinci_resolve: "ahk_exe Resolve.exe", illustrator: "ahk_exe Illustrator.exe", photoshop: "ahk_exe Photoshop.exe", audition: "ahk_exe Adobe Audition CC.exe"}
save_prompts := {camtasia: "TechSmith Camtasia", notepad:  "Notepad", xmind: "XMind", davinci_resolve: "Message", illustrator: "Adobe Illustrator", photoshop: "Adobe Photoshop", audition: "Audition"}
save_dialogs := {camtasia: "Save As", notepad: "Save As", xmind: "Save As", davinci_resolve: "Create New Project", illustrator: "Save As", photoshop: "Save As", audition: "Save As"}
autosave_folders := {camtasia: A_MyDocuments "\Autosaves", notepad: A_MyDocuments "\Autosaves", xmind: A_MyDocuments "\Autosaves", davinci_resolve: A_MyDocuments "\Autosaves", illustrator: A_MyDocuments "\Autosaves", photoshop: A_MyDocuments "\Autosaves", audition: A_MyDocuments "\Autosaves"}
~~~

Then it validates the folders where the saved files will be saved
~~~
Validate_Autosave_Folders(){
    global

    for key, folder in autosave_folders {
        if (!FileExist(folder)) {
            FileCreateDir, %folder%
        }
    }
}
~~~

Now it will start a Timer, which is monitoring for the IDLE time using the "A_TimeIdle" variable first and then doing an hour of the day verification to trigger the main process or not.\
The main process is restricted to work only after 10 pm (22 hours) and before 7 am (7 hours), those are stored in the main variables
~~~
timestamp_A := 22
timestamp_B := 7
msgbox_timer := 300 ; In seconds. 5 minutes
idle_miliseseconds := 4800000 ; In ms. 80 minutes
save_as_wintitle := "Save As"
filename_suffix := " - autosave before shutdown"
~~~

The timer looks like this

~~~
Check_IDLE:
If (A_TimeIdle > idle_miliseseconds){

    FormatTime, current_hour, , H
    FormatTime, current_minute, , m
    FormatTime, current_seconds, , s

    if (current_hour >= timestamp_A || current_hour <= timestamp_B) { 
        Safe_Exit_Main()         
    }
}
Return
~~~

The main process is within the "Safe_Exit_Main()" function, which also contains the "SafeExit_Target()" function to target the specific apps.\
First it checks if the Camtasia Recorder process exists, if so, it means it typically recording the screen, so hibernate the PC instead of shutdown, otherwise, it continues
~~~
if (WinExist("ahk_exe CamRecorder.exe")) {
    MsgBox, , % "Alert", % "Camtasia Recorder is running, proceeding to hibernate.", % 5

    DllCall("PowrProf\SetSuspendState", "int", 1, "int", 0, "int", 0)
    return
}
~~~

For practical purposes, there was called a function that will be sent individual and hardcoded values from the apps, as shown below
~~~
SafeExit_Target("camtasia", apps_wintitles["camtasia"], save_prompts["camtasia"], save_dialogs["camtasia"], autosave_folders["camtasia"])
SafeExit_Target("davinci_resolve", apps_wintitles["davinci_resolve"], save_prompts["davinci_resolve"], save_dialogs["davinci_resolve"], autosave_folders["davinci_resolve"])
SafeExit_Target("notepad", apps_wintitles["notepad"], save_prompts["notepad"], save_dialogs["notepad"], autosave_folders["notepad"])
SafeExit_Target("xmind", apps_wintitles["xmind"], save_prompts["xmind"], save_dialogs["xmind"], autosave_folders["xmind"])
SafeExit_Target("illustrator", apps_wintitles["illustrator"], save_prompts["illustrator"], save_dialogs["illustrator"], autosave_folders["illustrator"])
SafeExit_Target("photoshop", apps_wintitles["photoshop"], save_prompts["photoshop"], save_dialogs["photoshop"], autosave_folders["photoshop"])
SafeExit_Target("audition", apps_wintitles["audition"], save_prompts["audition"], save_dialogs["audition"], autosave_folders["audition"])
~~~

This was to easily visualize and differentiate each app, but this can be done also using a For-Loop and a master array that will contain all the values, which is a shorter and quick way.\
There was many tests to comprehend the manual process and be able to design it in a coded way. Inside the "SafeExit_Target()" function those steps were programmed involving mane windows-related and the Send commands with occasionally Sleep commands to avoid any crashes or conflicts.\
Note that the apps differ in their way to save a project, they require different steps that need to be programmed and/or specific windows checks, but some are similar, that's why inside this function there are the variable "identifier" to identify the current working app
Down here are extracts of the function to get an idea of what is inside
~~~
SafeExit_Target(identifier, se_target_wintitle, target_save_promt, target_save_dialog, target_save_folder){
    if (WinExist(se_target_wintitle)) {
        SetTitleMatchMode, 2
        WinActivate, % se_target_wintitle
        WinWaitActive, % se_target_wintitle        

        current_index := 0
        Loop, 
        {
            if (A_Index = 1) {
                current_index := A_Index
            }
            else {
                current_index++
            }
            
            if (Check_For_Promtpt(se_target_wintitle, target_save_promt)) {

                WinActivate, % se_target_wintitle
                WinWaitActive, % se_target_wintitle

                ; To Save Dialog

                if (identifier = "camtasia" || identifier = "notepad"|| identifier = "illustrator" || identifier = "photoshop" || identifier = "audition") {
                    Send {Alt Down}f{Alt Up}
                    Sleep 1000
                    Send a
                }
                else if (identifier = "xmind") {
                    Send ^+s
                }
                else if (identifier = "davinci_resolve"){
                    SetTitleMatchMode, 2
                    Send ^+s                    
                    WinActivate, % target_save_dialog
                    WinWaitActive, % target_save_dialog,, 15
                    Send {End 2}
                    Sleep 100
                    Send % filename_suffix
                    Sleep 300
                    Send {Enter}
                }

[...]
                ; Save Dialog
                FormatTime, CurrentDateTime,, MM-dd-yy HH.mm.ss

                if (identifier = "camtasia" || identifier = "notepad" || identifier = "xmind" || identifier = "illustrator" || identifier = "photoshop") {
                    Targeted_Save_In_Save_Dialog(identifier, target_save_dialog, target_save_folder, " - autosave before shutdown (" CurrentDateTime ")" , "suffix", "direct") 
                }

                else if (identifier = "audition") {
                    Targeted_Save_In_Save_Dialog(identifier, target_save_dialog, target_save_folder, " - autosave before shutdown (" CurrentDateTime ")" , "suffix", "direct", "no") 
                }

                ; Extra dialogs
                if (identifier = "illustrator"){
                    SetTitleMatchMode, 2
                    WinActivate, % "Illustrator Options"
                    WinWaitActive, % "Illustrator Options",, 5

                    Send {Enter}
                }

                else if (identifier = "photoshop"){
                    SetTitleMatchMode, 2
                    WinActivate, % "Photoshop Format Options"
                    WinWaitActive, % "Photoshop Format Options",, 5

                    Send {Enter}
                }

[...]
~~~
When the save file dialog appears, it is handled by the "Targeted_Save_In_Save_Dialog()" function, where it establishes first the target save folder and then it sets the filename depending if it exists and the mode
~~~
[...]

if (directory) {
    ControlFocus, ToolbarWindow324, %save_dialog_wintitle% ; Filename
    ControlClick, ToolbarWindow324, %save_dialog_wintitle%,, right
    sleep 1000
    Send e ; To edit
    Sleep, 1000
    Send % directory
    Sleep 1000
    Send {Enter}
    Sleep 1000
}    

[...]

if (filename) {
    ControlFocus, Edit1, %save_dialog_wintitle% ; Filename
    ControlClick, Edit1, %save_dialog_wintitle%,, left

    ; Alternative
    ; ControlFocus, DirectUIHWND3, %save_dialog_wintitle%
    ; ControlClick, DirectUIHWND3, %save_dialog_wintitle%,, left

    if (filename_mode = "suffix") {
        Send {End 2}
        Sleep 1000
        Send % filename
        Sleep 1000
    }
}
[...]
~~~

After all apps are saved in a specific folder with their corresponding filenames, the script proceeds to shutdown the PC with the command
~~~
Shutdown, 1
~~~

The script includes 2 tray icon buttons. One to launch the main process of safe exit and the other one to quit the script
~~~
Menu, Tray, NoStandard,
Menu, Tray, Add, % "Safe Shutdown", Safe_Exit_Main
Menu, Tray, Add, Quit, Quit
~~~



# Print PDF's as they are added to a folder with Adobe Acrobat
**Goal**\
Monitor a folder and when a new PDF is added, print it with Adobe Acrobar through CMD by using "Run %ComSpec% /c"

**Overall Process**\
There is a printer data setup and other information as variables to manage the values easily later in the script.
~~~
target_folder := A_ScriptDir
app_path := "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
printer_name := "Brother QL-1110NWB"
drivername := "Brother QL-1110NWB"
portname := "BTH001"
delay := 15000 ; 15 seconds
~~~

The script uses the "WatchFolder" library, which is used like this in the auto-execute section of the script
~~~
WatchFolder(target_folder, "Watch_For_New_PDF", , Watch := 1)
~~~

The "Watch_For_New_PDF()" function contains the monitoring for a PDF file, by looking into the string of "change.Name" and matching any ".pdf" that means the file is an PDF. Of course there can be a case where the file does include ".pdf" and it is not necesarily a PDF file type, which can be further adjusted to verify the full path of the file and extracting the metadata to confirm its type, but for practical purposes and where a ".pdf" in a file that is not an PDF is very unlikely for this user, it was done as described
~~~
Watch_For_New_PDF(path, changes) {
    global
    for k, change in changes {
        ; 1 means new file was added
        if (change.action = 1) {            
            if (InStr(change.Name, ".pdf")) {
                Print_PDF_File(app_path, change.Name, printer_name, drivername, portname, delay)
            }            
            return
        }
    }
}
~~~

Now, the "Print_PDF_File()" is where the magic happens, it receives the parameters to use in the command line and execute as shown below
~~~
Print_PDF_File(received_app_path, pdf_path, received_printer_name, received_driver_name := "", received_port_name :="", received_delay :="") {
    Run %ComSpec% /c " "%received_app_path%" "/S" "/T" "/O" "/H" "%pdf_path%" "%received_printer_name%" "%received_driver_name%" "%received_port_name%"" ,,hide 
    Sleep %received_delay%
    MsgBox %  "To delete PDF" 
    FileDelete, %pdf_path%    
}
~~~
There is a delay in order to wait prudent time to give space for the PDF to fully print, then the PDF is deleted.






# Quick search in Artgrid website from anywhere on the desktop
**Goal**\
Focus the search bar on Artgrid in Google Chrome whether it is open or not and verify if the tab already exists or created a new one

**Overall Process**\
First, the script will check if the Google Chrome process exists. If it doesn't exist, it will use the "Run" command to open a new instance of chrome and going to the Artgrid website.\
If, on the other hand, the Google Chrome window does exist, it will activate it and call the "Find_Chrome_Tab()" within an "If" statement to look for the returned value (0 or 1).
~~~
!1::
if (WinExist("ahk_exe chrome.exe")) {
	WinActivate, ahk_exe chrome.exe
	WinWaitActive, ahk_exe chrome.exe
	if (!Find_Chrome_Tab("Artgrid.io", 15)) {
		Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" https://www.artgrid.io/
	}
}
else {
	Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" https://www.artgrid.io/
	WinActivate, ahk_exe chrome.exe
	WinWaitActive, ahk_exe chrome.exe
}

Search_Process()	

return
~~~

Inside the "Find_Chrome_Tab()" funciton, it will loop (by a predefined number of loops) through the existing Google Chrome tabs (using Control + Tab to change between tabs forward) and it use the "WinGetTitle, Title, A" command to look for a match in the title of "Artgrid.io". Because when you are navigating in Artgrid, and you are on home or in a search page, when you use the WinSpy, there will be always the "Artgrid.io" text on the window (tab) title.\
If the tab with the match is found by using the "InStr()" native function it will then stop. If there was not any match, will launch a new tab with the Artgrid website. This looping through the existing tabs is done to avoid any duplicated tabs and not have to launch a new tab everytime the hotkey is called, but rather use an existing one, to save resources

~~~
Find_Chrome_Tab(tab_name, max_loop := 15) {
    global
    
    Loop, % max_loop ; Max number of loops
    {
        WinGetTitle, Title, A  ; Get active window title (tab title)
        if (InStr(Title, tab_name))
        {
            return 1 ; Success. Tab found
        }
        Send ^{Tab} ; Go to next Tab forward
        Sleep, 50
    }

    return 0 ; Tab not found
}
~~~

Now, with an active Artgrid website, it will start the "Search_Process()" function that will look for the Magnifier icon in the page by image references using the "FindText" library.\
It was added a timeout to stop looking for the icon in case of any error or similar situation where the search icon is not found within, say, 10 seconds.\
The image references are stored in an INI file and it was used an INI library that instantiates every INI value as a variable to ease the use of them
~~~
Search_Process() {
	global
	
	search_bar := [Values_search_bar_ref_1, Values_search_bar_ref_2]
	StartTime := A_TickCount
	timeout := 10 ; Seconds
	
	Loop {
		for a, reference in search_bar {
			if (ok:=FindText(0-150000, 0-150000, 150000, 150000, 0, 0, reference))
			{
				CoordMode, Mouse
				X:=ok.1.x, Y:=ok.1.y, Comment:=ok.1.id
				Break, 2 ; Break the For and the Loop
			}
		}

		ElapsedTime := A_TickCount - StartTime
		Round(ElapsedTime / 1000)		
		if (ElapsedTime > timeout * 1000) {
			MsgBox %  "Search icon not found!"
			Sleep 1000 
			ToolTip
			return
		}
	}
	MouseClick, left, % X - 50, % Y, 3 ; Click 3 times to select all of the text if any and be able to type right away a new search term
}
~~~



# Perform specific actions overtime (time based triggers). Using Excel and INI data and OBS software
**Goal**\
Execute different types of processes after a specific time-window, for example, trigger one action 15 seconds after the script starts and another after 50 seconds. This, by also using Excel information to open links and starting and stopping OBS record (to record the script actions) and save it into a specific folder

**Overall Process**\
The environment is an Excel file that contains product information from Amazon, a main link and 5 links of similar products. The script will go thorugh each one of the products, performing specific actions inside the product pages and opening the similar products links.\
The script first connects with Excel using COM and using a function from Joe Glines in his Excel Library called "XL_Find_Headers_in_Cols_Number".
~~~
headers_numbers:=XL_Find_Headers_in_Cols_Number(XL,[Values_Product_Link_Header_Name, Values_Brand_Name_Header_Name, Values_Images_Type_Header_Name, Values_Link_1_Header_Name,Values_Link_2_Header_Name,Values_Link_3_Header_Name,Values_Link_4_Header_Name,Values_Link_5_Header_Name]) ; Send search terms as an array
; MsgBox %  "headers_numbers: " headers_numbers[Values_Images_Type_Header_Name] ; Test
~~~

Because this is a time-based script, to ease the triggers it was setup flags for each process, that later they will serve to differentiate and trigger the corresponding action
~~~
main_process := true
process_1_flag := true
process_2_flag := true
process_3_flag := true
process_4_flag := true
process_5_flag := true
process_6_flag := true
process_7_flag := true
process_8_flag := true
process_9_flag := true
process_10_flag := true

links_values := []
current_row := 0

~~~

Inside the main process, the script then finds the total number of items in the table (last row number) with:
~~~
last_row_number := XL.Application.ActiveSheet.Cells(headers_numbers[Values_Product_Link_Header_Name]).EntireColumn.Find("*",,,, 1, 2).Row 
~~~
Then it starts to loop one by one (row by row). First, it opens the main product link and then the 5 similar product links with another loop and adding them into an array.
~~~
product_link_value := XL.Application.ActiveSheet.Cells(current_row, headers_numbers[Values_Product_Link_Header_Name]).Value
Run % "chrome.exe " product_link_value
Sleep 1000 
links_values := []
Loop, 5 
{ 
	links_values.Insert("Link_" A_Index, XL.Application.ActiveSheet.Cells(current_row, headers_numbers[Values_Link_%A_Index%_Header_Name]).Value)
	Run % "chrome.exe " links_values["Link_" A_Index]
}
~~~
It switches back to the first tab, start recording with OBS assign a time variable as reference to make further calculation with time and start the "Time_Logic()" function.\
Note that in order to get Autohotkey working with OBS, you need to setup a delay between the sent keys to let OBS capture the keys, otherwise it probably won't work. I first relied on using the ACC Viewer to start and stop recording, and I was able to do it, but then I found this link, giving me the magic answer to just  add some delays between the Send Commands and it worked very well! Here is the link: ; https://obsproject.com/forum/threads/ahk-not-working-with-obs-studio.70321/
~~~
Send ^1 ; Hotkey to go to first tab
Sleep 1000 

; Start OBS record
Send {F8 Down} ; Native OBS shortcut to start recording
Sleep 500 
Send {F8 Up}

; Initiate time
StartTime := A_TickCount ; Save the current "A_TickCount" to "StartTime"
Time_Logic()

~~~

Now, it starts a recursive function that will call a "Check_Time()" function that will monitor the time passed and will make use of stored timings in the INI file and flags to determine what process is the next one and execute only that one

~~~
Time_Logic() {
	global
	CoordMode, Mouse, Screen

	; P1	0 - 10 seconds: Stay with the mouse centered, not moving
	if (Check_Time() > Values_Process_1_Target_Timing && process_1_flag) { 
		Process_1() 
	}

	; P2	11 - 15 seconds: Move the mouse through the small images
	if (Check_Time() > Values_Process_2_Target_Timing && process_2_flag) {
		Process_2()
	}

	; P3	16 - 20 seconds: Scroll down until the "product description" appears
	if (Check_Time() > Values_Process_3_Target_Timing && process_3_flag) {
		Process_3()
	}

	; P4	21 - 30 seconds: Scroll through the product description (a bit down, a bit up, a bit down again)
	if (Check_Time() > Values_Process_4_Target_Timing && process_4_flag) {
		Process_4()
	}

	; P5	31 - 35 seconds: Home key, to jump to the top
	if (Check_Time() > Values_Process_5_Target_Timing && process_5_flag) {
		Process_5()
	}

	; P6	36 - 40 seconds: Go through the other tabs one by one
	if (Check_Time() > Values_Process_6_Target_Timing && process_6_flag) {
		Process_6()
	}

	; P7	41 - 45 seconds: Switch to the first tab again
	if (Check_Time() > Values_Process_7_Target_Timing && process_7_flag) {
		Process_7()
	}

	; P8	46 - 50 seconds: Hold on the first tab
	if (Check_Time() > Values_Process_8_Target_Timing && process_8_flag) {
		Process_8()
	}

	; P9	50 - 60 seconds: Stop screen recording
	if (Check_Time() > Values_Process_9_Target_Timing && process_9_flag) {
		Process_9()

		return ; Exit recursive function
	}

	Time_Logic()	
}

~~~
The "Check_Time()" function looks like this, is a simple time difference in seconds and then rounded up with the "Round" native function
~~~
Check_Time() {
	global

	ElapsedTime := A_TickCount - StartTime ; Evaluate the difference between the current A_TickCount and the StartTime, it will result in the elapsed time
	; MsgBox,  %  "Milliseconds have elapsed: " ElapsedTime ; Show the results in milliseconds and seconds
	; 	. "`nSeconds passed: " Round(ElapsedTime/1000, 2) ; Calculate the seconds from the milisecconds. 1 second = 1000. CL-1

	return Round(ElapsedTime/1000, 2)	
}
~~~
Each process have their own actions that will be preiodically triggered by time and the flags. Examples of the processes include
* Calculate the center of an area (where the images are) using coordinates and then moving the mouse to hover each individual image in order to preview it. The time hovering an image will be determined by a random number where the user set the minimum and maximum range in the INI file
~~~
[...] 
if (Check_Image_Type() = "vertical") {
	; Coordinates
	first_image_coordinates_values := StrSplit(Values_First_Amazon_Image_Coordenates_Vertical, ",")
	second_image_coordinates_values := StrSplit(Values_Second_Amazon_Image_Coordenates_Vertical, ",")

	first_center := Get_Center_Of_Rectangle(first_image_coordinates_values)
	second_center := Get_Center_Of_Rectangle(second_image_coordinates_values)
	images_offset := second_center[2] - first_center[2]

	move_mouse := first_center.Clone()

	; Go to first image
	MouseMove, % first_center[1], % first_center[2]

	Loop, % Values_Maximum_Image_Number - 1 ; Because above was moved to the first image already
{	
		Random, random_mod_delay, % Values_Seconds_Delay_Per_Image_Min, % Values_Seconds_Delay_Per_Image_Max

		delay_mod := Values_Seconds_Delay_Per_Image + random_mod_delay
			
		Sleep % delay_mod * 1000
		move_mouse[1] := move_mouse[1]
		move_mouse[2] := move_mouse[2] + images_offset ; Increase the value to each iteration and save it
		MouseMove, % move_mouse[1], % move_mouse[2]		
	}
}
[...]
~~~
- Retrieving data from the Excel file to differ the types of image distribution the script will work with, if horizontal or vertical
~~~
Check_Image_Type() {
	global

	; MsgBox %  "" headers_numbers[Values_Images_Type_Header_Name]
	; current_row := 3
	images_type_value := XL.Application.ActiveSheet.Cells(current_row, headers_numbers[Values_Images_Type_Header_Name]).Value

	; MsgBox %  headers_numbers[Values_Images_Type_Header_Name] " " images_type_value

	if (images_type_value = "V" || images_type_value = "v") { 
		return "vertical"	
	}
	else if (images_type_value = "H" || images_type_value = "h") {
		return "horizontal"	
	}

	else {
		MsgBox %  "Incorrect Images Type" 
	}
	
}
~~~
* Scroll down with a random "force" (which is the times of the WheelDown sent) and stop until there is a image reference that serves as a limit, meaning, that when that image is shown on the page, it means the scroll reached the end of the desired area to display.
~~~
[...]
scroll_limit_references := [Values_Scroll_Limit_Reference_1, Values_Scroll_Limit_Reference_2]

	Loop {
		Random, random_mod, % Values_Scroll_Limit_Force_Min, % Values_Scroll_Limit_Force_Max
		scroll_force := Values_Scroll_Limit_Force + random_mod

		Send {WheelDown %scroll_force%}

		Random, random_mod_scroll_delay, % Values_Scroll_Limit_Delay_Min, % Values_Scroll_Limit_Delay_Max
		scroll_delay := Values_Scroll_Limit_Delay + random_mod_scroll_delay

		Sleep % scroll_delay * 1000

		for a, reference in scroll_limit_references {
			; MsgBox % a "-" reference

			if (ok:=FindText(0-150000, 0-150000, 150000, 150000, 0, 0, reference))
			{
				Break, 2 ; Break the For-Loop and the Loop
			}
		}	
	}
[...]
~~~

The recursion on the "Time_Logic()" funciton will stop until it reaches the last process (Process 9), then it stop the recording and rename the newest OBS recording as the name of the Brand, which is on the Excel sheet, and moving the file to a specific folder. Here is how the "Process_9" and the "Renaming_Process()" functions look
~~~

Process_9() {
	global
	process_9_flag := false

	WinClose, ahk_exe chrome.exe
	Stop_Record()

	Renaming_Process()
	Reset_Processes_Flags()

}


Renaming_Process() {
	global
	file_path := FF_Check_For_Files_Quantity(Values_Raw_Folder, 1)
	brand_name_value := XL.Application.ActiveSheet.Cells(current_row, headers_numbers[Values_Brand_Name_Header_Name]).Value
	FormatTime, CurrentDate,, MM.dd.yyyy
	FormatTime, CurrentTime,,  HH.mm
	FileMove, % file_path[1,"path"] ,  % Values_Target_Folder "\" brand_name_value "." file_path[1,"ext"]	
}

~~~

Then it will get prepared for the next Excel row and resetting the flag values with "Reset_Processes_Flags()" to set them all to true, so it can continue with the next product fresh and start the main process again with that new product













# Title
**Goal**\
aaa

**Overall Process**\
aaa












