
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


# Display a message after 7 days, as a expiration date measure
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

The objective is to mimic a expiration date security measure of a script in order to avoid its usage after 7 days.\


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







# Title
**Goal**\
aaa

**Overall Process**\
aaa












