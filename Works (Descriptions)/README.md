
My goal with this is you can see many ways in what Autohotkey can help and get some ideas and insights of work you can do

Some of the works I've done for people include the following:

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
**Notes**\
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





# Title
**Goal**\
aaa

**Overall Process**\
aaa












