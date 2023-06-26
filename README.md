# Community_Tips_Reascript
A reascript used to draw creative tips from a spreadsheet that will be stored within your local machine.


Hi, thank you for downloading this Reascript! 

The idea behind this was to create something light and 'fun' that wouldn't require a mass install of other supporting environments. So, this script has been created using LUA, rather than Python or C-based languages. Also, the spreadsheet with the tips & tricks is contained within this repository and will copy to your local machine. Should you like, you can edit this to add your own tips and it'll only reflect on your copy. I researched keeping this in the cloud to keep refreshes exciting, but this would require the user (you!) to install further modules and so on.


========================================================================================================================================

So, here are the instructions to get this up and running (it's not a lot, I promise!)



1. Download this repository as a .zip file (likely already done), place the folder in your desired destination and unzip it (important to do this first)

2. Open Reaper, go to 'Actions' -> 'Show Actions List'

3. Within the Actions List box, select the "New Action" button and then select "Load Reascript" 

4. From here, select the .lua file you've recently downloaded within the "Community_Tips_Reascript-main" folder (assuming you have kept the default name)

5. This will now show as the top result in the Actions list. Select it, and then select the 'Edit Action...' button
  - You will now see the Reascript Development Environment and the lua code for the script

6. Find the line that says (local spreadsheetPath = "PASTE YOUR FILE PATH HERE - KEEP QUOTATIONS") and replace the TEXT IN CAPITALS with the file path of the .csv file and make sure to keep the quotation marks. This should be Line 14 (the environment doesn't have a line counter)

7. Once this has been pasted in, 'save' the file (CTRL+S on Windows, CMND+S on Mac) and this will execute the script
  - You should see the dialogue box appear with a random prompt from the spreadsheet!

8. Close the dialogue box and the Reascript Development Environment

9. Assign your preferred action to the script by clicking the 'add' button with your new script highlighted (as you would when editing exisiting actions)


And that should be you go to go!


Thank you again for checking this out. Hopefully it'll help in moments or creative block.

Aaron (@ARAudio_)
