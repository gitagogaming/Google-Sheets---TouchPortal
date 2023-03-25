


## Sample Leaderboard Google Sheet
![image](https://user-images.githubusercontent.com/76603653/227666672-8e476636-2993-40f1-9b6a-ac979526ca4b.png)
Copy this and use as your own. - https://docs.google.com/spreadsheets/d/18qDz-0B0negMaXONQr1wtnc0prk1Eop3bBdygJsWaRw/edit#gid=1710931458
 <br><br>
##Sample Scoreboard - Production
![image](https://user-images.githubusercontent.com/76603653/227666898-c8aea73a-ff4f-4a82-a48f-7c578b4c7f87.png)
This is the the scoreboard tob e used with the Example Config Below

 <br><br>
## Example Config
Please be sure to mimic the format used in the example config.  
 <br><br>
## Sample Page
![image](https://user-images.githubusercontent.com/76603653/226514197-a6a1bbd4-9aab-47b5-a9f2-d53908b751da.png)

```ini
worksheetName = "Sample_Scoreboard"     ## This is the name of the worksheet within the spreadsheet
spreadsheetId = "YOUR_SPREADSHEET_ID"   # This is the spreadsheet URL / ID


### This example below will take I8 and name it "BLUE_TEAM_NAME" inside of TouchPortal.. 
### A TouchPortal state will be created with the name provided and its value will match the cell on your google sheet 

[Text]
I9 = "BLUE_NAME"
I7 = "BLUE_SCORE"
K9 = "ORANGE_NAME"
K7 = "ORANGE_SCORE"
E13 = "CASTER1"
E14 = "CASTER1_TWITTER"
H13 = "CASTER2"
H14 = "CASTER2_TWITTER"



## Input Cells here that contain image URLs
## If wanting an image saved as another format then you may put into Images.JPG or similar
[Images]
    ## Images put in this section will retain their original file type


	## Images put below this will be changed to the file type suggested
    [Images.PNG]

    [Images.JPG]
       

    [Images.GIF]
        


## Input cells containing video URLs here
## Videos will retain original file type
[Videos]
    
```
