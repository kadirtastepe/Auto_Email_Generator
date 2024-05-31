@echo off

echo Welcome to Automated Reminder Mail Template Generator v-1.1b updated on (29/02/2024)
echo Maintaned by Kadir Tastepe
echo ========================================================================================================
echo Please make sure that the Microsoft Edge browser is located on the Laptop screen not in 2nd monitor.
echo Please make sure that reset the Microsoft Edge settings to Zoom is 50%
echo Screen Resolution must be set to 1920x1080
echo ========================================================================================================
echo Checking the python installation ...

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed.
    exit /b
)


REM Check if the necessary python packages are installed. If not then install.
echo Python installation is found.
echo Checking if required libraries are installed...
python -c "import time" && echo Python package TIME found || (echo Python package TIME is not found && pip install time)
python -c "import PIL" && echo Python package PIL found || (echo Python package PIL is not found && pip install pillow)
python -c "import pyautogui" && echo Python package Pyautogui found || (echo Python package Pyautogui not found && pip install pyautogui)

REM Check if python requires update
echo Checking python updates...
python -m pip install --upgrade pip

REM Check the currrent date and time
echo Getting the current date and time automatically from the system time...
for /F "tokens=1-3 delims=/ " %%A in ('date /T') do (
    set "day=%%A"
    set "month=%%B"
    set "year=%%C"
)

REM For reminder date
set "rem_cur_date=%month%/%day%"
echo Today: %rem_cur_date%
echo Calculating the date that is 7 days before...
powershell -Command "(Get-Date).AddDays(-7).ToString('MM/dd')" > temp.txt
set /p rem_pre_date=<temp.txt
del temp.txt

echo Date Interval "ForRem=%rem_pre_date%-%rem_cur_date%"

REM Create the current date in the YYYYMMDD format
set "current_date=%year%%month%%day%"

REM Calculate the date that is 7 days before
powershell -Command "(Get-Date).AddDays(-7).ToString('yyyyMMdd')" > temp.txt
set /p previous_date=<temp.txt
del temp.txt

set "link1=https://___.LINK.__.com&period_type=dp&date_from=%previous_date%&date_to=%current_date%&"

echo Link1 opening...

start msedge "%link1%" --force-device-scale-factor=0.5
timeout /t 8 

echo ScreenShot in action...

REM Create a python file for screenshots
(
echo import pyautogui
echo from PIL import ImageGrab
echo import time

echo screenWidth, screenHeight ^= pyautogui.size^(^)

echo pyautogui.click^(1542, 515^)

echo pyautogui.click^(1542, 515, duration^=1^)

echo time.sleep^(1^)

echo image^=ImageGrab.grab^(bbox^=^(0,215,1884,1000^)^)

echo output^_path ^= ^'ScreenShot.png^'

echo image.save^(output^_path^)

echo print^(^"Screenshot saved to^:^", output^_path^)

) > capture.py

python ./capture.py

REM Automatically finds the script location
set "script_location=%~dp0"
set "script_location=%script_location:~0,-1%"
REM echo Script location: %script_location%


REM Set the folder path for images
set "image_folder=%script_location%"

REM Create an HTML file with the desired content
(
    echo ^<html^>
    echo ^<body^>
    echo ^<p^>Hello colleagues, ^</p^>
    echo ^<p^>You get this mail as you are subscribed to ^<a href="__LINK__.com"^> Gorgeous_Mail_List. ^</a^> ^</p^>
    echo ^<p^>Please find below the tracking items from day period week %rem_pre_date%-%rem_cur_date% from. ^</p^> 
    echo ^<p^> ^<b^>Week ^(^(%rem_pre_date%-%rem_cur_date%^)^)^</b^> ^</p^>
    echo ^<p^> 1^. Weekly Tracking^: ^</p^>
    echo ^<p^> ^<a href="%link1%"^>Title of Link 1^</a^>^</p^>
    echo ^<p^> ^<img src="%ScreenShot.png" alt="Example Image"^>^</p^>
    echo ^<p^> Thanks and best regards, ^</p^>
    echo ^</body^>
    echo ^</html^>
) > output.html

REM Set the path to the HTML file
set "html_file=./output.html"

REM Read the HTML content from the file
set "html_content="
for /f "usebackq delims=" %%A in ("%html_file%") do (
    set "html_content=!html_content!%%A"
)

set "recipient=List_Of_Recipients"
set "subject=Weekly Reporting - Week (%rem_pre_date%-%rem_cur_date%)"
set "cc=Tastepe, Kadir <kadir.tastepe@cern.ch>;"

REM Send an email using Outlook with the HTML content as plain text

start outlook.exe /c ipm.note /a "%html_file%" /m "%recipient%?Cc=%cc%&subject=%subject%"

del capture*




