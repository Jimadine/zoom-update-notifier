
# Zoom Update Notifier for Windows

An unofficial Powershell script to notify you when there is a Zoom update available

## Why this script exists
Out of frustration that — despite Zoom Workplace running in the background all of the time — the Zoom Updater functionality only tells me that an update is required at the very moment I attempt to join a meeting. An update usually takes a few minutes to install, so has made me late for meetings!

:rage: :rage: :rage:

## What the script does
* It extracts the version number of the currently installed version of Zoom from the `zoom.exe` executable file on your computer
* It then makes a single `HEAD` request to a specific Zoom URL. The response returned includes the version number of the latest version in the `Location` header
* It then compares the two version numbers, and if the latest version is greater than the installed version, it will pop a dialogue box (or send an email to you) to let you know:

![An image of the test window](./window_screenshot_live.png?raw=true)

## Setting things up
The script is intended to be set up as a daily scheduled task with one or more time triggers (depending on how many times a day you wish to have the script perform the check).

Before setting up a scheduled task, it is suggested that you test the script by opening a Powershell window and passing the `-IsTest` switch parameter:
```
Set-Location "C:\path\to\repo"
.\zoom-update-notifier.ps1 -IsTest
```
You should see a pop-up as follows:

![An image of the test window](./window_screenshot_test.png?raw=true)

<sub>In the above screenshot, the installed version and latest version match, which is expected when passing the `-IsTest` switch parameter</sub>

If you want to test the emails you can also do that by passing the `-EnableEmail` parameter and the other email-related parameters to have the script send you an email:

```
.\zoom-update-notifier.ps1 -IsTest -EnableEmail -MailFromAddress fromaddress@example.org -MailToAddress youraddress@example.org -MailServer smtp.example.org
```

When creating the scheduled task `Action`, here are the commands and arguments that I would suggest adding:
```
Program/script: C:\Windows\System32\wscript.exe
Arguments: //B "C:\path\to\silent.vbs" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy ByPass -File "C:\path\to\zoom-update-notifier.ps1"
```
Note `silent.vbs` is a small generic wrapper script included in this repo that allows you to run Powershell silently. It's the only bulletproof method to avoid seeing a Powershell window.

Example scheduled task `Action` for emails:
```
Program/script: C:\Windows\System32\wscript.exe
Arguments: //B "C:\path\to\silent.vbs" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy ByPass -File "C:\path\to\zoom-update-notifier.ps1" -EnableEmail -MailFromAddress fromaddress@example.org -MailToAddress youraddress@example.org -MailServer smtp.example.org
```

## Notes
* If there are any errors then these will usually be displayed in pop-up windows, not by email even if you've passed the email parameters
* If the script presents an error about `zoom.exe` not existing at the standard path of `C:\Program Files\Zoom\bin\zoom.exe`, you can pass the `-ZoomExePath` parameter to change where the script looks for `zoom.exe`
* If you're unsure where `zoom.exe` is, `where.exe /F /R \ zoom.exe` is your friend.
* When creating the scheduled task, and setting time triggers for times that your computer is likely to be turned off at, it's probably worth enabling the `Run the task as soon as possible after a scheduled start is missed` option under the `Settings` tab
* The script's email functionality relies on the deprecated `Send-MailMessage` cmdlet and so only supports unauthenticated email
