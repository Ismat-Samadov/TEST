# report_automation
we can automate our reporting with two way:

which daily_report_scheduled.py file running in background with loop.
also we can do this with taskschd.msc which  report_daily.py file belong second way.

here are detailed description for running report_daily.py in windows:


To ensure that your script runs daily, you can utilize external task scheduling tools such as Windows Task Scheduler (on Windows) or cron (on Unix-like systems) to run your Python script at the desired time each day. These tools are designed for this purpose and are more reliable than embedding the scheduling logic within your script.

Here are the steps to set up a scheduled task using Windows Task Scheduler:

Save your Python script (daily_report.py) to a location on your computer.

Open Windows Task Scheduler by searching for it in the Windows Start menu or by running taskschd.msc in the Run dialog (Win + R).

In Task Scheduler, click on "Create Basic Task" on the right-hand side.

Follow the wizard to set up your task:

Give your task a name and description.
Choose "Daily" as the trigger and click "Next."
Set the start date and time you want the script to run daily.
Choose "Repeat task every" and set it to "1 days."
Specify how many days you want to recur (leave it as 1 for daily execution).
Select "Start a program" as the action and click "Next."
In the "Program/script" field, browse and select the Python executable (python.exe) on your system. You should provide the full path to python.exe, for example: C:\Python39\python.exe.

In the "Add arguments (optional)" field, provide the full path to your Python script (daily_report.py), for example: "C:\path\to\daily_report.py"

Click "Next" and review your settings. If everything looks correct, click "Finish" to create the task.

The task is now scheduled to run daily at the specified time.

With this setup, Windows Task Scheduler will take care of running your script daily without the need for the scheduling logic inside the script. Make sure to keep your system running or awake at the scheduled time for the task to execute as expected.

