<h1>How To Use</h1>

You will definitely need to re-jig this for your own uses. I haven't touched the code since then. You will need an Excel file and a .ics file, both of which were given to you from the UofT scheduler (from J. Trinh or at least she was for mine). To run:

<code>python3 DentalCalendar2020.py \<input.xls\> \<input2.ics\></code>

and it will create customized .ics calendars for students 1-120 inside the folder "Dental Calendars/". To get them into PDF format is a manual chore: open Microsoft Outlook, create a new empty calendar (as to not screw up your current one if you have one), import one of the .ics files, and then export to PDF given a particular time range (likely early September to June-Aug).

You can see what the output of my program is within "Dental Calendars/" with the example input files that were given to me back then for an idea of what it produces. Note that I only had 118 students in my cohort that year (no 61 or 120).

From 3rd year to 4th year of dental school, I had to spend a few hours tinkering with it as since I was graduating when the pandemic occurred, they added a triple clinical session (AM, PM1, and PM2) on certain days. So if you want to use this for your own purposes, you'll have to read my code, figure out how it works, and then re-jig it for your purposes. Godspeed.

<h2>Dependencies</h2>
<ul>
  <li><a href="https://pandas.pydata.org/">pandas</a> - For reading the Excel file</li>
  <li><a href="https://pypi.org/project/icalendar/">icalendar</a> - For reading and writing .ics files</li>
</ul>

<h1>How It Works (Broadly Speaking)</h1>

<ol>
  <li>Read the Excel file and store into a Python data structure</li>
  <li>At specific rows (clinic days) and columns (student ID) - which have to be figured out by hand by the way - fetch the clinic session and store into a 2D array</li>
  <li>Read the .ics file and start a counter from 1 to 120</li>
  <ol>
    <li>Go through each event in the .ics file one by one</li>
    <ul>
      <li>If it is a clinical event or ancillary clinic then intercept it and inject the particular data from the 2D array</li>
      <li>If not, just change its colour to something aesthetically pleasing</li>
    </ul>
    <li>Output to create a new .ics file for said counter</li>
  </ol>
</ol>
