# Timeclock Program 24-25
1. Trigger created and Script Properties used to store which employee the script is up to so if program times out, will restart on proper row
2. Loops through dates in the current week
3. For hourly employees, loops through each day and checks time worked
4. For period employees, loops through times for each period they work each day
5. Checks punches to calculate time late/ left early/absences based on specific teacher schedule
6. Logs information on master sheet in proper tab per employee
7. Form is filled out for lateness/absence/early leaving with reasons and times and this gets added to masterSheet on proper line
8. Teachers emailed when they come late, leave early, or are absent to ask the reason. Form can then be filled out and added to masterSheet
9. masterSheet is then copied to individual sheets so each employee can view their reports
10. Admin gets reports of teachers over their allocated time at the end of each pay-period
11. Schedule changes: planned are accounted for and form can be submitted for unplanned (per division or whole school)

    ## Other functions:
    * Handling times in different formats and subtracting times
    * Dealing with differences in scheduling including exceptions
    * Creating sheets and tabs automatically for new employees and sorting the tabs
    * Creating a backup masterSheet every week
    * Converting between period based schedules and time based schedules
    * Formatting sheets properly including merging cells, putting equations into sheets, colors and fonts, hiding and resizing columns
    * Adds new month line whenever the newest punch record is a new month
