# Bnos Yisroel Efficiency Scripts
## Timeclock Program 24-25
* See README in Timeclock folder

## Bloomerang Reports
* Takes excel reports emailed monthly and uploads them to drive
* Extracts data from files in the drive to populate a report spreadsheet, matching the files to the correct place in the report

## Daycare Payroll
* Triggered on paydays
* Takes monthly timeclock data, matches to individual reports and calculates time worked per pay period
* Catches errors when missing a clock in or clock out
* Updates individual tabs and google sheets for each employee for reporting
* Formats sheets for easy navigation including highlighting totals, red highlighting errors, bolding titles etc
* Updates consolidated report for business office

## Payroll
* processes information from forms empolyees fill out
* Deals with hourly vs period employees, specific rates and jobs for each employee
* Calculates payments and populates individual tabs and sheets for reporting including formating of sheets
* Automatically creates a report each payday to consolidate total payments per job per employee since last payday
* Emails employees to confirm form submissions and errors. Sends daily reminders for days employees are scheduled to work and haven't submitted a form yet
* Deals with side cases such as jobs that have a submission cap, get extra pay for working a certain amount, or can be filled out for extended work periods
  
## Electives Sorting
  * Creates reports of which classes students still need to take based on class categories, previous classes taken, and class requirements

## Carpool Slot Assigning
* Sort through all form submissions that represent carpools and divide into tab for each division's submissions
* On the division tabs that hold all current students, assign what their method of carpool is and comments submitted
* If method type is carpool, assign them a spot. Spots are A-Z and 1-12. If it is a big van carpool, they get a spot past P. Start with alternating spots then fill in spots once all alternates are filled
* If a student goes home with a different division, find them and fill in division and carpool type
