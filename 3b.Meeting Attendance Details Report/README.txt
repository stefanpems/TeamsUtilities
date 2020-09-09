This PowerShell script creates a CSV report with the details of the users' attendance to the meetings planned in the organizational's Office 365 groups/teams.

The main CSV output report containes the following header row:
"UserLoginName","StartDateTime","EndDateTime","TeamEmail","TeamDisplayName","TeamVisibility"

The script's execution ha 3 phases. If you have MS Access installed on your PC, you may want to speed up the complete execution by configuring the script to skip the third phase: it can be done much quicker by using the delivered Access file with the sample queries.

Please refer to the initial script preamble for any other detail.

UPDATE: in "Output Sample 1.png" you can see the output (imported in Access) of the class meeting attendance in the last 30 days in a real school 
