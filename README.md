# 17TrackDashboard
Given tracking number of shopify orders in a Google Sheets, collect the details from 17track API and make a dashboard out of it.

Steps 
    1. Pop-up to prompt to select a file with .xlsx or .csv extension
    2. Read the file and store it in a dataframe
    3. Check if the file has the required columns (tracking numbers importantly)
    4. Then query the 17track API for each tracking number (register first and then gettrackinginfo) 
    5. Store the response in a dataframe
    6. Sent the dataframe to to a google sheet via gspread
    7. Copy the looker dashboard  using lookerstudio api 
    8. Update the dashboard with the new data (create link and paste in the terminal)
    9. Send an email to the user with the link to the dashboard (optional)