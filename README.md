# QuickPick
Quick Picker PCF that allows you to pick the queue item from the entity you are on


##Features
Pick a record without navigating back to the queue items list

##Limitations
Binds to a field that isnt used.

##Wish List
- Make generic for any queue enabled record
- fail more gracefully on creation of new record
- refresh after save of new record
- Swap to only do API call when button is clicked for pick, this will reduce the number of API calls
- Swap to only do API call when button is clicked for release, this will reduce the number of API calls
- Swap to only do API call when button is clicked for route to user search, this will reduce the number of API calls
- Have the option to "wuick route" to another user/team
- Have the option to only show team members for routing or more detailed config options for teams and/or users

## Enable To
Bind to any single text field as it isn't used
- SingleLine.Text

##Setup
Use the solution.zip to install managed component to your environment

##Solution creation
navigate to QuickPick folder
create folder called QuickPickSolution
pac solution init --publisher-name XXX --publisher-prefix YYY

pac solution add-reference --path ../

msbuild /t:build /restore

##Solution install
Manually download zip file from releases and import.

##Solution Deployment
To deploy to your environment using the command line
copy the repository locally
connect to remote env
pac auth create --url XXX
enter credentials
command to push to env
pac pcf push --publisher-prefix YYY
