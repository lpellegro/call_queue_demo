INSTRUCTIONS

These scripts create call queues in Webex Calling and assign agents to the queues based on provisioning details included in a .xlsx spreadsheet.

If you don't a Webex org you can use dcloud.cisco.com to setup a Webex Calling lab. 
Just log into dcloud and look for "Cisco Webex Calling v3" lab.

Once you have an org, you can go to developer.webex.com and login using your org admin credentials.
Then click on "Documentation". 

![img.png](img.png)

On the left hand side you have the list of APIs. Click on "Full API Reference" and then on any of the APIs. On the API page, right hand side, copy the "Bearer".
This is a limited-duration access token that you can paste in the "credentials" file. After 12 hours this token is no longer valid and you have to grab a new one from the same web page.
A full OAuth flow will automatically renew access token before they expire, but this is out of the scope of these scripts.

Next, create an Excel spreadsheet by populating the first row with the queue name (i.e. "Queue1", "Queue2"), and include the agent's email under the specific queue.
Make sure that the Excel spreadsheet doesn't have blank rows or columns.
Launch the script: the queues will be populated and agents assigned.

If you want to delete those queues, just run the delete_queue script. This script deletes all queues that have been created