Lync 2013 Get External Conference IDs (GUI)
===========================================

            

This is just a quick PowerShell GUI I wrote for someone who requested it.  I thought I'd put a copy up here as a contribution. 


To run it, just double click the script, enter your Front End server or pool FQDN and hit Enter or click Refresh.  It should retrieve all dedicated external conference IDs known to the pool.  You will need administrative access to the localrtc
 database on the Front End you're querying.


A few notes:


This is the first draft version, it forces the script to run with elevated privledges to get around an issue where running on the Front End server itself won't allow access to the localrtc database.  This version only displays
 the external conference ID as seen in the meet URL.  It does not display the DTMF conference ID as displayed in the Outlook invite.  This is mainly because I haven't found where this information is stored (or is it calculated?). 
 If you have insight into the location of that please let me know.  You can contact me directly via LinkedIn in my TechNet profile or add to the Q/A here.


Future versions will include code cleanup, error checking, and hopefully the DTMF ID.


Edit: Minor bug fix, removed issue with spaces in the folder names.


        
    
