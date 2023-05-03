# location_transfers

A previous employer utilized one vendor to track the location and performance data of their rental units and another to track billing and contracts.  

The ecosystem of the first vendor ("Vendor1" in the code) used an asset tree structure with the asset site, then asset number at the lowest end.   

My team had a daily task of reconciling the site of each unit in that ecosystem with the location listed by the billing software ("Vendor2" in the code).  Vendor1 offers an API with an endpoint suited to the task, while Vendor2 only allowed custom API's with additional cost and time needed for development.

My approach was to create a dictionary of the unit numbers and their assigned site from each vendor, then create a CSV with the entries that differ.  Since Vendor2 doesn't offer an API, there's code to navigate the desktop software, download an Excel report, then build the dictionary from that report.  

Al Sweigart has my unending appreciation (and some of my money (: ) for his Automate the Boring Stuff with Python course and book. 
