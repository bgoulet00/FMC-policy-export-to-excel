# FMC-policy-export-to-excel
Export Cisco FMC access control policy to excel sheet

This script will generate an excel xlsx file dump of a Cisco FMC access control policy
The output is meant to be a reference of the policy that appears similar in layout to the web GUI
The output is NOT suitable or intended to be a backup of the policy
Additionally, Users, Source Dynamic Attribute and Destination Dynamic Attribute are not accounted for but
updates to include those items should be trivial if you use those technologies

Developed and tested with the following environment
OS: windows10
Python: 3.11.5
Target platform:  FMC 7.0.4
Dependencies: protocols.csv input file with protocol number to name mappings. this is provided in the repo
Limitations: 
  - Users, Source Dynamic Attribute and Destination Dynamic Attribute are not accounted for
  - FMC queries are limited to 1000 results. paging to support greater than 1000 has not been implemented
Caveats: sometimes rules are added incorectly in FMC and appear to be part of a category when they are actually
   'Undefined'.  this script will ignore those anomolies in an attempt to replicate the layout of the FMC GUI
