# NSX-T-Reporting
A collection of python scripts which pulls out information from your NSX-T environment.
The majority of these scripts will output to a spreadsheet detailing the NSX-T environment.
These scripts primarily run against the Policy API, except for those components which use the manager API.
Some of these scripts use the Python SDK for NSX-T so you will need this installed, and can be found here:

https://code.vmware.com/web/sdk/3.0/nsx-t-python

Other scripts use the API directly

## NSX-Manager Info
Information from the NSX Manager Cluster 

## SERVICES
A script extracting all of the services in NSX-T along with their service entries, types, ports, tags and scope and outputs to an excel file.

## GROUPS
Information pertaining to security groups configured in the NSX environment

## Logical Switcches / Segments
Information pertaining to segments & logical switches configured in the NSX environment

