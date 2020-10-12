This solution is comprised of three scripts:
1. Collects free disk space information on VMware volumes, Windows drives and mount points, and stores them in a
database.
2. Queries database in order to generate a trend report
3. Queries database in order to generate alerts based on specifc thresholds

These scripts requires SQLPS (SQL) and PowerCLI (VMware) Powershell modules

SQL database should be setup with a primary key on HOST, DRIVE and DATE, that way there will only be one data point for each drive per day.
This script will display an error message when trying to populate SQL if an entry already exists for that day.

