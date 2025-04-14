# AD_Server_Audit
## Objective
Find all servers in AD, determine if they are up or down through ICMP and/or WinRM. If WinRM is running, not blocked by firewall, then gather the servers Physical and Logical processor count, and the hardware model the server is running on. 
## Parameters
- *ExcelFile* - The path and file name of the Excel file that will be exported. Defaults to your system environments "My Documents" path and saves the file as AD Server Audit.xlsx. 
## Output the following column headers
- Server
- OS
- AD Tree
- Created
- Last Changed
- ICMP
- WinRM
- Physical Proc Count
- Logical Proc Count
- Model
