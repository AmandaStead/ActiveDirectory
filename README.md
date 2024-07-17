# ActiveDirectory

Powershell project used for work related tasks to manage active directory users and security groups.

This user interface allows to connect administrative users to their administrative account using their credentials. 
Functions include:

1. Termination:
   -Resets their network password to the default termination password
   -Moves their account to the disabled accounts in active directory
   -Disables their network account
   -Removes all security groups from their account

2. New User:
   -Creates a new network account by first checking for already existing usernames
   -Creates a default password
   -Adds the default security groups and maps their user account "profile" directly to their user drive

3. Modify user account
   -Reset password
   -Enable account
   -Disable account

4. Add security group

All functions have the ability to search active directory for users and security groups. 
