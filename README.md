# Teams Class Provisioner

A simple Python 3 project that gets a list of current sections from PowerCampus, along with teachers and students, and creates/updates matching Classes in Microsoft Teams.

Uses Microsoft Graph API v1.0 and SQL connection to PowerCampus database. Requires SQL Server 2016 or newer for JSON-Path support.

# Faculty Team
All faculty members returned by the sections query will be assigned to this Team. The Team GUID should be placed in settings.json

# Students Team
All students returned by the sections query will be assigned to this Team. The Team GUID should be placed in settings.json

# settings.json
## Microsoft section
`application_id`: Found in the "Application (CLIENT) ID" column under App registrations in the Azure portal. The application must have the following API permissions:

  * EduRoster.ReadWrite.All
  * Group.ReadWrite.All
  * Member.Read.Hidden
  * User.ReadWrite.All
  
`secret`: Client secret generated in the Azure portal under App registrations.

`registrar_id`: The registrar will be added as a teacher to all classes. Set to `null` to disable this behavior.

## PowerCampus section
`database_string`: A [pyodbc connection string](https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-SQL-Server-from-Windows) to your PowerCampus SQL server. The example setup is for Kerberos authentication on Windows, but you can modify it for Linux or other platforms.

## Other settings
`debug`: If true, prints a lot of extra information.

`dry_run`: If true, only simulates making changes to Graph API. Useful with debug, which will print simulated changes.

`clear_cache_sections`: If true, pulls sections from PowerCampus. If false, loads cached sections from last run when this setting was true. This option exists for speed when debugging/testing.

`clear_cache_users`: If false, loads the user cache from the last run. New users would still be looked up and added to the cache. This option exists for speed when debugging/testing.
