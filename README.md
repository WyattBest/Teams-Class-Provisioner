# Teams Class Provisioner

A simple Python 3 project that gets a list of current sections from PowerCampus, along with teachers and students, and creates/updates matching Classes in Microsoft Teams.

Uses Microsoft Graph API v1.0 and SQL connection to PowerCampus database. Requires SQL Server 2016 or newer for JSON-Path support.

# settings.json
debug: prints a lot of extra information if true
dry_run: if true, only simulates making changes to Graph API. Useful with debug, which will print simulated changes.
refresh_sections: if true, pulls sections from PowerCampus. If false, uses cached sections from last run when refresh_sections was true. This option exists for speed when debugging/testing; I was using a linked server with JSON-Path support that couldn't optimize the sections query.
