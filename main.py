import requests
import graph_auth_helper
import graph_api_helper
import json
import pyodbc


def debug_print(x):
    """Attempt to print JSON without altering it, serializable objects as JSON, and anything else as default."""
    if config['debug'] and len(x) > 0:
        if isinstance(x, str):
            print(x)
        else:
            try:
                print(json.dumps(x, indent=4))
            except:
                print(x)


def clean_sql_json(x):
    """Cleans up JSON produced by SQL Server by reducing this pattern:
        [{"Key": [{"Key": "Value"}]}]
    to this:
        [{'Key': ['Value']}]

    Also removes duplicates (and ordering) from the reduced list.
    Returns an object, not JSON.
    """

    datas = json.loads(x)

    for data in datas:
        for key, value in data.items():
            if (isinstance(value, list)
                    and isinstance(value[0], dict)
                    and len(value[0]) == 1
                    ):
                data[key] = list({
                    list(item.values())[0]
                    for item in value
                })

    return datas


def get_userPrincipalName(PEOPLE_CODE_ID):
    """Sub-function called by get_user_id(). Executes SQL to get userPrincipalName from PowerCampus."""
    cursor.execute(get_userPrincipalName_sql, PEOPLE_CODE_ID)
    try:
        userPrincipalName = cursor.fetchone()[0]
        return userPrincipalName
    except TypeError:
        return None


def get_user_id(PEOPLE_CODE_ID):
    """Looks up userPrincipalName in PowerCampus based on PCID, then looks up userId in Graph API.
    Keeps in-memory cache to reduce querying. Return None if user not found or if user is unlicensed.
    """

    # Set up a persistent HTTP session
    if 'sess_gui' not in globals():
        global sess_gui
        sess_gui = requests.Session()
        sess_gui.headers.update(
            {"Authorization": graph_auth_helper.get_auth_header()})

    # Init cache
    if 'cached_users' not in globals():
        global cached_users
        cached_users = {}

    if PEOPLE_CODE_ID not in cached_users:
        cached_users[PEOPLE_CODE_ID] = {}

    if 'userPrincipalName' in cached_users[PEOPLE_CODE_ID]:
        userPrincipalName = cached_users[PEOPLE_CODE_ID]['userPrincipalName']
    else:
        userPrincipalName = get_userPrincipalName(PEOPLE_CODE_ID)
        cached_users[PEOPLE_CODE_ID]['userPrincipalName'] = userPrincipalName

    if userPrincipalName is None:
        debug_print({'lookup user': PEOPLE_CODE_ID,
                     'response': 'No record in PersonUser!'})
        return None

    if 'userId' not in cached_users[PEOPLE_CODE_ID]:
        r = sess_gui.get(graph_endpoint + '/users/' +
                         userPrincipalName + '?$select=displayName,id')

        if r.status_code == 404:
            cached_users[PEOPLE_CODE_ID]['userId'] = None
            debug_print({'lookup person': PEOPLE_CODE_ID,
                         'userPrincipalName': userPrincipalName, 'response': json.loads(r.text)})
        else:
            r.raise_for_status()
            response = json.loads(r.text)['id']
            debug_print({'lookup person': PEOPLE_CODE_ID,
                         'userPrincipalName': userPrincipalName, 'response': response})

            # Check that found user has an O365 license
            r = sess_gui.get(graph_endpoint + '/users/' +
                             response + '/licenseDetails')
            r.raise_for_status()

            if json.loads(r.text)['value'] is None:
                cached_users[PEOPLE_CODE_ID]['userId'] = None
            else:
                cached_users[PEOPLE_CODE_ID]['userId'] = response

    return cached_users[PEOPLE_CODE_ID]['userId']


# Read config file
with open('settings.json') as config_file:
    config = json.load(config_file)
debug_print(config)

graph_endpoint = config['Microsoft']['graph_endpoint']

# Microsoft SQL Server connection.
cnxn = pyodbc.connect(config['PowerCampus']['database_string'])
cursor = cnxn.cursor()
# Check connection. Authentication should be Kerberos.
cursor.execute(
    'SELECT auth_scheme FROM sys.dm_exec_connections WHERE session_id = @@spid;')
row = cursor.fetchone()
debug_print({
    'sql connection check':
    {
        'database': cnxn.getinfo(pyodbc.SQL_DATABASE_NAME),
        'auth_scheme': row[0]
    }
})
# Cache query text
with open('get_userPrincipalName.sql') as sql:
    get_userPrincipalName_sql = sql.read()

# Load cached users
if config['clear_cache_users'] == False:
    with open('cached_users.json') as file_users:
        cached_users = json.load(file_users)['cache']

# Get sections
if config['clear_cache_sections'] == True:
    # Get list of sections and students from PowerCampus
    print('Querying PowerCampus sections list...')
    with open('get_current_sections.sql') as sql:
        cursor.execute(sql.read())
    rows = cursor.fetchall()
    sections = clean_sql_json(''.join([row[0] for row in rows]))

    # Save sections to cache file
    with open('cached_sections.json', mode='w') as file_sections:
        json.dump(sections, file_sections, indent=4)
else:
    # Load cached instead of live sections list
    print('Using cached sections list.')
    with open('cached_sections.json') as file_sections:
        sections = json.load(file_sections)

debug_print(sections)

print('Fetching Teams classes.')
# Get list of Teams classes.
teams_classes = graph_api_helper.get_classes()
debug_print({'current Teams classes': teams_classes})

print('Updating classes.')
# Compare to sections and create any new classes.
# Newly-created classes will not have members added immediately; Office 365 usually takes some minutes to provision a new class.
for sect in sections:
    if sect['classCode'] in [t_class['classCode'] for t_class in teams_classes]:
        debug_print({'no action': sect['classCode']})
    else:
        debug_print({'create class': sect['classCode']})
        graph_api_helper.create_class(sect['EVENT_LONG_NAME'], sect['classCode'],
                                      sect['classCode'], sect['SectionId'], sect['mailNickname'], sect['term'][0])

# For any Teams classes not in sections, archive the Teams class and mark for removal.
for t_class in teams_classes:
    pos = teams_classes.index(t_class) + 1
    print(str(pos) + ' of ' + str(len(teams_classes) + 1))

    if t_class['classCode'] in [sect['classCode'] for sect in sections]:
        debug_print({'no action': t_class['classCode']})
    else:
        debug_print({'archive class': t_class['classCode']})
        graph_api_helper.archive_team(t_class['id'])
        t_class['Delete'] = True

# Remove archived teams classes from list.
teams_classes[:] = [t_class for t_class in teams_classes if 'Delete' not in t_class]

print('Updating members in classes.')
for t_class in teams_classes:
    pos = teams_classes.index(t_class) + 1
    print(str(pos) + ' of ' + str(len(teams_classes) + 1))

    # For each Teams class, get the list of teachers.
    t_teachers = []
    t_teachers = [t_teacher['id']
                  for t_teacher in graph_api_helper.get_class_teachers(t_class['id'])]

    # Lookup PowerCampus teachers from sections by classCode, then translate PCID list to O365 userId's.
    pc_teachers = []
    pc_teachers_pcid = [sec['SECTIONPER']
                        for sec in sections if sec['classCode'] == t_class['classCode']][0]
    if pc_teachers_pcid is not None:
        pc_teachers = [get_user_id(t_user) for t_user in pc_teachers_pcid]
    # Add registrar(s) to each class. Set setting to null to make this stop.
    pc_teachers = pc_teachers + config['Microsoft']['registrars']
    debug_print({'class': t_class['classCode'],
                 'pc_teachers': pc_teachers, 't_teachers': t_teachers})
    # Make lists into unordered, unique sets and remove None
    t_teachers = set(t_teachers) - {None}
    pc_teachers = set(pc_teachers) - {None}

    # Add new teachers from sections.
    for teacher in pc_teachers.difference(t_teachers):
        debug_print({'class': t_class['classCode'], 'add teacher': teacher})
        graph_api_helper.add_class_teacher(t_class['id'], teacher)

    # Remove extra teachers not in sections.
    for teacher in t_teachers.difference(pc_teachers):
        debug_print({'class': t_class['classCode'], 'remove teacher': teacher})
        graph_api_helper.remove_class_teacher(t_class['id'], teacher)

    # For each Teams class, get the list of students (actually all members; there's no API call for just students).
    t_members = []
    t_members = [t_member['id']
                 for t_member in graph_api_helper.get_class_members(t_class['id'])]

    # Lookup PowerCampus students from sections by classCode, then translate PCID list to O365 userId's.
    pc_students = []
    pc_students_pcid = [sec['TRANSCRIPTDETAIL']
                        for sec in sections if sec['classCode'] == t_class['classCode']][0]
    if pc_students_pcid is not None:
        pc_students = [get_user_id(t_user) for t_user in pc_students_pcid]
    debug_print({'class': t_class['classCode'],
                 'pc_students': pc_students, 't_members': t_members})
    # Make lists into unordered, unique sets and remove None
    t_members = set(t_members) - {None}
    pc_students = set(pc_students) - {None}

    # Add new students from sections.
    for student in pc_students.difference(t_members):
        debug_print({'class': t_class['classCode'], 'add student': student})
        graph_api_helper.add_class_student(t_class['id'], student)

    # Remove extra students not in sections.
    # Because get_class_members() returns students + teachers, include teachers set when comparing.
    for student in set(t_members - t_teachers).difference(pc_students):
        debug_print({'class': t_class['classCode'], 'remove student': student})
        graph_api_helper.remove_class_student(t_class['id'], student)

with open('cached_users.json', mode='w') as dump_file:
    json.dump({'description': 'A dump of the cached_users object from last time sync was completed.',
               'cache': cached_users
               }, dump_file, indent=4)

print('Updating Faculty group members.')
# Update members of existing Faculty team
faculty_team = config['Microsoft']['faculty_team']
t_owners = []
t_members = []
pc_faculty_pcid = []

t_owners = [owner['id']
            for owner in graph_api_helper.get_group_owners(faculty_team)]
t_members = [member['id']
             for member in graph_api_helper.get_group_members(faculty_team)]

# List PowerCampus teachers from sections, then translate PCID list to O365 userId's.
pc_faculty_pcid = [item for sublist in sections if sublist['SECTIONPER']
                   is not None for item in sublist['SECTIONPER']]
pc_faculty = [get_user_id(t_user) for t_user in pc_faculty_pcid]

pc_faculty = set(pc_faculty) - {None}
t_owners = set(t_owners) - {None}
t_members = set(t_members) - {None}


print('Updating Student group members.')
# Update members of existing Student team
student_team = config['Microsoft']['student_team']
t_owners = []
t_members = []
pc_students_pcid = []

t_owners = [owner['id']
            for owner in graph_api_helper.get_group_owners(student_team)]
t_members = [member['id']
             for member in graph_api_helper.get_group_members(student_team)]

# List PowerCampus students from sections, then translate PCID list to O365 userId's.
pc_students_pcid = [item for sublist in sections if sublist['TRANSCRIPTDETAIL']
                    is not None for item in sublist['TRANSCRIPTDETAIL']]
pc_students = [get_user_id(t_user) for t_user in pc_students_pcid]

pc_students = set(pc_students) - {None}
t_owners = set(t_owners) - {None}
t_members = set(t_members) - {None}

# Add new students from sections.
for student in pc_students.difference(t_members):
    debug_print({'add to Students team': student})
    graph_api_helper.add_group_member(student_team, student)

# Remove extra students not in sections.
for student in set(t_members - t_owners).difference(pc_students):
    debug_print({'remove from Students team': student})
    graph_api_helper.remove_group_member(student_team, student)

# Parse cached_users and output suspicious entries to file
error_users = {}

for PCID, results in cached_users.items():
    if 'userId' not in results or 'userPrincipalName' not in results:
        error_users[PCID] = cached_users[PCID]

for PCID, results in cached_users.items():
    for key, value in results.items():
        if value is None:
            error_users[PCID] = cached_users[PCID]

with open('error_users.json', mode='w') as dump_file:
    json.dump({'description': 'Users with possible error states.',
               'users': error_users
               }, dump_file, indent=4)

print('Finished!')
