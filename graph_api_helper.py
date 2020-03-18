import json
import requests
import graph_auth_helper
import pyodbc

# Read config file
with open('settings.json') as config_file:
    config = json.load(config_file)
graph_endpoint = config['Microsoft']['graph_endpoint']

# Create persistent HTTP session without Content-Type header
sess_graph = requests.Session()
sess_graph.headers.update({
    "Authorization": graph_auth_helper.get_auth_header()
})

# Create persistent HTTP session with Content-Type: application/json header
sess_graph_j = requests.Session()
sess_graph_j.headers.update({
    'Authorization': graph_auth_helper.get_auth_header(),
    'Content-Type': 'application/json'
})


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


def get_classes():
    """Returns a list of class-type Teams. Does not return classes missing the classCode property or archived classes."""

    r = sess_graph.get(graph_endpoint + '/education/classes')
    r.raise_for_status()

    response = json.loads(r.text)
    teams_classes = response['value']

    # Get additional pages from server
    while '@odata.nextLink' in response:
        r = sess_graph.get(response['@odata.nextLink'])
        r.raise_for_status()
        response = json.loads(r.text)
        teams_classes.extend(response['value'])

    # Add isArchived property to each class
    parameters = {'$select': 'isArchived'}
    for t_class in teams_classes:
        # Print progress, since this is really slow
        pos = teams_classes.index(t_class) + 1
        print(str(pos) + ' of ' + str(len(teams_classes) + 1))

        r = sess_graph_j.get(graph_endpoint + '/teams/' +
                             t_class['id'], params=parameters)
        try:
            r.raise_for_status()
            t_class['isArchived'] = json.loads(r.text)['isArchived']
        except requests.exceptions.HTTPError:
            # Graph API tends to 404 or 500 on newly-created Teams
            if r.status_code == 404 or r.status_code == 500:
                t_class['isArchived'] = None
            # Retry bad gateway errors up to 10 times
            elif r.status_code == 502:
                debug_print(r.text)
                for attempt in range(10):
                    try:
                        r = sess_graph_j.get(
                            graph_endpoint + '/teams/' + t_class['id'], params=parameters)
                        r.raise_for_status()
                        t_class['isArchived'] = json.loads(r.text)[
                            'isArchived']
                    except:
                        if r.status_code == 502:
                            debug_print(r.text)
                            continue
                        else:
                            break
            else:
                raise

    return [t_class for t_class in teams_classes if 'classCode' in t_class and t_class['isArchived'] != True]


def get_class_members(class_id):
    """Returns a list of students and teachers for the given class."""

    parameters = {'$select': 'id,displayName,mail,userType,userPrincipalName'}
    r = sess_graph.get(graph_endpoint + '/education/classes/' +
                       class_id + '/members', params=parameters)
    r.raise_for_status()
    response = json.loads(r.text)
    members = response['value']

    # Get additional pages from server
    while '@odata.nextLink' in response:
        r = sess_graph.get(response['@odata.nextLink'])
        r.raise_for_status()
        response = json.loads(r.text)
        members.extend(response['value'])

    return members


def get_class_teachers(class_id):
    """Returns a list of teachers for the given class."""

    parameters = {'$select': 'id,displayName,mail,userType,userPrincipalName'}
    r = sess_graph.get(graph_endpoint + '/education/classes/' +
                       class_id + '/teachers', params=parameters)
    r.raise_for_status()
    members = json.loads(r.text)['value']
    return members


def create_class(name, description, class_code, external_id, mail, term, external_name=None):
    """Creates a class. External_name defaults to name. Term is an object with nested elements; see Graph API docs."""

    if external_name is None:
        external_name = name

    body = {
        'description': description,
        'classCode': class_code,
        'displayName': name,
        'externalId': external_id,
        'externalName': external_name,
        'externalSource': 'sis',
        'mailNickname': mail,
        'term': term
    }

    if config['dry_run']:
        return None
    else:
        r = sess_graph_j.post(
            graph_endpoint + '/education/classes', data=json.dumps(body))
        debug_print({'new class response': r.text})
        r.raise_for_status()
        return json.loads(r.text)


def add_class_teacher(class_id, teacher_id):
    """Adds a teacher to a Team. Returns HTTP status code; 204 indicates success."""

    body = {
        '@odata.id': graph_endpoint + '/education/users/' + teacher_id
    }

    if config['dry_run']:
        return None
    else:
        r = sess_graph_j.post(graph_endpoint + '/education/classes/' +
                              class_id + '/teachers/$ref', data=json.dumps(body))
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code


def add_class_student(class_id, student_id):
    """Adds a student to a Team. Returns HTTP status code; 204 indicates success."""

    body = {
        '@odata.id': graph_endpoint + '/education/users/' + student_id
    }

    if config['dry_run']:
        return None
    else:
        try:
            r = sess_graph_j.post(graph_endpoint + '/education/classes/' +
                                  class_id + '/members/$ref', data=json.dumps(body))
            debug_print(r.text)
            r.raise_for_status()
        except requests.HTTPError:
            # Why does this 404 sometimes? User licensing issue?
            if r.status_code == 404:
                debug_print(r.text)
            else:
                raise
        return r.status_code


def remove_class_teacher(class_id, teacher_id):
    """Removes the specified teacher from the specified Teams class. Returns 204 if successful."""

    if config['dry_run']:
        return None
    else:
        r = sess_graph.delete(graph_endpoint + '/education/classes/' +
                              class_id + '/teachers/' + teacher_id + '/$ref')
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code


def remove_class_student(class_id, student_id):
    """Removes the specified student from the specified Teams class. Returns 204 if successful."""

    if config['dry_run']:
        return None
    else:
        r = sess_graph.delete(graph_endpoint + '/education/classes/' +
                              class_id + '/members/' + student_id + '/$ref')
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code


def archive_team(team_id):
    """Archives an existing team and makes SharePoint site read-only. Returns status code 202 if successful."""

    # Not currently working. See https://github.com/microsoftgraph/microsoft-graph-docs/issues/4944
    # body = {
    #     "shouldSetSpoSiteReadOnlyForMembers": True
    # }

    if config['dry_run']:
        return None
    else:
        r = sess_graph_j.post(
            graph_endpoint + '/teams/' + team_id + '/archive')

        debug_print(r.text)

        try:
            r.raise_for_status()
        except requests.exceptions.HTTPError:
            # Archiving tends to bomb out  while waiting for backend state consistency.
            # We'll log the error and keep going.
            if r.status_code == 404:
                debug_print({'Error archiving Team:': team_id})
                return 500
            else:
                raise

        return r.status_code


def get_group_owners(group_id):
    """Returns a list of owners of the given group."""

    parameters = {'$select': 'id,displayName,mail,userType,userPrincipalName'}
    r = sess_graph.get(graph_endpoint + '/groups/' +
                       group_id + '/owners', params=parameters)
    r.raise_for_status()
    owners = json.loads(r.text)['value']
    return owners


def get_group_members(group_id):
    """Returns a list of members of the given group."""

    parameters = {'$select': 'id,displayName,mail,userType,userPrincipalName'}
    r = sess_graph.get(graph_endpoint + '/groups/' +
                       group_id + '/members', params=parameters)
    r.raise_for_status()
    response = json.loads(r.text)
    members = response['value']

    # Get additional pages from server
    while '@odata.nextLink' in response:
        r = sess_graph.get(response['@odata.nextLink'])
        r.raise_for_status()
        response = json.loads(r.text)
        members.extend(response['value'])

    return members


def add_group_member(group_id, user_id):
    """Adds a member to an Office 365 Group. Returns HTTP status code; 204 indicates success."""

    body = {
        '@odata.id': graph_endpoint + '/directoryObjects/' + user_id
    }

    if config['dry_run']:
        return None
    else:
        try:

            r = sess_graph_j.post(
                graph_endpoint + '/groups/' + group_id + '/members/$ref', data=json.dumps(body))
            debug_print(r.text)
            r.raise_for_status()
            return r.status_code
        except requests.HTTPError:
            # Why does this 404 sometimes? User licensing issue?
            if r.status_code == 404:
                debug_print(r.text)
            else:
                raise
        return r.status_code


def remove_group_member(group_id, user_id):
    """Removes a member from an Office 365 Group. Returns HTTP status code; 204 indicates success."""

    if config['dry_run']:
        return None
    else:
        r = sess_graph.delete(graph_endpoint + '/groups/' +
                              group_id + '/members/' + user_id + '/$ref')
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code
