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
    """Returns a list of class-type Teams. Does not return classes missing the classCode property."""

    r = sess_graph.get(graph_endpoint + '/education/classes')
    r.raise_for_status()

    teams_classes = json.loads(r.text)['value']
    return [t_class for t_class in teams_classes if 'classCode' in t_class]


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
        r = sess_graph_j.post(graph_endpoint + '/education/classes/' +
                              class_id + '/members/$ref', data=json.dumps(body))
        debug_print(r.text)
        r.raise_for_status()
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

    body = {
        "shouldSetSpoSiteReadOnlyForMembers": True
    }

    if config['dry_run']:
        return None
    else:
        r = sess_graph_j.post(graph_endpoint + '/teams/' +
                              team_id + '/archive', data=json.dumps(body))
        debug_print(r.text)
        r.raise_for_status()
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
        r = sess_graph_j.post(graph_endpoint + '/groups/' +
                              group_id + '/members/$ref', data=json.dumps(body))
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code


def remove_group_member(group_id, user_id):
    """Removes a member from an Office 365 Group. Returns HTTP status code; 204 indicates success."""

    if config['dry_run']:
        return None
    else:
        r = sess_graph.delete(graph_endpoint + '/groups/' +
                              group_id + '/members/' + user_id + '$ref')
        debug_print(r.text)
        r.raise_for_status()
        return r.status_code
