import tableauserverclient as TSC
import datetime
import requests
import os

# Configurtion
server_url = 'https://prod-useast-b.online.tableau.com/'
site_name = 'b360bi'
token_name = 'myToken'
access_token = 'OQlxX9hPR+qhWzx3bOlv/Q==:SYpTs58I4M2xFCv4mcF6xbhbOzxDv1c7'
api_version = '3.22'

ppt_location = r'C:\Workspace\Automations\content'
today_date = datetime.datetime.today().strftime('%Y-%m-%d')

# Specify Workbook Name
project_name = 'Hyper API'
workbook_name = 'Fraud Analysis'

def connect_tableau():
    # Step 2: Connect to the Tableau Server
    try:
        tableau_auth = TSC.PersonalAccessTokenAuth(token_name=token_name,
                                                   personal_access_token=access_token,
                                                   site_id=site_name)
        server = TSC.Server(server_url)
        server.auth.sign_in(tableau_auth)
        print("DEBUG: Connected to Tableau Server")
        return server
    except Exception as err:
        print("ERR: Error connecting to Tableau Server")
        raise err

def get_workbook_id(server, workbook_name, project_name):
    """ Retrieves the workbook ID based on the given workbook name and project name.

    Args:
        server (TableauServer): The Tableau Server object.
        workbook_name (str): The name of the workbook to search for.
        project_name (str): The name of the project where the workbook resides.

    Returns:
        int: The ID of the matching workbook.

    Raises:
        FileExistsError: If no workbook with the specified name is found or if multiple workbooks
                        share the same name.
    """
    # Get all workbooks in the specified project
    all_workbooks = list(TSC.Pager(server.workbooks))
    matching_workbooks = [workbook for workbook in all_workbooks if
                          workbook.name == workbook_name and workbook.project_name == project_name]
    if len(matching_workbooks) == 0:
        error = f"Workbook '{workbook_name}' not found in project '{project_name}'."
        raise FileExistsError(error)
    elif len(matching_workbooks) > 1:
        error = f"Multiple workbooks with name '{workbook_name}' found in project '{project_name}'."
        raise FileExistsError(error)

    return matching_workbooks[0].id


def make_api_request(api_url, headers):
    """
    Makes an API request and returns the response.

    Args:
        api_url (str): The URL of the API endpoint.
        headers (dict): Headers to include in the request.

    Returns:
        requests.Response: The API response.
    """
    response = requests.get(api_url, headers=headers)
    return response


def parse_and_save_pptx(response, output_dir = "", workbook_name = "target"):
    """
    Parses the API response and saves the .pptx content to a local file.

    Args:
        response (requests.Response): The API response.
    """
    if response.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        # Save the .pptx content to a local file
        filename = f"{workbook_name}.pptx"
        output_path = os.path.join(output_dir, filename)
        with open(output_path, "wb") as pptx_file:
            pptx_file.write(response.content)
        print(f"File saved to {output_path}")
    else:
        print("Unexpected response type:", response.headers.get("Content-Type"))

server = connect_tableau()
server.version = '3.22'

workbook_id = get_workbook_id(server, workbook_name, project_name)

print(workbook_id)

authentication_token = server._auth_token
site_luid = server._site_id
api_url = f"{server_url}/api/{api_version}/sites/{site_luid}/workbooks/{workbook_id}/powerpoint"
headers = {'X-tableau-auth': authentication_token}

api_response = make_api_request(api_url,headers)
parse_and_save_pptx(api_response, ppt_location, workbook_name)

server.auth.sign_out()

