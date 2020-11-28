from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime
import pandas as pd
import pygsheets

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def make_query():
    """
    Returns a list of queries(string) in
    the same query format as the Gmail search box.
    """
    # Add queries here if you have additional queries
    query_lst = ['label:ocbc']
    return query_lst


def tracker(query):
    """
    Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    try:
        response = service.users().messages().list(userId='me', q=query).execute()

        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId = 'me', q=query, pageToken=page_token).execute()
            messages.extend(response['messages'])

        msg_list = []

        for item in messages:
            msg = service.users().messages().get(userId='me', id=item['id'], format = 'metadata').execute()
            msg_list.append(msg)

        return msg_list

    except Exception as e:
        print('An error occured: ', e)




def parse_ocbc(lst):
    """
    Takes in a list of messages.
    Returns a pandas dataframe with 4 columns (Date, Desc, Amount, and Source).
    """
    dates = []
    descs = []
    amounts = []

    for msg in lst:
        # Get date
        header = msg['payload']['headers']
        for dct in header:
            if dct['name'] == 'Date':
                date = dct['value']
                stop = date.find(':') - 3
                date = datetime.strptime(date[:stop], '%d %b %Y')
                dates.append(date)

        # Get description and amount
        snippet = msg['snippet']
        desc_idx = snippet.find('.')
        amt_start = snippet.find('SGD') + 4
        amt_end = snippet.find('From') - 1

        descs.append(snippet[:desc_idx])
        amounts.append(snippet[amt_start:amt_end])

    # Reverse the order
    dates.reverse()
    descs.reverse()
    amounts.reverse()

    df = pd.DataFrame(data = {'Date': dates, 'Description': descs, 'Amount': amounts, 'Source': ['OCBC Pay Anyone']*len(dates)})
    return df


def to_google_sheets(df):
    """
    Takes in a pandas dataframe.
    Modifies the Google Sheet spreadsheet.
    Returns None.
    """
    # Authorisation
    gc = pygsheets.authorize(service_file='/Users/kevin/PycharmProjects/ExpensesTracker/creds.json')
    # Open the Google Sheets spreadsheet
    sheet = gc.open('Expense Tracker')
    # Select the first sheet
    curr = sheet[0]
    # Update the sheet with the df supplied as input starting at cell A1 (1, 1)
    curr.set_dataframe(df, (1, 1))


def main():
    # Build query
    query = make_query()
    # Build list of messages
    msg_lst_ocbc = tracker(query[0])
    # Create dataframe
    df_ocbc = parse_ocbc(msg_lst_ocbc)
    # Modifies spreadsheet
    to_google_sheets(df_ocbc)

    print('Edit Done Succesfully.')


if __name__ == '__main__':
    main()
