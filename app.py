import os.path
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "120zYmwpnKiOLPjWQXfKuJyMjrNjucsvONkka6LqzajM"
SAMPLE_RANGE_NAME = "engenharia_de_software!C4:F27"


def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        values = result.get("values", [])

        results = [ ]
        for linha in values:
            situation = ""
            score = 0
            absent = int(linha[0])
            average = (float(linha[1])+float(linha[2])+float(linha[3]))/3
            if ((absent/60)>0.25):
                situation = "Reprovado por falta"
            elif (average>=70):
                situation = "Aprovado"
            elif (average>=50):
                situation = "Exame Final"
                score = 100 - average
            else:
                situation = "Reprovado por Nota"
            tupla = (situation, math.ceil(score))
            results.append(tupla)

        result = (
            sheet.values()
            .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="G4", valueInputOption="USER_ENTERED", body={'values':results})
            .execute()
        )

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()