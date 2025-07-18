from google.oauth2.credentials import Credentials
from googleapiclient.discovery import Resource, build

from src.settings import GOOGLE_ACCESS_TOKEN, GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN


class GoogleAuthClient:
    TOKEN_URI = "https://oauth2.googleapis.com/token"

    def __init__(self, params: dict[str, str]) -> None:
        self.scopes = params["scopes"]
        self.service_name = params["service_name"]
        self.version = params["version"]

    def build_service(self) -> Resource:
        if not (GOOGLE_CLIENT_ID or GOOGLE_ACCESS_TOKEN or GOOGLE_CLIENT_SECRET):
            print("Google API credentials have not been configured.")

        credentials = Credentials(
            token=GOOGLE_ACCESS_TOKEN,
            client_id=GOOGLE_CLIENT_ID,
            scopes=self.scopes,
            token_uri=self.TOKEN_URI,
            refresh_token=GOOGLE_REFRESH_TOKEN,
            client_secret=GOOGLE_CLIENT_SECRET,
        )

        return build(credentials=credentials, serviceName=self.service_name, version=self.version)
