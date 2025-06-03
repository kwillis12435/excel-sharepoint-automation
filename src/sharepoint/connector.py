from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

class SharePointConnector:
    def __init__(self, site_url, client_id, client_secret, tenant_id):
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.ctx = None

    def connect(self):
        credentials = ClientCredential(self.client_id, self.client_secret)
        self.ctx = ClientContext(self.site_url).with_credentials(credentials)

    def download_file(self, file_path, local_path):
        if not self.ctx:
            self.connect()
        response = self.ctx.web.get_file_by_server_relative_url(file_path).download(local_path).execute_query()
        return local_path

    def list_files(self, folder_path):
        if not self.ctx:
            self.connect()
        folder = self.ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        self.ctx.load(files)
        self.ctx.execute_query()
        return [file.properties["Name"] for file in files]