import json
import requests


class MicrosoftGraph:
    def __init__(self, token_url, client_id, client_secret, tenant_id, proxy_url, mail_box):
        """
        Initializes the MicrosoftGraph object with configuration parameters.

        Args:
            token_url (str): URL for obtaining access tokens.
            client_id (str): Client ID for authentication.
            client_secret (str): Client secret for authentication.
            tenant_id (str): Tenant ID for authentication.
            proxy_url (str): URL of the proxy server.
            mail_box (str): Mailbox address.
        """
        self.token_url = token_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.proxy_url = proxy_url
        self.mail_box = mail_box

    def get_token(self):
        """
        Retrieves the access token for Microsoft Graph API.

        Returns:
            str: Access token if successful, else None.
        """
        url = "https://login.microsoftonline.com/{}/oauth2/v2.0/token".format(
            self.tenant_id)
        body = {
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        }
        proxies = {'https': self.proxy_url}
        response = requests.post(url, data=body, proxies=proxies)
        if response.status_code < 300:
            return json.loads(response.text)['access_token']
        return None

    def update_message_isread(self, msg_id):
        """
        Updates the read status of a message in the mailbox.

        Args:
            ticket (dict): Dictionary containing message details.

        Returns:
            bool: True if update is successful, else False.
        """

        token = self.get_token()

        URLpatch = f"https://graph.microsoft.com/v1.0/users/{self.mail_box}.eg/messages/{msg_id}"
        proxies = {'https': self.proxy_url}
        headers = {
            "authorization": f"Bearer {token}",
            "content-type": "application/json"
        }

        body = """{
            "isRead": "true"
        }"""
        response = requests.patch(
            URLpatch, headers=headers, proxies=proxies, data=body)

        print('------------ Response -------------')
        print(f'update_message_isread response code={response.status_code}')
        print(f'update_message_isread response text={response.text}')
        print('------------ End of Response -------------')

        if response.status_code < 300:
            return True
        else:
            return False

    def search_new_emails(self, subject, sender):
        """
        Searches for new emails in the mailbox.

        Args:
            subject (str): Subject of the email to search for.
            sender (str): Sender's email address to filter by.

        Returns:
            list: List of email details if successful, else False.
        """

        query = create_query_from_subject(subject=subject)
        token = self.get_token()
        URLget = ""

        if sender:
            URLget = f"https://graph.microsoft.com/v1.0/users/{self.mail_box}.eg/mailFolders('inbox')/messages?$filter=(sender/emailAddress/address) eq '{sender}' and {query} isRead+eq+false"
        else:
            URLget = f"https://graph.microsoft.com/v1.0/users/{self.mail_box}.eg/mailFolders('inbox')/messages?$filter={query} isRead+eq+false"

        proxies = {'https': self.proxy_url}
        headers = {
            "authorization": f"Bearer {token}",
            "content-type": "application/json"
        }

        response = requests.get(URLget, headers=headers, proxies=proxies)

        print('------------ Response -------------')
        print(response)
        print(response.text)
        print(response.status_code)
        print('------------ End of Response -------------')

        if response.status_code < 300:
            return json.loads(response.text)['value']
        else:
            return False

    def send_email(self, sender, mail_body, subject, to=[], cc_emails=[]):
        """
        Sends an email using Microsoft Graph API.

        Args:
            sender (str): Sender's email address.
            mail_body (str): Body of the email.
            subject (str): Subject of the email.
            to (list): List of recipient email addresses.
            cc_emails (list): List of CC email addresses.

        Returns:
            bool: True if email is sent successfully, else False.
        """

        cust_emails_str = self.get_object_mail_addresses(to)
        cc_emails = self.get_object_mail_addresses(cc_emails)
        mail_sender = sender
        token = self.get_token()
        headers = {
            "authorization": f"Bearer {token}",
            "content-type": "application/json"
        }
        URLsend = f"https://graph.microsoft.com/v1.0/users/{mail_sender}/sendMail"

        message_payload = {
            "Message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": mail_body
                },
                "toRecipients": cust_emails_str,
                "ccRecipients": cc_emails,
            },
            "saveToSentItems": False
        }
        message_payload = json.dumps(message_payload)

        print("inside send message")
        print(message_payload, URLsend)
        proxies = {'https': self.proxy_url}
        response = requests.post(
            URLsend, headers=headers, proxies=proxies, data=message_payload)

        print('------------ Response -------------')
        print(f'send mail response status code={response.status_code}')
        print(f'send mail response text={response.text}')
        print('------------ End of Response -------------')

        return response.status_code < 300

    def get_object_mail_addresses(self, customer_list):
        """
        Converts a list of email addresses to a format suitable for Microsoft Graph API.

        Args:
            customer_list (list): List of email addresses.

        Returns:
            list: List of email details.
        """

        out_mail_list = []

        for mail in customer_list:
            out_mail_list.append({
                "emailAddress": {
                    "address": mail
                }
            })

        return out_mail_list

    def receive_message(self, is_read, subject):
        """
        Retrieves messages based on read status and subject.

        Args:
            is_read (bool): Read status of the message.
            subject (str): Subject of the message to search for.

        Returns:
            list: List of message details if successful, else False.
        """

        token = self.get_token()
        URLget = f"https://graph.microsoft.com/v1.0/users/{self.mail_box}.eg/mailFolders('inbox')/messages?$filter=contains(subject,'{subject}') and isRead+eq+{str(is_read).lower()}"
        proxies = {'https': self.proxy_url}
        headers = {
            "authorization": f"Bearer {token}",
            "content-type": "application/json"
        }
        response = requests.get(URLget, headers=headers, proxies=proxies)

        if response.status_code < 300:
            return json.loads(response.text)['value']
        else:
            return False




    def clear_box(self):
        """
        Marks all unread messages in the mailbox as read.
        """

        token = self.get_token()
        URLget = f"https://graph.microsoft.com/v1.0/users/{self.mail_box}.eg/mailFolders('inbox')/messages?$filter=isRead+eq+false"
        proxies = {'https': self.proxy_url}
        headers = {
            "authorization": f"Bearer {token}",
            "content-type": "application/json"
        }
        response = requests.get(URLget, headers=headers, proxies=proxies)

        for mail in json.loads(response.text)['value']:
            self.update_message_isread(dict(mail))

        return 

    def create_query_from_subject(subject):
        """
        Creates a query filter based on subject keywords.

        Args:
            subject (str): Subject of the email.

        Returns:
            str: Query filter string.
        """
        subj_keyword = subject.split(' ')
        subj_keyword = [i for i in subj_keyword if i != "||"]
        query = ''
        for i in subj_keyword:
            query += "contains(subject,'"+str(i)+"') and "
        return query
