---

# Microsoft Graph API Wrapper

This is a Python wrapper for interacting with the Microsoft Graph API to manage email communication using the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview).

## Features

- Send emails using the Microsoft Graph API.
- Search for new emails based on subject and sender.
- Mark messages as read.
- Retrieve messages based on read status and subject.

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/OmarTaher11/microsoft-graph-api-wrapper.git
   ```

2. Navigate to the project directory:

   ```bash
   cd microsoft-graph-api-wrapper
   ```

3. Install the required dependencies using pip:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Import the `MicrosoftGraph` class from the module:

   ```python
   from microsoft_graph_api import MicrosoftGraph
   ```

2. Initialize the `MicrosoftGraph` object with your configuration parameters:

   ```python
   graph = MicrosoftGraph(
       token_url='your_token_url',
       client_id='your_client_id',
       client_secret='your_client_secret',
       tenant_id='your_tenant_id',
       proxy_url='your_proxy_url',
       mail_box='your_mailbox_address'
   )
   ```

3. You can now use the methods provided by the `MicrosoftGraph` class:

   ```python
   # Send an email
   graph.send_email(
       sender='sender@example.com',
       mail_body='Hello, this is a test email.',
       subject='Test Email',
       to=['recipient1@example.com', 'recipient2@example.com'],
       cc_emails=['cc1@example.com', 'cc2@example.com']
   )

   # Search for new emails
   emails = graph.search_new_emails(
       subject='Important',
       sender='sender@example.com'
   )

   # Mark a message as read
   message_id = 'message_id_here'
   graph.update_message_isread(msg_id=message_id)

   # Retrieve messages based on read status and subject
   unread_messages = graph.receive_message(is_read=False, subject='Notification')

   # Mark all unread messages as read
   graph.clear_box()
   ```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
