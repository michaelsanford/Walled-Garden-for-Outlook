# Walled-Garden Add-in for Outlook

A compliance Add-in for Microsoft Outlook to analyse sender lists for recipients outside your organization.

## How it Works

The Add-In will extract the sender's (`From: `) organization domain to match it against all recipients.

When the email is Sent, all Recipients - in `To: `, `CC: ` and `BCC: ` - will be matched against the sender's domain.

Any recipient whose domain does not match the sender's will trigger a popup, listing each "outsider" recipient and confirming if the messgae should be send anyway.

Answering `No` will return to the compose window.

Answering `Yes` will release the message.

## Installing in Outlook

//TODO once the artifact is published on Releases.

## Contributing

I welcome contributions and bug reports; please file an issue.