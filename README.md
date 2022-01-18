# Walled-Garden Add-in for Outlook

A compliance Add-in for Microsoft Outlook to analyse sender lists for recipients outside your organization.

## How it Works

The Add-In will extract the sender's (`From: `) organization domain to match it against all recipients.

When the email is Sent, all Recipients - in `To: `, `CC: ` and `BCC: ` - will be matched against the sender's domain.

Any recipient whose domain does not match the sender's will trigger a popup, listing each "outsider" recipient and confirming if the messgae should be send anyway.

Answering `No` will return to the compose window.

Answering `Yes` will release the message.

## Installing

1. Go to [Releases](https://github.com/michaelsanford/Outlook-Walled-Garden/releases) and choose the `Latest`.
2. Download the Zip and validate the checksum.
3. Run `setup.exe`
4. Restart Outlook to load the Add-in

## Uninstalling

1. Press Start
2. Type "Add or remove programs"
3. Search for "Walled Garden for Outlook"
4. Choose it and select `Uninstall`.

## Contributing

I welcome contributions and bug reports; please file an issue.

### Attributions

Add-in icon courtesy [bitfreak86](https://www.iconfinder.com/bitfreak86)
