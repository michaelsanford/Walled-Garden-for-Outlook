# Walled-Garden Add-in for Outlook

A compliance Add-in for Microsoft Outlook to analyse sender lists for recipients outside your organization.

## How it Works

The Add-In will extract the sender's (`From: `) organization domain to match it against all recipients.

When the email is Sent, all Recipients - in `To: `, `CC: ` and `BCC: ` - will be matched against the sender's domain.

Any recipient whose domain does not match the sender's will trigger a popup, listing each "outsider" recipient and confirming if the message should be sent anyway.

Answering `No` will return to the compose window.

Answering `Yes` will release the message.

## Installing

1. Go to [Releases](https://github.com/michaelsanford/Outlook-Walled-Garden/releases) and choose the `Latest`.
2. Download the Zip and validate the checksum.
3. Run `setup.exe`
4. Restart Outlook to load the Add-in

## Troublehooting

1. Open Outlook
2. Ensure there is no yellow warning under `Slow and Disabled COM Add-ins`. You can re-enable it there.
3. Navigate to `File` > `Options` (bottom-left, above `Exit`) > `Add-ins`
4. Ensure "Walled Garden for Outlook" is installed and not disabled.

ℹ️ If Outlook has disabled this add-in on your system because it was flagged as being slow to start, [please file a bug](https://github.com/michaelsanford/Outlook-Walled-Garden/issues/new/choose).

## Uninstalling

### From Windows
1. Press Start
2. Type "Add or remove programs"
3. Search for "Walled Garden for Outlook"
4. Choose it and select `Uninstall`.

### Forcing Removal in Outlook

If you have installed this Add-in using a Release package (ran a `setup.exe`) use the above method instead.

1. Open Outlook
2. Navigate to `File` > `Options` (bottom-left, above `Exit`) > `Add-ins`
3. Choose `Manage: COM Add-ins` and click `Go`.
4. Choose "Walled Garden for Outlook" and click `Remove`.
5. Restart Outlook.



## Contributing

I welcome contributions and bug reports; please file an issue.

### Attributions

Add-in icon courtesy [bitfreak86](https://www.iconfinder.com/bitfreak86)
