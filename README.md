# powershell-bpac
Brother P-Touch bPAC SDK in PowerShell

# Get Brother b-PAC SDK and DLL

https://www.brother.co.jp/eng/dev/bpac/index.aspx

# Troubleshooting
1. Use Brother LE Drivers<br />https://stackoverflow.com/questions/23155315/brother-label-printer-sdk-bpac-3-1-failed-to-print
2. Ensure correct paper set in Printer Preferences
3. Ensure correct printer and paper set on label; be aware of width vs length and start from a template and get that printing first.

## Error numbers

If an error is returned from:
- SetPrinter()
- StartPrint()
- PrintOut()
- Close()
- EndPrint()

Be sure that the points above are checked.

## Printer flashes, or has error

Ensure you are sending the correct size paper setting.

# Modifying Label

Install P-Touch Editor software

https://support.brother.ca/app/answers/detail/a_id/133156/~/download-and-install-the-p-touch-editor-software
