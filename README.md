# Architect Add-in for Excel
An Excel Add-In to access Architect backend via the Typescript API
https://github.com/architect-xyz/architect-ts

For users of the Architect trading platform who want to access some of the trading functionality via Excel.

This add-in allows Architect users to connect see prices, positions, balances, and pnl.
Users should already have an account with Architect, along with API key / secret.

https://excel.architect.co


## Lower Latency
For clients needing a lower latency add-in in C#, please contact support@architect.co

## For Maintainers


### To Add Functions
Add functions to src/functions.ts

### To Build
```bash
npm install
npx webpack
```

### To Validate manifest.xml

office-addin-manifest validate manifest.xml


#### Sideloading App
npx office-addin-debugging start manifest.xml
npx office-addin-debugging stop manifest.xml


#### Remove Cache
In File Explorer, go to
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\

(Change 16.0 to your Office version if needed, e.g., 15.0 or 17.0.)

Delete the folder contents (this will clear the cache for all sideloaded add-ins).

Restart Excel.


#### TO DO:

- use web workers? https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-add-in-custom-functions-using-web-workers



#### Helpful Resources:

Shared Runtime:
https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/excel-shared-runtime-global-state/manifest.xml
https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime?tabs=xmlmanifest


