# architect-excel
An Excel Add-In to access Architect backend via the Typescript API
https://github.com/architect-xyz/architect-ts

For users of the Architect trading platform who want to access some of the trading functionality via Excel.

This add-in allows Architect users to connect see prices, positions, balances, and pnl.
Users should already have an account with Architect, along with API key / secret.

https://architect-xyz.github.io/architect-excel/


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

### To Test
npx office-addin-debugging start docs/manifest.xml
npx office-addin-debugging stop docs/manifest.xml
