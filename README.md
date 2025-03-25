# architect-excel
Excel plugin to access Architect backend via the Typescript API
https://github.com/architect-xyz/architect-ts

## Lower Latency
For clients needing a lower latency plugin in C#, please contact support@architect.co

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
