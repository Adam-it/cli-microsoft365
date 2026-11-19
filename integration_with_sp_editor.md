# Integration with SP Editor

## Overview

The **CLI for Microsoft 365** can be integrated with the SP Editor to run commands directly in the browser, using the context of the signed-in SharePoint user. This allows for executing SharePoint (`spo`) commands without leaving the SP Editor environment. 

> **Note:** This integration is experimental and intended for proof-of-concept scenarios.

## Build

`@pnp/cli-microsoft365` is not installed from npm but from a local clone next to this folder (`file:../cli-microsoft365`).
Before you start building and testing it you need to have both CLI for Microsoft 365 and SP Editor repo checked out locally next to each other.

For example, your folder structure should look like this:

```
/src
  /cli-microsoft365
  /sp-editor
```

For SP Editor you need to checkout the [sp-editor-with-cli-m365](https://github.com/Adam-it/sp-editor/tree/sp-editor-with-cli-m365) branch.

Once you have both repositories checked out, you can perform the build running the below steps/script:

```powershell
cd ../cli-microsoft365 # the branch with the browser bundle
npm i
npm run build:browser # builds dist-browser
cd ../sp-editor
npm i
npm run build # copies the bundle from node_modules to public/bundles/m365 first
```

`npm run m365cli` copies the bundle again after the CLI was rebuilt; reload the extension afterwards.

## How to test the extension

1. In Edge, go to `edge://extensions`, turn on Developer mode, choose "Load unpacked" and pick `sp-editor\build` location. Disable the store version of SP Editor if you have it installed.
2. Open a SharePoint site, open DevTools, go to the "SharePoint" panel and pick "CLI for Microsoft 365".
3. Run for e.g. `spo web get --url <site>`, then a write such as `spo list add --title CliTest --baseTemplate GenericList --webUrl <site>` etc.