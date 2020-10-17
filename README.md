## Overview

This project creates a chrome extension which allows for sample data population into lists. The extension can be viewed on the [chrome web store](https://chrome.google.com/webstore/detail/sharepoint-population-too/eaoeogggkmgpmbigihinlmjhefpeaefa?hl=en&authuser=0).

## Development Testing

This project was bootstrapped with [Create React App](https://github.com/facebook/create-react-app).

### Getting up and running in Chrome

```
cd .\src\popup
npm install
cd ..\..\
npm install
npm run clean
npm run watch
```
Open chrome -> ... -> More Tools -> Extensions -> Load Unpacked -> Select dist directory created from watch command

### Test Platform

test\templates\testingGrounds.xml is a [PnP Template](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/introducing-the-pnp-provisioning-engine) that can be applied on an existing site for configuring a wide range of site columns/content types to test.

``` ps
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/targetcommunicationsite"
Apply-PnPProvisioningTemplate -Path "./test/templates/testingGrounds.xml"
```

## Production Build

```
npm run clean
npm run build
npm run zip
```