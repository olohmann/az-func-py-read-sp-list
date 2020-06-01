# For Local Testing

Make sure you have installed the latest [https://github.com/Azure/azure-functions-core-tools](Azure Function Core Tools).

Place a `local.settings.json` file with the following format in the root directory:
```
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "",
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "O365HostUri": "https://oliverlo365dev.sharepoint.com",
    "O365UserName": "sp@oliverlo365dev.onmicrosoft.com",
    "O365UserPassword": "***",
    "O365SiteUri": "https://oliverlo365dev.sharepoint.com/sites/demo/",
    "O365ListName": "sample-list"
  }
}
```

Make sure that you replace the values with your O365 environment.

Issue `func start` and perform a `curl` request against the local Azure Function endpoint.
