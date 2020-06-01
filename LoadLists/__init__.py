import logging

import os
import json
import azure.functions as func
from shareplum import Site
from shareplum.site import Version
from shareplum import Office365

def main(req: func.HttpRequest) -> func.HttpResponse:
    # Configuration Sanity Check
    o365HostUri = os.environ['O365HostUri']
    o365UserName = os.environ['O365UserName']
    o365UserPassword = os.environ['O365UserPassword']
    o365SiteUri = os.environ['O365SiteUri']
    o365ListName = os.environ['O365ListName']

    validationError = False
    if not o365HostUri:
        validationError = True
        logging.error('O365HostUri configuration missing.')
    if not o365UserName:
        validationError = True
        logging.error('O365UserName configuration missing.')
    if not o365UserPassword:
        validationError = True
        logging.error('O365UserPassword configuration missing.')
    if not o365SiteUri:
        validationError = True
        logging.error('O365SiteUri configuration missing.')
    if not o365ListName:
        validationError = True
        logging.error('O365ListName configuration missing.')

    if validationError:
        return func.HttpResponse("Invalid configuration.", status_code=400)

    authcookie = Office365(o365HostUri, username=o365UserName, password=o365UserPassword).GetCookies()
    site = Site(o365SiteUri, version=Version.v365, authcookie=authcookie)
    sp_list = site.List(o365ListName)
    data = sp_list.GetListItems('All Items')
    logging.info(data)

    return func.HttpResponse(json.dumps(data), status_code=200, mimetype='application/json')
