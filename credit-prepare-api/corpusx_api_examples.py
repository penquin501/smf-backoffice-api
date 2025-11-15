import requests

#------------------------------------------------------------------------#
#--                              Get Token                             --#
#------------------------------------------------------------------------#

params = {
            "grant_type":"password",
            "username": "xxxxx",        # Required
            "password": "xxxxx"         # Required
         }

url_token = "https://corpusxapi.bol.co.th/api/v1/token/token"
r_token = requests.post(url_token, data = params)

if r_token.status_code == 200:
    print("# Response status code: " + str(r_token.status_code) + " Token retrieval successful #")
else:
    print("# Response status code: " + str(r_token.status_code) + " Token retrieval unsuccessful #")
  
my_token = r_token.json()

#------------------------------------------------------------------------#
#--                           Refresh Token                            --#
#------------------------------------------------------------------------#

params_re_token = {
                    "grant_type":"refresh_token",
                    "refresh_token" : my_token['access_token']       # my_token
                  }

url_re_token = "https://corpusxapi.bol.co.th/api/v1/token/token"
r_re_token = requests.post(url_re_token, data = params_re_token)

if r_re_token.status_code == 200:
    print("# Response status code: " + str(r_re_token.status_code) + " Refresh token successful #")
else:
    print("# Response status code: " + str(r_re_token.status_code) + " Refresh token unsuccessful #")

re_token = r_re_token.json()

#------------------------------------------------------------------------#
#--                              Clear Token                           --#
#------------------------------------------------------------------------#

params_clear = {
                "grant_type":"password",
                "username": "xxxxx",        # Required
                "password": "xxxxx"         # Required
                }

url_clear = "https://corpusxbackapi.bol.co.th/api/v1/session/clear"

r_clear= requests.post(url_clear, data = params_clear)

if r_clear.status_code == 200:
    print("# Response status code: " + str(r_clear.status_code) + " Clear token successful #")
else:
    print("# Response status code: " + str(r_clear.status_code) + " Clear token unsuccessful #")

#------------------------------------------------------------------------#
#--                              Check Data                            --#
#------------------------------------------------------------------------#

params_check_data = {
                        "systemId": "1",
                        "registrationId": "0107546000407",      # Required either companyName or registrationId
                        "companyName": "",                      # Required either companyName or registrationId
                        "language": "",       
                        "status": "",
                        "fsType": "",
                        "dataSet": ""
                    }

head = {"Authorization": "Bearer {}".format(my_token['access_token'])}  # my_token
url_check_data = "https://corpusxfrontapi.bol.co.th/ApiFrontend/bol_service/check/data"

r_check_dat = requests.post(url_check_data, data = params_check_data, headers=head)

if r_check_dat.status_code == 200:
    print("# Response status code: " + str(r_check_dat.status_code) + " Check data successful #")
else:
    print("# Response status code: " + str(r_check_dat.status_code) + " Check data unsuccessful #")

check_data = r_check_dat.json()

#------------------------------------------------------------------------#
#--                              Check Cost                            --#
#------------------------------------------------------------------------#

params_check_cost = {
                        "systemId": "1"
                        , "registrationId": "0107546000407"                 # Required either companyName or registrationId
                        , "companyName":""                                  # Required either companyName or registrationId
                        , "status":""
                        , "dataSet":"7"                                     # Required either dataSet or dataField
                        , "dataField":"201010400"                           # Required either dataSet or dataField -- Example of single field
                        # , "dataField":"201010400,201010600,201010900"     # Example of multiple fields
                        , "periodFrom":"2017"                               # Required
                        , "periodTo":"2018"                                 # Required
                        , "fsType":""
                        , "language":""
                    }

head = {"Authorization": "Bearer {}".format(my_token['access_token'])}     # my_token
url_check_cost = "https://corpusxfrontapi.bol.co.th/ApiFrontend/bol_service/check/cost"

r_check_cost = requests.post(url_check_cost, data = params_check_cost, headers=head)

if r_check_cost.status_code == 200:
    print("# Response status code: " + str(r_check_cost.status_code) + " Check cost successful #")
else:
    print("# Response status code: " + str(r_check_cost.status_code) + " Check cost unsuccessful #")

check_cost = r_check_cost.json()

#------------------------------------------------------------------------#
#--                              Get Data                              --#
#------------------------------------------------------------------------#

params_get_data = {
                    "systemId": "1"
                    , "registrationId": "0107546000407"                 # Required either companyName or registrationId
                    , "companyName":""                                  # Required either companyName or registrationId
                    , "status":""
                    , "dataSet":"7"                                     # Required either dataSet or dataField
                    , "dataField":"201010400"                           # Required either dataSet or dataField -- Example of single field
                    # , "dataField":"201010400,201010600,201010900"     # Example of multiple fields
                    , "periodFrom":"2017"                               # Required
                    , "periodTo":"2018"                                 # Required
                    , "fsType":""
                    , "language":""
                  }

head = {"Authorization": "Bearer {}".format(my_token['access_token'])}   # my_token
url_get_data = "https://corpusxfrontapi.bol.co.th/ApiFrontend/bol_service/get/data"
r_get_data = requests.post(url_get_data, data = params_get_data, headers=head)

if r_get_data.status_code == 200:
    print("# Response status code: " + str(r_get_data.status_code) + " Get data successful #")
else:
    print("# Response status code: " + str(r_get_data.status_code) + " Get data unsuccessful #")

data = r_get_data.json()

