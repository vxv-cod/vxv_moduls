import os
import sys
import json
import httpx
from loguru import logger
from requests_negotiate_sspi import HttpNegotiateAuth
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def decorTime(my_func):
    '''Обертка функции декоратором'''
    from time import time
    # *args — это сокращение от «arguments» (аргументы) в виде кортежа, 
    # **kwargs — сокращение от «keyword arguments» (именованные аргументы)
    @logger.catch
    def wrapper(*args):
        start_time = time()
        my_func(*args)
        wrapper_time = f'"Выполнено за: {round(time() - start_time, 8)} sec"'
        logger.success(wrapper_time)

    return wrapper



# @logger.catch
@decorTime
def requests_session(url):
    '''Убираем ошибку о не проверенной ссылке'''
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    '''Вместо session.trust_env = False можно использовать следующую строчку'''
    os.environ['no_proxy'] = '*' 

    with requests.Session() as session:
        # session.trust_env = False
        # session.auth = HttpNegotiateAuth()
        # session.verify = False
        # session.max_redirects = 5
        
        params = {
            "url": url, 
            "verify": False, 
            "auth": HttpNegotiateAuth(), 
            # "timeout": 5
            }
        
        # response = session.get(url)
        response = session.get(**params)
                
        if response.status_code != 200:
            logger.error(url, response)
            return 
        response = response.json()
        response = json.dumps(response, ensure_ascii = False, indent=4, sort_keys=True)
        logger.debug(response)

    return response




def requests_get(url):
    os.environ['no_proxy'] = '*' 
    params = {
        "url": url, 
        "verify": False, 
        "auth": HttpNegotiateAuth(), 
        "timeout": 5
        }
    response = requests.get(**params)
    '''Альтернативный запрос'''
    # response = requests.request("GET", **params)

    if response.status_code != 200:
        logger.error(response)
        return 
    response = response.json()

    # if response.headers['content-type'].split("; ")[0] == "application/json":
    #     response = response.json()
    # else:
    #     response = response.text

    response = json.dumps(response, ensure_ascii = False, indent=4, sort_keys=True)
    logger.debug(response)
    
    return response



if __name__ == "__main__":

    url = "http://127.0.0.1:8888/tasks"
    requests_session(url)
    
    url = "https://tnnc-pir-app.rosneft.ru/test-api/view-customer-plans"
    requests_get(url)




# proxy_string = r'http://tmn-tnnc-proxy.rosneft.ru:9090'
# proxies = {
#   'https' : proxy_string,
#   'http' : proxy_string,
# }