import asyncio
import json
import os
import sys
import httpx
import uvicorn
import orjson
from fastapi.responses import ORJSONResponse, StreamingResponse
from typing import Any
from fastapi import FastAPI, Response
from fastapi.responses import RedirectResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from a2wsgi import ASGIMiddleware
from loguru import logger
from requests_negotiate_sspi import HttpNegotiateAuth
from httpx_negotiate_sspi import HttpSspiAuth

import requests
import urllib3
from tqdm import tqdm
import aiohttp
from time import time
from rich.progress import track, Progress


import urllib

# os.environ['no_proxy'] = '*' 
'''Убираем ошибку о не проверенной ссылке'''
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# headers={"User-Agent": "Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148"}
# headers={'User-Agent': 'httpx/0.19.0'}


from time import time

def timer_(my_func):
    '''Время выполнения'''
    async def wrapper(*args):
        start_time = time()
        # res =  await my_func(track(*args))
        res =  await my_func(*args)
        logger.success(f'Итого : {round(time() - start_time, 8)} sec')
        return res
    return wrapper



# "Время": "\"0.03124642 sec\""
# @logger.catch
# @timer_
def requests_session(urls):
    content = []
    with requests.Session() as session:
        # session.verify = False
        session.trust_env = False
        # session.auth = HttpNegotiateAuth()
        
        for url in urls:
            response = session.get(url)
            if response.status_code == 200:
                content.append(response.json())
            else:
                content.append({url : response.status_code})

    return content


# proxies = urllib.request.getproxies()
# proxies = {
#     'http': 'http://tmn-tnnc-proxy.rosneft.ru:9090', 
#     'https': 'http://tmn-tnnc-proxy.rosneft.ru:9090', 
#     'ftp': 'http://tmn-tnnc-proxy.rosneft.ru:9090'
# }


async def async_http_get_aiohttp(urls: list):
    async with aiohttp.ClientSession() as session:
        content = []
        for url in urls:
            async with session.get(url=url) as response:
                content.append(await response.json())

        return content


@logger.catch
async def async_http_get_aiohttp_gather(urls: list):
    async with aiohttp.ClientSession() as session:
        tasks = []
        for url in urls:
            tasks.append(asyncio.create_task(session.get(url)))
        responses = await asyncio.gather(*tasks)
        return [await r.json() for r in responses]

# @timer_


CLIENT_PARAMS = {"verify" : False, "trust_env" : False, "auth" : HttpSspiAuth(),
        #   "headers" : {'Content-Type': 'application/json; charset=utf-8'}
            }

@timer_
async def async_httpx_AsyncClient_get(urls):
    async with httpx.AsyncClient(**CLIENT_PARAMS) as client:
        content = []
        for url in urls:
            obj = {"url" : url}
            response = await client.get(url)
            if response.status_code == 200:
                obj["data"] = response.json()
                content.append(obj)
            else:
                obj["status_code"] = response.status_code
                content.append(obj)                
                logger.error(f"{obj}")
    return content


@timer_
async def async_httpx_AsyncClient_gather_get(urls):
    async with httpx.AsyncClient(**CLIENT_PARAMS) as client:
        tasks = [client.get(url) for url in urls]
        response = await asyncio.gather(*tasks)
        # response = [item.json() for item in response]
        content = []
        for resp in response:
            obj = {"url" : str(resp.url)}
            if resp.status_code == 200:
                obj["data"] = resp.json()
                content.append(obj)
            else:
                obj["status_code"] = resp.status_code
                content.append(obj)                
                logger.error(f"{obj}")

    return content


# params = {"verify" : False, "trust_env" : False, "auth" : HttpSspiAuth()}
# async with httpx.AsyncClient(**params) as client:
#     content = {}
#     for url in urls:
#         req = client.build_request("GET", url)
#         response = await client.send(req, stream=True)
        
        # logger.error(response.json())
        # iter_text = response
        # logger.error(f"{iter_text}")
        # for text in  iter_text:            
        #     response = text
                
        # yield json.dumps(response.json()).encode("utf-8")
        # return StreamingResponse(response.aiter_text(), background=BackgroundTask(response.aclose))


async def async_httpx_AsyncClient_stream_GET(urls):
    # return StreamingResponse(async_httpx_AsyncClient_stream_GET)
    async with httpx.AsyncClient(**CLIENT_PARAMS) as client:
        for idx, url in enumerate(urls):
            async with client.stream('GET', url) as response:
                if response.status_code == 200:
                    async for line in response.aiter_lines():
                        # line = json.dumps(json.loads(line), ensure_ascii = False, indent=2)
                        yield f"query_{idx} : {line}\n"
                else:
                    line = response.status_code
                    text = f"query_{idx} : {response.status_code}\n"
                    logger.error(text)
                    yield f"query_{idx} : {line}\n"



async def fetch(client: httpx.AsyncClient, url: str):
    async with client.get(url) as r:
        if r.status_code != 200:
            r.status_code
            
            
        return await r.json()
    

async def fetch_all(s, urls):
    tasks = []
    for url in urls:
        task = asyncio.create_task(fetch(s, url))
        tasks.append(task)
    res = await asyncio.gather(*tasks)
    return res





if __name__ == "__main__":
    urls = [
        # "http://127.0.0.1:8888/tasks",
        # "http://127.0.0.1:8888/users",
        # "https://tnnc-pir-app.rosneft.ru/test-api/view-customer-plans",
        "https://jsonplaceholder.typicode.com/todos/1"
    ]
    urls = urls * 100
    # response = requests_session(urls)     # Время : ""0.12374616 sec"

    # func = async_http_get_aiohttp(urls)           # Время : "0.05312514 sec"
    # func = async_http_get_aiohttp_gather(urls)    # Время : "0.03124595 sec"
    
    # func =  async_httpx_AsyncClient_get(urls)
    
    # response = asyncio.run(func)
    func =  async_httpx_AsyncClient_gather_get(urls)
    asyncio.run(func)
    
    # logger.debug(response)
    # wrapper_time = f'"{round(time() - start_time, 8)} sec"'
    # logger.debug(f"Время : {wrapper_time}")

    # urls = [
    #         'https://jsonplaceholder.typicode.com/todos/1'
    #     ]
    # requests_session(urls)
    # logger.debug(requests_session(urls))