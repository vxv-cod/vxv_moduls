import uvicorn
import json


from typing import Any

import orjson
from fastapi import FastAPI, Response
from fastapi.responses import ORJSONResponse

app = FastAPI()

# Response.media_type = "application/json"


class CustomORJSONResponse(Response):
    media_type = "application/json"

    def render(self, content: Any) -> bytes:
        # assert orjson is not None, "orjson must be installed"
        # return orjson.dumps(content, option=orjson.OPT_INDENT_2)
        return json.dumps(
                content,
                ensure_ascii=False,
                allow_nan=False,
                indent=4,
                separators=(",", ":"),
            ).encode("utf-8")
        
        # return ORJSONResponse([json.dumps(content, ensure_ascii = False, indent=16, sort_keys=True)])


@app.get("/1", response_class=CustomORJSONResponse)
async def main():
    return {"item_id": "response_class"}


@app.get("/")
async def main():
    # Response.media_type = "application/json"
    content = {"message": "Без ссылки на response_class"}
    # response = Response(orjson.dumps(content, option=orjson.OPT_INDENT_2))
    response = Response(json.dumps(content, ensure_ascii = False, indent=5, sort_keys=True))
    return response


if __name__ == "__main__":
    uvicorn.run("test_Response:app", host="localhost", port=8000, reload=True, workers=3)
    # uvicorn.run(app, host="localhost", port=8000)