import sys
from loguru import logger
import urllib3

from response_407 import requests_session

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


logger.trace("A trace message.")
logger.debug("A debug message.")
logger.info("An info message.")
logger.success("A success message.")
logger.warning("A warning message.")
logger.error("An error message.")
logger.critical("A critical message.")

list_type = [
    "TRACE",
    "DEBUG",
    "INFO",
    "SUCCESS",
    "WARNING",
    "ERROR",
    "CRITICAL"
]


# sys.stderr
# sys.stdout
# "out.log", enqueue=True,

logger.remove()
logger.add(
    sys.stderr, 
    # "out.log", enqueue=True,
    format = str(
        "<green>{time:D.MM.YYYY - HH:mm:SSSS!UTC}</green> | "
        "<level>{level: <8}</level> | "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - "
        "<level>{message}</level> | "
        "<red>{extra}</red>"
    ),
    # backtrace=False,
    # diagnose=False,
    # level="SUCCESS",
)

extra = {"context": "foo"}
logger.configure(extra=extra)

# log_1 = logger.remove()

# logger.configure(
#     handlers=[
#         # dict(sink=sys.stderr, format="[{time}] {message}"),
#         dict(sink=sys.stderr, format="<green>{time:D.MM.YYYY - HH:mm:ssss}</> | <level>{level} | {message}</>"),
#         dict(sink="file.log", enqueue=True, serialize=True),
#     ],
#     levels=[dict(name="NEW", no=13, icon="¤", color="")],
#     extra={"common_to_all": "default"},
#     patcher=lambda record: record["extra"].update(some_value=42),
#     activation=[("my_module.secret", False), ("another_library.module", True)],
# )



requests_session("http://127.0.0.1:8888/tasks")


'''Если вам не нужно включать в запись журнала все, 
вы можете создать собственную serialize()функцию и использовать ее 
следующим образом:'''

'''https://betterstack.com/community/guides/logging/loguru/#formatting-log-records'''

# import sys
# import json
# from loguru import logger


# def serialize(record):
#     subset = {
#         "timestamp": record["time"].timestamp(),
#         "message": record["message"],
#         "level": record["level"].name,
#     }
#     # return json.dumps(subset)
#     return json.dumps(subset, ensure_ascii = False, indent=4, sort_keys=True)

# def patching(record):
#     record["extra"]["serialized"] = serialize(record)
# logger.remove(0)

# logger = logger.patch(patching)
# logger.add(sys.stderr, format="{extra[serialized]}")
# logger.debug("Happy logging with Loguru!")

