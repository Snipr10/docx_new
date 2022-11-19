import logging
import os

from dotenv import load_dotenv


load_dotenv()
load_dotenv(".version")


DOMAIN_URL = os.getenv("DOMAIN_URL", "https://api.glassen-it.com")
login_l = "java_api"
password_p = "4yEcwVnjEH7D"

if DOMAIN_URL == "https://api.glassen-it.com":
    NETWORK_IDS = [1, 2, 4, 5, 8, 9, 10]
else:
    NETWORK_IDS = [1, 2, 3, 4, 5, 7, 8, 9, 10]

if DOMAIN_URL == "https://isiao.glassen-it.com":
    login_l = "superadmin"
    password_p = "superadmin"

LOGIN_URL = os.getenv("LOGIN_URL", f"{DOMAIN_URL}/component/socparser/authorization/login")
SUBECT_URL = os.getenv("SUBECT_URL", f"{DOMAIN_URL}/component/socparser/users/getreferences")
STATISTIC_POST_URL = os.getenv("STATISTIC_POST_URL", f"{DOMAIN_URL}/component/socparser/content/posts")

SUBECT_TOPIC_URL = os.getenv("SUBECT_TOPIC_URL", f"{DOMAIN_URL}/component/socparser/stats/getMainTopics")
STATISTIC_URL = os.getenv("STATISTIC_URL", f"{DOMAIN_URL}/component/socparser/content/getpostcount")
STATISTIC_TRUST_GRAPH = os.getenv("STATISTIC_TRUST_GRAPH", f"{DOMAIN_URL}/component/socparser/stats/trustViewsGraph")
THREAD_URL = os.getenv("THREAD_URL", f"{DOMAIN_URL}/component/socparser/threads/get")
GET_TRUST_URL = os.getenv("GET_TRUST_URL", f"{DOMAIN_URL}/component/socparser/content/getposttoptrust")
GET_ATTENDANCE_URL = os.getenv("GET_ATTENDANCE_URL", f"{DOMAIN_URL}/component/socparser/stats/getpostattendance")

WEEK_TRUST = os.getenv("WEEK_TRUST", f"{DOMAIN_URL}/component/socparser/content/getweektrust")
THREAD_STATS = os.getenv("THREAD_STATS", f"{DOMAIN_URL}/component/socparser/stats/getThreadStats")
THREAD_DATA = os.getenv("THREAD_DATA", f"{DOMAIN_URL}/component/socparser/thread/additional_info")
OWNERS_TOP = os.getenv("OWNERS_TOP", f"{DOMAIN_URL}/component/socparser/stats/getOwnersTopByPostCount")

GET_RESOURCES_STATS = f"{DOMAIN_URL}/component/socparser/stats/getsourcestats"
GET_ADDITIONAL_INFO = f"{DOMAIN_URL}/component/socparser/thread/additional_info"
GET_TRUST_DAILY = f"{DOMAIN_URL}/component/socparser/stats/trustdaily"
GET_STATS = f"{DOMAIN_URL}/component/socparser/stats"
GET_AGES = f"{DOMAIN_URL}/component/socparser/stats/ages"
GET_CITY = f"{DOMAIN_URL}/component/socparser/thread/getcitytop"
GET_OWNER_TOP = f"{DOMAIN_URL}/component/socparser/stats/getOwnersTopByPostCount"
GET_POST_TOP = f"{DOMAIN_URL}/component/socparser/content/getposttoptrust"

def print_settings(logger: logging.Logger) -> None:
    # Note: Do not log sensitive information here like username, password, and api keys and secrets.
    logger.info("Python: " + "; ".join([
        f"{DOMAIN_URL}",
        f"{LOGIN_URL}",
        f"{SUBECT_URL}",
        f"{STATISTIC_POST_URL}",
        f"{SUBECT_TOPIC_URL}",
        f"{STATISTIC_URL}",
        f"{STATISTIC_TRUST_GRAPH}",
        f"{THREAD_URL}",
        f"{GET_TRUST_URL}",
        f"{GET_ATTENDANCE_URL}",
    ]))
