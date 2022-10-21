import logging.config

from settings import GET_ADDITIONAL_INFO, GET_STATS, GET_CITY, GET_OWNER_TOP, \
    GET_POST_TOP, GET_AGES, GET_TRUST_DAILY, GET_RESOURCES_STATS
from resp import post

logger = logging.getLogger('foo-logger')
TIMEOUT = 7 * 60


async def get_additional_info(session, thread_id):
    response = await post(session, GET_ADDITIONAL_INFO,
                          {
                              "thread_id": thread_id,

                          })
    return response.json()


async def get_stats(session, thread_id, _from, _to, type):
    response = await post(session, GET_STATS,
                          {
                              "thread_id": thread_id,
                              "from": _from,
                              "to": _to,
                              'type': type
                          })
    return response.json()


async def get_city(session, thread_id, _from, _to):
    response = await post(session, GET_CITY,
                          {
                              "thread_id": thread_id,
                              "from": _from,
                              "to": _to,
                              "limit": 100

                          })
    return response.json()


async def get_top(session, thread_id, _from, _to, type):
    response = await post(session, GET_OWNER_TOP,
                          {
                              "thread_id": thread_id,
                              "referenceFilter": [],
                              "from": _from,
                              "to": _to,
                              "type": type

                          })
    return response.json()


async def get_post_top(session, thread_id, _from, _to):
    response = await post(session, GET_POST_TOP,
                          {
                              "thread_id": thread_id,
                              "negative": None,
                              "from": _from,
                              "to": _to,
                              "post_count": 10

                          })
    return response.json()


async def get_ages(session, thread_id, _from, _to):
    response = await post(session, GET_AGES,
                          {
                              "thread_id": thread_id,
                              "from": _from,
                              "to": _to,
                              "group1_start": 18,
                              "group1_end": 25,
                              "group2_start": 26,
                              "group2_end": 39,
                              "group3_start": 40,
                              "group3_end": 54,
                              "group4_start": 55,
                              "group4_end": 170
                          })
    return response.json()


async def get_trustdaily(session, thread_id, _from, _to):
    response = await post(session, GET_TRUST_DAILY,
                          {
                              "thread_id": thread_id,
                              "from": _from,
                              "to": _to
                          })
    res_dict = {}
    TOTAL = "total"
    total_count = 0
    for trust in response.json():
        if trust != TOTAL:
            trust_dict = {}
            total_sm_count = 0
            for s in response.json().get(trust):
                if s != TOTAL:
                    count = 0
                    for c in response.json().get(trust).get(s):
                        count += c[-1]
                    trust_dict[s] = count
                    total_count += count
                    total_sm_count += count
            trust_dict[TOTAL] = total_sm_count

            res_dict[trust] = trust_dict

    return total_count, res_dict, response.json()


async def getSourceStats(session, thread_id, _from, _to):
    response = await post(session, GET_RESOURCES_STATS,
                          {
                              "thread_id": thread_id,
                              "from": _from,
                              "to": _to
                          })

    return response.json()
