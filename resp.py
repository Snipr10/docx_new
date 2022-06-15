
import logging.config

logger = logging.getLogger('foo-logger')
TIMEOUT = 7 * 60


async def post(session, url, body):
    attempt = 0
    response = None
    while attempt < 5:
        try:
            if body:
                response = await session.post(url, json=body, timeout=TIMEOUT)
            else:
                response = await session.post(url, timeout=TIMEOUT)
        except Exception as e:
            try:
                logger.error(f"post {e} {response.text} {response.status_code}")
            except Exception as e:
                logger.error(f"post {e}")
        attempt += 1
        if response is not None and response.status_code == 200:
            break
    return response
