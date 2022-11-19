
import logging.config

import httpx

logger = logging.getLogger('foo-logger')
TIMEOUT = 7 * 60


async def post(session, url, body):
    this_session = session
    attempt = 0
    response = None
    while attempt < 5:
        try:
            if body:
                response = await this_session.post(url, json=body, timeout=TIMEOUT)
            else:
                response = await this_session.post(url, timeout=TIMEOUT)
        except Exception as e:
            try:
                logger.error(f"post {e} {response.text} {response.status_code}")
            except Exception as e:
                logger.error(f"post {e}")
        try:
            if response.status_code == 403:
                response = None
                from word_media import login
                this_session = httpx.AsyncClient()
                await login(this_session)
        except Exception:
            pass
        attempt += 1
        if response is not None and response.status_code == 200:
            break
    return response
