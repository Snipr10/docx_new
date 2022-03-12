import asyncio
import re
from datetime import datetime

import docx
import httpx
from docx.oxml.shared import OxmlElement
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt, RGBColor

from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
CLEANR = re.compile('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});')

TIMEOUT = 15 * 60
LOGIN_URL = "https://api.glassen-it.com/component/socparser/authorization/login"
SUBECT_URL = "https://api.glassen-it.com/component/socparser/users/getreferences"
STATISTIC_POST_URL = "https://api.glassen-it.com/component/socparser/content/posts"
STYLE = "Times New Roman"
PT = Pt(10.5)


async def login(session, login="java_api", password="4yEcwVnjEH7D"):
    payload = {
        "login": login,
        "password": password
    }
    response = await session.post(LOGIN_URL, json=payload, timeout=TIMEOUT)
    if response.status_code != 200:
        raise Exception("can not login")
    return session


# Press the green button in the gutter to run the script.

def add_title_data(title, name, data):
    title_run = title.add_run(name, style=STYLE)
    title_run.font.color.rgb = RGBColor(118, 113, 113)
    title_run.bold = True
    title_run.font.size = Pt(12)
    title_run = title.add_run(f": {data}", style=STYLE)
    title_run.italic = True
    title_run.font.size = Pt(12)


async def subects_names(session, referenceFilter):
    response = await session.post(SUBECT_URL, timeout=TIMEOUT)
    names = []
    try:
        res = []
        for r in response.json():
            res.extend(r['items'] or [])
        for r in res:
            if r.get("id") in referenceFilter:
                names.append(r.get("keyword"))
        return names
    except Exception:
        return []


async def get_posts(session, thread_id, _from, _to, network_id, referenceFilter):
    limit = 200
    start = 0
    posts = []
    while True:
        payload = {
            "thread_id": thread_id,
            "from": _from,
            "to": _to,
            "limit": limit, "start": start, "sort": {"type": "date", "order": "desc", "name": "dateDown"},
            "filter": {"network_id": network_id,
                       "referenceFilter": referenceFilter, "repostoption": "whatever"}
        }
        response = await session.post(STATISTIC_POST_URL, json=payload, timeout=TIMEOUT)
        posts.extend(response.json().get("posts") or [])
        if not response.json().get("posts") or response.json().get("count") <= len(posts):
            break
        start += limit
    smi = 0
    social = 0
    for post in posts:
        if post.get("network_id") == "4":
            smi += 1
        else:
            social += 1
    return posts, smi, social


async def get_session_result(_login, _password, thread_id, _from, _to, referenceFilter, network_id):
    async with httpx.AsyncClient() as session:
        session = await login(session)

        (posts, smi, social), names = await asyncio.gather(
            get_posts(session, thread_id, _from, _to, network_id, referenceFilter),
            subects_names(session, referenceFilter)
        )
        return posts, smi, social, names


def insertHR(paragraph, line = "double"):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), line)
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)


def add_hyperlink_into_run(paragraph, run, url):
    runs = paragraph.runs
    for i in range(len(runs)):
        if runs[i].text == run.text:
            break
    # --- This gets access to the document.xml.rels file and gets a new relation id value ---
    part = paragraph.part
    r_id = part.relate_to(
        url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )
    # --- Create the w:hyperlink tag and add needed values ---
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
    hyperlink.append(run._r)
    paragraph._p.insert(i,hyperlink)
    run.font.color.rgb = docx.shared.RGBColor(0, 0, 255)


async def docx_media(login, password, thread_id, _from, _to, referenceFilter, network_id):
    print("docx_media")
    posts, smi, social, names = await get_session_result(login, password, thread_id, _from, _to, referenceFilter, network_id)
    document = Document()
    obj_styles = document.styles
    obj_charstyle = obj_styles.add_style(STYLE, WD_STYLE_TYPE.CHARACTER)
    obj_font = obj_charstyle.font
    obj_font.name = STYLE

    title = document.add_paragraph()
    title_run = title.add_run("СВОДКА ПУБЛИКАЦИЙ", style=STYLE)
    title_run.font.color.rgb = RGBColor(118, 113, 113)
    title_run.bold = True
    title_run.font.size = Pt(14)

    title = document.add_paragraph()

    if len(network_id) >= 2 and 4 in network_id:
        type_network = "СМИ и соцсети"
        number_networks = f"СМИ - {smi} публикаций, соцсети - {social} публикаций"
    elif 4 in network_id:
        type_network = "только СМИ"
        number_networks = f"СМИ - {smi} публикаций"
    else:
        type_network = "только соцсети"
        number_networks = f"соцсети - {social} публикаций"

    add_title_data(title, "Объекты", ", ".join(names))
    add_title_data(title, "\nИсточники публикаций", type_network)
    add_title_data(title, "\nВременной период", f"{_from} - {_to}")
    add_title_data(title, "\nДата подготовки отчета", datetime.today().strftime("%m/%d/%Y, %H:%M:%S"))
    add_title_data(title, f"\nВсего сообщений", number_networks)
    insertHR(document.add_paragraph(), line="single")

    for post in posts:
        post_paragraph = document.add_paragraph()
        paragraph_run = post_paragraph.add_run(post['author'], style=STYLE)
        paragraph_run.bold = True
        paragraph_run.font.size = Pt(12)
        paragraph_run = post_paragraph.add_run(f" {post['created_date']}"
                                               "\n ", style=STYLE)
        paragraph_run.font.size = Pt(12)

        paragraph_ilnk = document.add_paragraph()
        url = post['uri']
        if len(url) > 65:
            url = url[:65]
        paragraph_link = paragraph_ilnk.add_run(url, style=STYLE)
        paragraph_link.font.size = Pt(12)

        add_hyperlink_into_run(paragraph_ilnk, paragraph_link, post['uri'])

        if post['title']:
            post_paragraph_title = document.add_paragraph()
            paragraph_title = post_paragraph_title.add_run(post['title'], style=STYLE)
            paragraph_title.bold = True
            paragraph_title.font.size = Pt(12)

        post_paragraph_text = document.add_paragraph()
        cleantext = re.sub(CLEANR, '', post['text'])
        post_paragraph_text = post_paragraph_text.add_run(cleantext, style=STYLE)
        post_paragraph_text.font.size = Pt(12)

        hr = document.add_paragraph()
        hr.style.font.size = Pt(1)

        insertHR(hr)

    return document

