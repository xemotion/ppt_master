import os
import sys
import requests
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import html2text
import re
import urllib.parse

# .env 파일에서 환경변수 로드
env_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path=env_path)

CONFLUENCE_BASE_URL = os.getenv('CONFLUENCE_BASE_URL')
CONFLUENCE_API_USER = os.getenv('CONFLUENCE_API_USER')
CONFLUENCE_API_PASSWORD = os.getenv('CONFLUENCE_API_PASSWORD')

if not all([CONFLUENCE_BASE_URL, CONFLUENCE_API_USER, CONFLUENCE_API_PASSWORD]):
    raise ValueError('환경변수(.env) 설정을 확인하세요.')

def extract_page_id_from_url(url):
    """Confluence 페이지 URL에서 page ID 추출. 없으면 None 반환."""
    # 예시: .../pages/123456789/...
    match = re.search(r'/pages/(\d+)', url)
    if match:
        return match.group(1)
    # 예시: ...pageId=123456789
    match = re.search(r'pageId=(\d+)', url)
    if match:
        return match.group(1)
    return None

def extract_spacekey_title_from_url(url):
    """/display/{spaceKey}/{title} 패턴에서 spaceKey, title 추출"""
    match = re.search(r'/display/([^/]+)/([^/?#]+)', url)
    if match:
        spacekey = match.group(1)
        title = match.group(2)
        # URL 디코딩
        title = urllib.parse.unquote(title)
        return spacekey, title
    return None, None

def get_page_id_by_title(spacekey, title):
    """spaceKey와 title로 pageId 조회 (+를 공백으로 변환해서 우선 시도, 실패 시 + 그대로도 시도)"""
    # 1차 시도: +를 공백으로 변환
    title_for_query = title.replace('+', ' ')
    api_url = f"{CONFLUENCE_BASE_URL}/rest/api/content?title={urllib.parse.quote(title_for_query)}&spaceKey={spacekey}&expand=history"
    response = requests.get(
        api_url,
        auth=(CONFLUENCE_API_USER, CONFLUENCE_API_PASSWORD)
    )
    if response.status_code == 200:
        data = response.json()
        if 'results' in data and len(data['results']) > 0:
            return data['results'][0]['id']
    # 2차 시도: +를 그대로 두고 시도
    api_url = f"{CONFLUENCE_BASE_URL}/rest/api/content?title={urllib.parse.quote(title)}&spaceKey={spacekey}&expand=history"
    response = requests.get(
        api_url,
        auth=(CONFLUENCE_API_USER, CONFLUENCE_API_PASSWORD)
    )
    if response.status_code == 200:
        data = response.json()
        if 'results' in data and len(data['results']) > 0:
            return data['results'][0]['id']
    raise ValueError('title/spaceKey로 pageId를 찾을 수 없습니다.')

def save_confluence_page(page_id, output_format='txt', include_images=False):
    api_url = f"{CONFLUENCE_BASE_URL}/rest/api/content/{page_id}?expand=body.view"
    response = requests.get(
        api_url,
        auth=(CONFLUENCE_API_USER, CONFLUENCE_API_PASSWORD)
    )

    if response.status_code != 200:
        raise Exception(f"Confluence API 요청 실패: {response.status_code} {response.text}")

    content = response.json()
    html_body = content['body']['view']['value']

    if output_format == 'txt':
        soup = BeautifulSoup(html_body, 'html.parser')
        output_text = soup.get_text(separator='\n', strip=True)
        output_file = f"confluence_page_{page_id}.txt"
    elif output_format == 'md':
        soup = BeautifulSoup(html_body, 'html.parser')
        if include_images:
            for img in soup.find_all('img'):
                alt = img.get('alt', '')
                src = img.get('src', '')
                md_img = f'![{alt}]({src})'
                img.replace_with(md_img)
            for tag in soup.find_all(['iframe', 'object', 'embed']):
                src = tag.get('src') or tag.get('data') or ''
                if src:
                    md_img = f'![gliffy]({src})'
                    tag.replace_with(md_img)
            output_text = html2text.html2text(str(soup))
        else:
            output_text = html2text.html2text(html_body)
        output_file = f"confluence_page_{page_id}.md"
    else:
        raise ValueError('output_format은 "txt" 또는 "md"만 지원합니다.')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(output_text)

    print(f"{output_format.upper()} 파일로 저장 완료: {output_file}")

if __name__ == "__main__":
    # 사용법: python confluence_to_md.py <confluence_page_url> [md|txt] [include_images]
    if len(sys.argv) < 2:
        print('사용법: python confluence_to_md.py <confluence_page_url> [md|txt] [include_images]')
        sys.exit(1)
    page_url = sys.argv[1]
    output_format = sys.argv[2] if len(sys.argv) > 2 else 'txt'
    include_images = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else False
    page_id = extract_page_id_from_url(page_url)
    if not page_id:
        spacekey, title = extract_spacekey_title_from_url(page_url)
        if not (spacekey and title):
            print('URL에서 pageId 또는 spaceKey/title을 찾을 수 없습니다.')
            sys.exit(1)
        page_id = get_page_id_by_title(spacekey, title)
    save_confluence_page(page_id, output_format=output_format, include_images=include_images) 
