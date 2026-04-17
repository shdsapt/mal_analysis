#!/usr/bin/env python3
"""
EML Attachment Extractor
EML 파일에서 첨부파일을 자동으로 추출하여 각 EML 파일명별 폴더에 저장합니다.
Python 내장 email 모듈을 사용하므로 추가 패키지가 필요 없습니다.
"""

import os
import sys
import email
import email.policy
import hashlib
import re
import argparse
import html
from email.header import decode_header
from urllib.parse import urlparse

# 윈도우 cp949 인코딩 에러 방지 설정 (asyncio 충돌 방지를 위해 reconfigure 사용)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')



# file_analysis.py에서 지원하는 확장자 목록 (768~780 라인 기준)
SUPPORTED_EXTENSIONS = {
    'pdf', 
    'xls', 'xlsx', 'xlsm', 'doc', 'docx', 'docm',
    'png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'jfif',
    'exe', 'dll', 'sys', 'ocx',
    'ppt', 'pptx', 'pptm', 'potx', 'pps', 'ppsx'
}

# 기본 경로 설정 (스크립트 위치 기준)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EML_DIR = os.path.join(SCRIPT_DIR, "eml")
DEFAULT_OUTPUT_DIR = os.path.join(SCRIPT_DIR, "attachfiles")


def sanitize_filename(filename):
    """파일명에서 OS에서 허용하지 않는 문자를 제거합니다."""
    # Windows 금지 문자 제거
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    
    # 너무 긴 파일명/폴더명 자르기 (Windows MAX_PATH 제한 방지)
    if len(filename) > 80:
        name, ext = os.path.splitext(filename)
        # 자른 후 즉시 공백을 제거하여 "폴더명 " 방지
        filename = name[:80].strip() + ext
        
    # 최종적으로 앞뒤 공백 및 점 제거
    return filename.strip(' .')


def decode_mime_header(header_value):
    """MIME 인코딩된 헤더 값을 디코딩합니다."""
    if not header_value:
        return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            charset = charset or 'utf-8'
            try:
                result.append(part.decode(charset, errors='replace'))
            except (LookupError, UnicodeDecodeError):
                result.append(part.decode('utf-8', errors='replace'))
        else:
            result.append(part)
    return ''.join(result)


def check_body_urls(msg):
    """
    이메일 본문(text/plain, text/html)에서 하이퍼링크(URL)가 존재하는지 검사합니다.
    Returns: 발견된 URL 목록 (없으면 빈 리스트)
    """
    urls = []
    url_pattern = re.compile(r'https?://[^\s<>"\')]+', re.IGNORECASE)
    
    for part in msg.walk():
        content_type = part.get_content_type()
        if content_type in ('text/plain', 'text/html'):
            try:
                body = part.get_content()
                if isinstance(body, str):
                    # HTML 엔티티 (예: &amp;, &#1092; 등) 디코딩
                    body = html.unescape(body)
                    found = url_pattern.findall(body)
                    urls.extend(found)
            except Exception:
                pass
    return urls


def load_safe_domains():
    """safe_domains.txt에서 안전 도메인 목록을 로드합니다."""
    safe_domains_file = os.path.join(SCRIPT_DIR, "safe_domains.txt")
    if os.path.exists(safe_domains_file):
        with open(safe_domains_file, "r", encoding="utf-8") as f:
            return {line.strip().lower() for line in f
                    if line.strip() and not line.startswith("#")}
    return set()


def extract_domain_from_url(url):
    """URL에서 도메인을 추출합니다. (서브도메인 포함)"""
    try:
        parsed = urlparse(url)
        domain = parsed.hostname or ''
        return domain.lower()
    except Exception:
        return ''


def is_safe_domain(url, safe_domains):
    """
    URL의 도메인이 safe_domains에 포함되는지 확인합니다.
    서브도메인도 매칭합니다. (예: mail.google.com → google.com에 매칭)
    """
    domain = extract_domain_from_url(url)
    if not domain:
        return False
    for safe in safe_domains:
        if domain == safe or domain.endswith('.' + safe):
            return True
    return False


def is_image_url(url):
    """URL 경로가 이미지 확장자로 끝나는지 확인합니다."""
    try:
        parsed = urlparse(url)
        path = parsed.path.lower()
        image_extensions = ('.png', '.gif', '.jpg', '.jpeg', '.svg', '.bmp', '.webp', '.ico')
        return path.endswith(image_extensions)
    except Exception:
        return False


def extract_attachments(eml_path, output_base_dir):
    """
    단일 EML 파일에서 첨부파일을 추출합니다.
    EML 파일명으로 폴더를 생성하고 그 안에 첨부파일을 저장합니다.
    
    Returns: (추출된 첨부파일 수, 생성된 폴더 경로, 파싱된 msg 객체, 첨부파일 존재 여부)
    """
    eml_filename = os.path.basename(eml_path)
    eml_name_no_ext = os.path.splitext(eml_filename)[0]
    folder_name = sanitize_filename(eml_name_no_ext)

    # EML 파일 파싱
    try:
        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=email.policy.default)
    except Exception as e:
        print(f"  [!] Failed to parse: {e}")
        return 0, None, None, False, False

    # 훈련 메일 여부 판단 (본문 어디든 dtsfm.shinhan.com 문자열이 포함되면 훈련 메일로 간주)
    is_training = False
    try:
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type in ('text/plain', 'text/html'):
                body_content = part.get_content()
                if isinstance(body_content, str) and "dtsfm.shinhan.com" in body_content.lower():
                    is_training = True
                    break
    except Exception:
        pass

    if is_training:
        folder_name = f"[훈련메일]{folder_name}"

    # 첨부파일 추출
    attachment_count = 0
    output_dir = None
    
    # 훈련 메일인 경우 파일 저장 여부와 상관없이 폴더 선제 생성
    if is_training:
        output_dir = os.path.join(output_base_dir, folder_name)
        os.makedirs(output_dir, exist_ok=True)

    has_any_attachment = False  # 지원/비지원 관계없이 첨부파일이 존재하는지 여부

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        content_disposition = str(part.get("Content-Disposition", ""))
        filename = part.get_filename()
        
        # 파일명이 없으면 Content-Type의 name 파라미터에서 시도
        if not filename:
            filename = part.get_param("name")

        if filename:
            filename = decode_mime_header(filename)
            filename = sanitize_filename(filename)

        # 첨부파일 또는 인라인 파일 판별 (파일명이 있으면 우선 대상으로 간주)
        if "attachment" in content_disposition or "inline" in content_disposition or filename:
            has_any_attachment = True
            
            if not filename:
                # 파일명이 끝내 없는 경우 Content-Type에서 확장자라도 추출 시도
                ext = part.get_content_type().split('/')[-1]
                filename = f"attachment_{attachment_count + 1}.{ext}"

            # 지원하는 확장자인지 검사 (file_analysis.py 768~780라인 기준)
            file_ext = os.path.splitext(filename)[1].lower().strip('.')
            if file_ext not in SUPPORTED_EXTENSIONS:
                # 메일 제목 추출 및 디코딩 추가
                mail_subject = msg.get("Subject", "제목없음")
                mail_subject = decode_mime_header(mail_subject)
                
                print(f"  [-] Skipped (Unsupported extension '.{file_ext}'): {filename} (Mail: {mail_subject})")
                
                # 로그 파일에 메일 제목과 함께 기록하기
                log_file = os.path.join(output_base_dir, "skipped_mails.txt")
                os.makedirs(output_base_dir, exist_ok=True)
                with open(log_file, "a", encoding="utf-8") as lf:
                    lf.write(f"[{mail_subject}] 스킵된 첨부파일 (지원하지 않는 확장자 '.{file_ext}'): {filename}\n")
                
                continue

            # 출력 폴더 생성 (첫 번째 지원하는 첨부파일 발견 시)
            if output_dir is None:
                output_dir = os.path.join(output_base_dir, folder_name)
                os.makedirs(output_dir, exist_ok=True)

            # 파일 저장
            filepath = os.path.join(output_dir, filename)
            
            # 동일 파일 존재 여부 확인 (중복 방지: 이미 있으면 새로 추출하지 않고 건너뜀)
            if os.path.exists(filepath):
                print(f"  [-] Skipped (Already exists): {filename}")
                continue

            try:
                payload = part.get_payload(decode=True)
                if payload:
                    # 첨부파일 내용(해시) 중복 체크
                    payload_hash = hashlib.sha256(payload).hexdigest()
                    hash_history_file = os.path.join(SCRIPT_DIR, "attachfiles", "extracted_hash_history.txt")
                    is_duplicate = False
                    if os.path.exists(hash_history_file):
                        with open(hash_history_file, 'r', encoding='utf-8') as hf:
                            if payload_hash in hf.read():
                                is_duplicate = True
                                
                    if is_duplicate:
                        print(f"  [-] Skipped (Duplicate Payload Hash): {filename}")
                        continue
                    else:
                        with open(hash_history_file, 'a', encoding='utf-8') as hf:
                            hf.write(f"{payload_hash}\n")

                    with open(filepath, 'wb') as f:
                        f.write(payload)
                    attachment_count += 1
                    print(f"  [+] Saved: {filename} ({len(payload):,} bytes)")
            except Exception as e:
                print(f"  [!] Failed to save {filename}: {e}")

    # 첨부파일 3개 이상 시 폴더명 앞에 [첨부파일 3개 이상] 접두사 붙이기
    LARGE_ATTACHMENT_THRESHOLD = 3
    if attachment_count >= LARGE_ATTACHMENT_THRESHOLD and output_dir and os.path.isdir(output_dir):
        parent = os.path.dirname(output_dir)
        old_name = os.path.basename(output_dir)
        if not old_name.startswith("[첨부파일 3개 이상]"):
            new_name = f"[첨부파일 3개 이상]{old_name}"
            new_dir = os.path.join(parent, new_name)
            try:
                if not os.path.exists(new_dir):
                    os.rename(output_dir, new_dir)
                    output_dir = new_dir
                    print(f"  [*] 폴더명 변경: {old_name} → {new_name}")
            except Exception as e:
                print(f"  [!] 폴더 이름 변경 실패: {e}")

    return attachment_count, output_dir, msg, has_any_attachment, is_training


def main():
    parser = argparse.ArgumentParser(description="EML Attachment Extractor")
    parser.add_argument("-dir", dest="eml_dir", default=DEFAULT_EML_DIR,
                        help=f"Directory containing .eml files (default: ./eml/)")
    parser.add_argument("-out", dest="output_dir", default=DEFAULT_OUTPUT_DIR,
                        help=f"Output directory for attachments (default: ./attachfiles/)")
    
    args = parser.parse_args()

    eml_dir = os.path.abspath(args.eml_dir)
    output_dir = os.path.abspath(args.output_dir)

    if not os.path.isdir(eml_dir):
        print(f"[!] EML directory not found: {eml_dir}")
        sys.exit(1)

    # EML 파일 스캔
    eml_files = [f for f in os.listdir(eml_dir) if f.lower().endswith('.eml')]
    
    if not eml_files:
        print(f"[*] No .eml files found in: {eml_dir}")
        return

    print(f"[*] Found {len(eml_files)} EML file(s) in: {eml_dir}")
    print(f"[*] Output directory: {output_dir}")
    print(f"{'='*60}\n")

    total_attachments = 0
    files_with_attachments = 0
    total_urls_saved = 0

    for eml_file in sorted(eml_files):
        eml_path = os.path.join(eml_dir, eml_file)
        print(f"[*] Processing: {eml_file}")
        
        count, folder, msg, has_any_attachment, is_training = extract_attachments(eml_path, output_dir)
        
        if count > 0:
            print(f"  [*] {count} attachment(s) extracted → {folder}")
            total_attachments += count
            files_with_attachments += 1
        else:
            print(f"  [-] No attachments found.")
        
        # 본문 URL 추출 (첨부파일 유무와 무관하게 모든 EML에서 수행)
        if msg is not None:
            if is_training:
                print("  [-] Skipped URL analysis (Training Mail)")
                print()
                continue
                
            body_urls = check_body_urls(msg)
            
            # 안전 도메인 필터링 + 이미지 URL 제외 + 중복 제거 (Base URL 기준 정규화)
            safe_domains = load_safe_domains()
            
            def smart_normalize(url):
                """추적용 인코딩 주소(Stibee, HubSpot 등)를 감지하여 정규화합니다."""
                import base64
                import binascii
                import re
                
                # 기본 정규화 (파라미터 제거)
                base = url.split('?')[0].split('#')[0].rstrip('/')
                
                # 1. 스티비(Stibee) 추적 링크 처리
                if "/v2/click/" in base:
                    try:
                        prefix = "/v2/click/"
                        idx = base.find(prefix)
                        if idx != -1:
                            after_prefix = base[idx + len(prefix):]
                            cleaned = re.sub(r'[^A-Za-z0-9+/=]', '', after_prefix)
                            start_match = re.search(r'aHR0[A-Za-z0-9+/=]+', cleaned)
                            if start_match:
                                token = start_match.group(0)
                                pad = len(token) % 4
                                if pad: token += '=' * (4 - pad)
                                decoded = base64.b64decode(token).decode('utf-8', errors='ignore')
                                if decoded.startswith('http'):
                                    return decoded.split('?')[0].split('#')[0].rstrip('/')
                    except Exception: pass

                # 2. 허브스팟(HubSpot) 추적 링크 처리
                # 허브스팟은 파라미터가 너무 길어 분석 오류를 일으키므로 파라미터 제거를 우선시함
                if "hubspotlinks.com" in base or "hubspotemail" in base or "hs-sites" in base:
                    return base

                return base

            unique_urls = set()
            domain_counts = {} # 도메인별 수집 개수 제한용
            
            for url in body_urls:
                if is_safe_domain(url, safe_domains) or is_image_url(url):
                    continue
                try:
                    # 스마트 정규화 적용
                    base_url = smart_normalize(url)
                    
                    # 정규화된 결과가 안전 도메인이거나 이미지면 제외 (스티비 내부 주소 검사)
                    if is_safe_domain(base_url, safe_domains) or is_image_url(base_url):
                        continue
                        
                    parsed = urlparse(base_url)
                    domain = (parsed.hostname or '').lower()
                    if not domain:
                        continue
                        
                    # 도메인별 샘플링 (최대 5개)
                    if domain not in domain_counts:
                        domain_counts[domain] = 0
                    
                    if domain_counts[domain] >= 5:
                        continue
                        
                    if base_url not in unique_urls:
                        unique_urls.add(base_url)
                        domain_counts[domain] += 1
                        
                except Exception:
                    unique_urls.add(url)
                    
            filtered_urls = sorted(list(unique_urls))
            
            if filtered_urls:
                # 메일 제목으로 폴더 생성
                eml_name_no_ext = os.path.splitext(eml_file)[0]
                url_folder_name = sanitize_filename(eml_name_no_ext)
                url_folder = os.path.join(output_dir, url_folder_name)
                os.makedirs(url_folder, exist_ok=True)
                
                urls_file = os.path.join(url_folder, "urls.txt")
                with open(urls_file, "w", encoding="utf-8") as uf:
                    for url in sorted(filtered_urls):
                        uf.write(url + "\n")
                print(f"  [URL] {len(filtered_urls)} URL(s) saved to urls.txt (filtered from {len(body_urls)} total)")
                total_urls_saved += len(filtered_urls)
            elif not has_any_attachment and not filtered_urls:
                # 첨부파일도 없고 분석할 유효 URL도 전혀 없는 경우
                subject = decode_mime_header(msg.get('Subject', '')) or '(제목 없음)'
                log_msg = f"[{subject}] 스킵된 깡통 메일 (분석 가능한 첨부파일 및 유효 URL 없음)"
                
                log_file = os.path.join(output_dir, "skipped_mails.txt")
                os.makedirs(output_dir, exist_ok=True)
                with open(log_file, "a", encoding="utf-8") as lf:
                    lf.write(log_msg + "\n")
        print()

    # 요약
    print(f"{'='*60}")
    print(f"[*] Summary:")
    print(f"    Total EML files processed: {len(eml_files)}")
    print(f"    Files with attachments: {files_with_attachments}")
    print(f"    Total attachments extracted: {total_attachments}")
    print(f"    Total URLs saved for analysis: {total_urls_saved}")
    print(f"    Output location: {output_dir}")


if __name__ == "__main__":
    main()
