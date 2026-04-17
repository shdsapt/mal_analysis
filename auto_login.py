#!/usr/bin/env python3
"""
mail.shinhan.com 자동 로그인 스크립트
- Selenium: 브라우저 자동화 (로그인 + 2차 인증 입력)
- IMAP: Gmail에서 2차 인증코드 자동 읽기
"""

# Windows cp949 인코딩 문제 방지
import os
os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"
import re
import sys
import time
import email
import hashlib
from urllib.parse import urlparse

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
import imaplib
import configparser
from datetime import datetime, timedelta

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        TimeoutException, NoSuchElementException, WebDriverException
    )
except ImportError:
    print("[!] selenium이 설치되어 있지 않습니다.")
    print("    실행: pip install selenium")
    sys.exit(1)

try:
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager
    HAS_WEBDRIVER_MANAGER = True
except ImportError:
    from selenium.webdriver.chrome.service import Service as ChromeService
    HAS_WEBDRIVER_MANAGER = False


try:
    if getattr(sys, 'frozen', False):
        # PyInstaller: EXE 파일 위치 최우선
        frozen_base = os.path.dirname(sys.executable)
        sys.path.insert(0, frozen_base)
except Exception:
    pass

# ─────────────────────────────────────────────
# URL 필터링 헬퍼 함수
# ─────────────────────────────────────────────
def _load_safe_domains():
    """safe_domains.txt에서 안전 도메인 목록을 로드합니다."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    safe_domains_file = os.path.join(script_dir, "safe_domains.txt")
    if os.path.exists(safe_domains_file):
        with open(safe_domains_file, "r", encoding="utf-8") as f:
            return {line.strip().lower() for line in f
                    if line.strip() and not line.startswith("#")}
    return set()


def _is_safe_domain(url, safe_domains):
    """URL의 도메인이 안전 도메인 목록에 있는지 확인합니다."""
    try:
        domain = (urlparse(url).hostname or '').lower()
        for safe in safe_domains:
            if domain == safe or domain.endswith('.' + safe):
                return True
    except Exception:
        pass
    return False


def _is_image_url(url):
    """URL 경로가 이미지 확장자로 끝나는지 확인합니다."""
    try:
        path = urlparse(url).path.lower()
        return path.endswith(('.png', '.gif', '.jpg', '.jpeg', '.svg',
                              '.bmp', '.webp', '.ico'))
    except Exception:
        return False


# ─────────────────────────────────────────────
# 설정 로드
# ─────────────────────────────────────────────
def load_config(config_path=None):
    """config.ini 파일에서 설정을 로드합니다."""
    if config_path is None:
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_dir, "config.ini")

    if not os.path.exists(config_path):
        print(f"[!] 설정 파일을 찾을 수 없습니다: {config_path}")
        print("    config.ini 파일을 생성하고 계정 정보를 입력해주세요.")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")
    
    # 런타임 속도 단축을 위해 polling 기본값 조정 (5초 -> 2초)
    if not config.has_section("gmail_imap"):
        config.add_section("gmail_imap")
    if not config.has_option("gmail_imap", "poll_interval_seconds"):
        config.set("gmail_imap", "poll_interval_seconds", "2")
        
    return config



# ─────────────────────────────────────────────
# Gmail IMAP: 2차 인증코드 읽기
# ─────────────────────────────────────────────
def get_verification_code(config):
    """
    Gmail IMAP에 접속하여 최신 인증코드 메일에서 코드를 추출합니다.
    최대 max_wait_seconds 동안 poll_interval_seconds 간격으로 폴링합니다.
    """
    gmail_email = config.get("gmail_imap", "email")
    app_password = config.get("gmail_imap", "app_password")
    imap_server = config.get("gmail_imap", "imap_server", fallback="imap.gmail.com")
    sender_filter = config.get("gmail_imap", "sender_filter", fallback="shinhan")
    code_pattern = config.get("gmail_imap", "code_pattern", fallback=r"\d{6}")
    max_wait = config.getint("gmail_imap", "max_wait_seconds", fallback=60)
    # 강제로 폴링 주기를 1초로 세팅하여 초고속 수신
    poll_interval = 1
    
    # 로그인 시도 시각 기록 (이전 메일 무시용)
    search_after = datetime.now() - timedelta(minutes=2)
    
    print(f"[*] Gmail IMAP 접속 중... ({gmail_email})")
    
    elapsed = 0
    while elapsed < max_wait:
        try:
            # IMAP 서버 접속
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(gmail_email, app_password)
            mail.select("INBOX")
            
            # 최근 메일 검색 (날짜 기준)
            date_str = search_after.strftime("%d-%b-%Y")
            search_criteria = f'(SINCE "{date_str}" FROM "{sender_filter}")'
            
            status, message_ids = mail.search(None, search_criteria)
            
            if status == "OK" and message_ids[0]:
                ids = message_ids[0].split()
                # 가장 최근 메일부터 확인 (역순)
                for msg_id in reversed(ids):
                    status, msg_data = mail.fetch(msg_id, "(RFC822)")
                    if status != "OK":
                        continue
                    
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # 메일 수신 시간 확인
                    msg_date = email.utils.parsedate_to_datetime(msg["Date"])
                    if msg_date.replace(tzinfo=None) < search_after:
                        continue
                    
                    # 메일 본문에서 인증코드 추출
                    body = _get_email_body(msg)
                    if body:
                        match = re.search(code_pattern, body)
                        if match:
                            code = match.group(0)
                            print(f"[+] 인증코드 발견: {code}")
                            mail.logout()
                            return code
            
            mail.logout()
            
        except imaplib.IMAP4.error as e:
            print(f"[!] IMAP 오류: {e}")
            print("    Gmail 앱 비밀번호를 확인해주세요.")
            return None
        except Exception as e:
            print(f"[!] 메일 확인 중 오류: {e}")
        
        elapsed += poll_interval
        if elapsed < max_wait:
            remaining = max_wait - elapsed
            print(f"[*] 인증코드 메일 대기 중... ({elapsed}s / {max_wait}s, 남은 시간: {remaining}s)")
            time.sleep(poll_interval)
    
    print(f"[!] {max_wait}초 동안 인증코드 메일을 수신하지 못했습니다.")
    return None


def _get_email_body(msg):
    """이메일 메시지에서 텍스트 본문을 추출합니다."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                    break
                except Exception:
                    continue
            elif content_type == "text/html":
                try:
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                except Exception:
                    continue
    else:
        try:
            body = msg.get_payload(decode=True).decode("utf-8", errors="replace")
        except Exception:
            pass
    return body


# ─────────────────────────────────────────────
# Selenium: 브라우저 자동화
# ─────────────────────────────────────────────
def create_driver(config):
    """Selenium WebDriver를 생성합니다."""
    browser_type = config.get("browser", "browser_type", fallback="chrome")
    headless = config.getboolean("browser", "headless", fallback=False)
    timeout = config.getint("browser", "page_load_timeout", fallback=30)
    
    if browser_type.lower() == "chrome":
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1280,900")
        # SSL 인증서 오류 무시 (사내망 등)
        options.add_argument("--ignore-certificate-errors")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        # 다운로드 디렉토리 설정: EXE 위치 기준
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        download_dir = os.path.join(base_dir, "eml")
        os.makedirs(download_dir, exist_ok=True)
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)
        
        # 드라이버 경로 탐색: 1순위 동일 폴더 chromedriver.exe
        if getattr(sys, 'frozen', False):
            cwd_driver = os.path.join(os.path.dirname(sys.executable), "chromedriver.exe")
        else:
            cwd_driver = os.path.join(os.getcwd(), "chromedriver.exe")
            
        try:
            if os.path.exists(cwd_driver):
                print(f"[*] 로컬 ChromeDriver 발견: {cwd_driver}")
                service = ChromeService(executable_path=cwd_driver)
                driver = webdriver.Chrome(service=service, options=options)
            elif HAS_WEBDRIVER_MANAGER:
                print("[*] WebDriverManager를 사용하여 드라이버 설치 시도...")
                driver_path = ChromeDriverManager().install()
                # ... (이후 기존 로직 동일하게 처리 가능하지만 간단하게 처리)
                service = ChromeService(executable_path=driver_path)
                driver = webdriver.Chrome(service=service, options=options)
            else:
                driver = webdriver.Chrome(options=options)
        except Exception as e:
            print(f"[!] ChromeDriver 실행 실패: {e}")
            print("[*] 시스템 기본 설치된 브라우저 제어 시도...")
            driver = webdriver.Chrome(options=options)
    
    elif browser_type.lower() == "edge":
        from selenium.webdriver.edge.service import Service as EdgeService
        options = webdriver.EdgeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--ignore-certificate-errors")
        driver = webdriver.Edge(options=options)
    
    else:
        print(f"[!] 지원하지 않는 브라우저: {browser_type}")
        sys.exit(1)
    
    driver.set_page_load_timeout(timeout)
    driver.implicitly_wait(10)
    return driver


def login_shinhan_mail(config):
    """
    mail.shinhan.com에 자동 로그인합니다.
    1단계: ID/PW 입력 → 로그인 클릭
    2단계: Gmail에서 2차 인증코드 읽기
    3단계: 인증코드 입력 → 최종 로그인
    """
    url = config.get("shinhan_mail", "url")
    username = config.get("shinhan_mail", "username")
    password = config.get("shinhan_mail", "password")
    
    id_selector = config.get("shinhan_mail", "id_selector", fallback="#userId")
    pw_selector = config.get("shinhan_mail", "pw_selector", fallback="#userPw")
    login_btn_selector = config.get("shinhan_mail", "login_btn_selector", fallback="#loginBtn")
    otp_input_selector = config.get("shinhan_mail", "otp_input_selector", fallback="#otpNo")
    otp_submit_selector = config.get("shinhan_mail", "otp_submit_selector", fallback="#otpBtn")
    login_success_selector = config.get("shinhan_mail", "login_success_selector", fallback=".mail-list")
    
    # 계정 정보 확인
    if username == "your_id" or password == "your_password":
        print("[!] config.ini에 실제 계정 정보를 입력해주세요.")
        sys.exit(1)
    
    driver = None
    try:
        # ─── Step 1: 로그인 페이지 접속 ───
        print(f"[*] 브라우저 시작...")
        driver = create_driver(config)
        
        print(f"[*] {url} 접속 중...")
        driver.get(url)
        
        # 페이지 로딩을 title 존재 여부로 동적 대기
        try:
            WebDriverWait(driver, 10).until(lambda d: d.title)
        except Exception:
            pass
            
        print(f"[*] 현재 URL: {driver.current_url}")
        print(f"[*] 페이지 제목: {driver.title}")
        
        # ─── Step 2: ID/PW 입력 ───
        print(f"[*] ID 입력 중... (셀렉터: {id_selector})")
        try:
            id_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, id_selector))
            )
            id_field.clear()
            id_field.send_keys(username)
        except TimeoutException:
            print(f"[!] ID 입력 필드를 찾을 수 없습니다. 셀렉터를 확인해주세요: {id_selector}")
            _print_page_debug(driver)
            return False
        
        print(f"[*] Password 입력 중... (셀렉터: {pw_selector})")
        try:
            pw_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, pw_selector))
            )
            pw_field.clear()
            pw_field.send_keys(password)
        except TimeoutException:
            print(f"[!] PW 입력 필드를 찾을 수 없습니다. 셀렉터를 확인해주세요: {pw_selector}")
            _print_page_debug(driver)
            return False
        
        # ─── Step 3: 로그인 실행 ───
        # 방법 1: PW 필드에서 Enter 키 전송 (가장 자연스러운 방식)
        print("[*] 로그인 실행 중... (Enter 키 전송)")
        pw_field.send_keys(Keys.RETURN)
        
        print("[*] 로그인 요청 전송. 페이지 전환 대기 중...")
        # 기존 5초 고정 대기를 로그인 후의 URL 변동이나 알럿 발생 여부로 동적으로 기다림
        try:
            WebDriverWait(driver, 10).until(
                lambda d: d.current_url != url and "login" not in d.current_url.lower() or EC.alert_is_present()(d)
            )
        except TimeoutException:
            pass
        
        # 로그인 후 에러 메시지 확인
        print(f"[*] 로그인 후 URL: {driver.current_url}")
        
        # Alert 팝업 확인
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            print(f"[!] Alert 메시지: {alert_text}")
            alert.accept()
        except Exception:
            pass
        
        # 에러 메시지 요소 확인 (일반적인 로그인 에러 패턴)
        error_selectors = [".error", ".alert", ".err-msg", ".login-error", 
                          "[class*='error']", "[class*='alert']", "[class*='warn']"]
        for es in error_selectors:
            try:
                err_el = driver.find_element(By.CSS_SELECTOR, es)
                if err_el.text.strip():
                    print(f"[!] 에러 메시지 감지: {err_el.text.strip()}")
            except Exception:
                continue
        
        # ─── Step 4: 2차 인증코드 입력 필드 확인 ───
        # OTP 입력 필드가 나타나는지 확인 (15초 대기)
        try:
            otp_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, otp_input_selector))
            )
            print("[*] 2차 인증 입력 화면 감지!")
        except TimeoutException:
            print("[*] 기본 OTP 셀렉터로 감지 실패. 페이지 구조를 분석합니다...")
            print(f"[*] 현재 URL: {driver.current_url}")
            
            # /twoFactorAuth 페이지인 경우 → OTP 필드를 다양한 방법으로 찾기
            if "twoFactor" in driver.current_url or "otp" in driver.current_url.lower():
                print("[*] 2차 인증 페이지 감지! OTP 입력 필드를 탐색합니다...")
                _print_page_debug(driver)
                
                # 다양한 셀렉터로 OTP 입력 필드 탐색
                otp_candidates = [
                    otp_input_selector,
                    "input[type='text']",
                    "input[type='number']",
                    "input[type='tel']",
                    "input[name*='otp']",
                    "input[name*='code']",
                    "input[name*='auth']",
                    "input[id*='otp']",
                    "input[id*='code']",
                    "input[id*='auth']",
                    "input[placeholder*='인증']",
                    "input[placeholder*='코드']",
                ]
                otp_field = None
                for candidate in otp_candidates:
                    try:
                        found = driver.find_elements(By.CSS_SELECTOR, candidate)
                        if found:
                            otp_field = found[0]
                            print(f"[+] OTP 입력 필드 발견! (셀렉터: {candidate})")
                            break
                    except Exception:
                        continue
                
                if otp_field is None:
                    print("[!] OTP 입력 필드를 찾을 수 없습니다.")
                    print("[*] 위 디버그 정보에서 input 요소를 확인하고 config.ini의 otp_input_selector를 수정해주세요.")
                    return False
            else:
                # URL이 변경되었는지 확인 (2FA가 아닌 다른 페이지)
                if driver.current_url != url and driver.current_url != url + "/login":
                    print("[+] URL이 변경되었습니다. 로그인 성공으로 추정됩니다.")
                    return True
                
                # 로그인 성공 여부 확인
                try:
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, login_success_selector))
                    )
                    print("[+] 로그인 성공! (2차 인증 불필요)")
                    return True
                except TimeoutException:
                    print("[!] 2차 인증도 없고, 로그인도 안 됨. 페이지를 확인해주세요.")
                    _print_page_debug(driver)
                    return False
        
        # ─── Step 5: 메일 인증 모드 선택 + 인증코드 발송 ───
        print("[*] 2차 인증 외부메일 선택 및 연계 처리 가동...")
        send_clicked = False
        
        try:
            # 1단계: 메일 인증 라디오 버튼 즉시 클릭 (DOM JS 활용)
            driver.execute_script("""
                var radio = document.getElementById('authMode_mail');
                if(radio && !radio.checked) {
                    radio.click();
                } else {
                    var spans = document.querySelectorAll('span.option_wrap, label');
                    for(var i=0; i<spans.length; i++) {
                        if(spans[i].innerText.includes('외부 메일') || spans[i].innerText.includes('메일 인증')) {
                            spans[i].click();
                            break;
                        }
                    }
                }
            """)
            print("[*] 메일 인증 모드 선택 완료 (JS Fast Path)")
            # UI 전환을 위한 최소한의 프레임 대기
            time.sleep(0.2)
            
            # 2단계: 인증코드 발송 버튼 즉시 타격 (id='issueAuthCode' 우선순위)
            clicked = driver.execute_script("""
                var btn = document.getElementById('issueAuthCode');
                if(btn) {
                    btn.click();
                    return true;
                }
                var links = document.querySelectorAll('a, button, input');
                for(var i=0; i<links.length; i++) {
                    var t = links[i].innerText || links[i].value || '';
                    if(t.includes('요청') || t.includes('발송') || t.includes('전송') || t.includes('인증코드')) {
                        links[i].click();
                        return true;
                    }
                }
                return false;
            """)
            if clicked:
                print("[*] 인증코드 발송 버튼 클릭 (JS Fast Path)")
                send_clicked = True
        except Exception:
            pass
            
        # JS 패스 실패 시, 기존 브로드 탐색 수행
        if not send_clicked:
            print("[*] 기존 브로드캐스트 버튼 탐색 활성화...")
            send_btn_selectors = [
                "button", "input[type='submit']", "input[type='button']",
                "a[class*='btn']", "[class*='send']", "[class*='request']",
                "[onclick*='send']", "[onclick*='auth']",
            ]
            for sel in send_btn_selectors:
                try:
                    btns = driver.find_elements(By.CSS_SELECTOR, sel)
                    for btn in btns:
                        btn_text = btn.text.strip() if btn.text else ""
                        btn_value = btn.get_attribute("value") or ""
                        if any(kw in (btn_text + btn_value) for kw in ["발송", "전송", "요청", "send", "인증"]):
                            driver.execute_script("arguments[0].click();", btn)
                            print(f"[*] 인증코드 발송 버튼 클릭: '{btn_text or btn_value}'")
                            send_clicked = True
                            break
                except Exception:
                    continue
                if send_clicked:
                    break
        
        if not send_clicked:
            # 모든 버튼/링크를 출력하여 디버깅
            print("[!] 인증코드 발송 버튼을 찾을 수 없습니다.")
            # (디버그 로그 출력 생략)
            pass
        
        # ─── Step 6: Gmail에서 인증코드 읽기 ───
        print("[*] Gmail에서 인증코드 읽기 시작...")
        verification_code = get_verification_code(config)
        
        if not verification_code:
            print("[!] 인증코드를 가져오지 못했습니다.")
            print("[*] 수동으로 인증코드를 입력해주세요:")
            verification_code = input(">>> 인증코드: ").strip()
            if not verification_code:
                print("[!] 인증코드가 입력되지 않았습니다. 종료합니다.")
                return False
        
        # ─── Step 6: 인증코드 입력 + 제출 ───
        print(f"[*] 인증코드 입력 중: {verification_code}")
        otp_field.clear()
        otp_field.send_keys(verification_code)
        
        # 버튼 탐색을 제한적으로 빠르게 수행 (요청/발송 제외)
        print("[*] 인증 제출 버튼 즉시 탐색...")
        submit_clicked = False
        try:
            # 보통 OTP 입력창 근처나 특정 영역에 "확인" 버튼이 있음
            btns = driver.find_elements(By.CSS_SELECTOR, "button, input[type='submit'], a, span")
            for btn in btns:
                t = btn.text.strip()
                if "확인" in t and "요청" not in t and "발송" not in t and btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    print(f"[*] OTP 제출 버튼 클릭: '{t}'")
                    submit_clicked = True
                    break
        except Exception:
            pass
            
        if not submit_clicked:
            otp_field.send_keys(Keys.RETURN)
            print("[*] 폼 제출 직행... (Enter)")
            
        # 서버에서 로그인 검증 처리할 물리적 시간 확보 (필수)
        time.sleep(1)
        
        # ─── Step 7: 인증 완료 팝업 처리 ───
        # 1차: JavaScript Alert 확인
        popup_handled = False
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            print(f"[+] Alert 팝업 감지: '{alert.text}'")
            alert.accept()
            popup_handled = True
            time.sleep(1)
        except (TimeoutException, Exception):
            pass
        
        # 2차: 커스텀 모달 팝업 초고속 네이티브 이벤트 처리
        if not popup_handled:
            print("[*] 커스텀 팝업 확인창 탐색 시도...")
            
            try:
                # 팝업 컨테이너 자체의 렌더링 완료 대기 (최대 5초)
                # 이 요소 안에 있는 진짜 팝업용 버튼만 취급
                popup_container = WebDriverWait(driver, 5).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, ".btn_layer_wrap, [class*='popup'], .modal[style*='block'], .swal2-popup")
                )
                time.sleep(0.5) # 애니메이션 및 브라우저 이벤트 큐 안정화
                
                # '화면 전체'가 아니라 오직 '팝업 컨테이너 내부'의 확인 버튼만 검색
                btns = popup_container.find_elements(By.CSS_SELECTOR, ".btn_layer_wrap, button, a")
                for btn in btns:
                    if btn.is_displayed() and "확인" in btn.text.strip():
                        btn.click()
                        print(f"[+] 팝업 내 확인 버튼 타격 완료: '{btn.text.strip()}'")
                        popup_handled = True
                        break
                        
                # 팝업 내부 검색 실패 시 최후 보루: btn_layer_wrap 클래스를 가진 녀석을 직접 클릭
                if not popup_handled:
                    fallback_btn = driver.find_element(By.CSS_SELECTOR, ".btn_layer_wrap")
                    fallback_btn.click()
                    print("[+] btn_layer_wrap 클래스 직접 클릭 완료")
                    popup_handled = True
                    
            except Exception:
                pass
            
            # 여기서도 못 찾으면 일반 버튼 탐색 로직 수행 but 'submit'이나 'OTP버튼' id 솎아내기
            if not popup_handled:
                print("[*] 전체 페이지에서 보이는 '확인' 버튼 탐색...")
                all_visible = driver.find_elements(By.CSS_SELECTOR, 
                    "button, a, input[type='button'], input[type='submit'], [class*='btn'], span")
                visible_confirms = []
                for el in all_visible:
                    try:
                        if el.is_displayed():
                            t = el.text.strip()
                            el_id = el.get_attribute("id") or ""
                            el_class = el.get_attribute("class") or ""
                            el_tag = el.tag_name
                            if t:
                                visible_confirms.append({
                                    "element": el, "text": t, "id": el_id, 
                                    "class": el_class, "tag": el_tag
                                })
                    except Exception:
                        continue
                
                # 보이는 텍스트 요소 탐색 (로그 출력 생략)
                pass
                
                # 정확히 "확인"만 포함된 버튼 찾기 (OTP 확인과 겹치지 않게 필터링)
                for vc in visible_confirms:
                    if vc["text"] == "확인":
                        # 이전 단계에서 클릭한 id='submit' 등은 제외
                        if vc["id"] == "submit" or "btn_major" in vc["class"]:
                            continue
                            
                        driver.execute_script("arguments[0].click();", vc["element"])
                        print(f"[+] 최종 확인 버튼 클릭: <{vc['tag']}> id={vc['id']}, class={vc['class'][:30]}")
                        popup_handled = True
                        break
        
        if popup_handled:
            # URL 변경 대기 (최대 10초)
            print("[*] 페이지 전환 대기 중...")
            for _ in range(20):
                time.sleep(0.5)
                if "twoFactorAuth" not in driver.current_url:
                    print(f"[+] URL 변경 감지: {driver.current_url}")
                    break
        else:
            print("[*] 팝업이 감지되지 않았습니다.")
        
        # ─── Step 7: 로그인 완료 확인 ───
        try:
            WebDriverWait(driver, 15).until(
                lambda d: d.find_elements(By.CSS_SELECTOR, login_success_selector) or "mailCommon.do" in d.current_url
            )
            print("[+] ✅ 로그인 성공!")
            print(f"[+] 최종 URL: {driver.current_url}")
            return driver
        except TimeoutException:
            print("[*] 로그인 성공 요소를 확인할 수 없습니다.")
            print(f"[*] 현재 URL: {driver.current_url}")
            print(f"[*] 페이지 제목: {driver.title}")
            # URL 변화로 성공 여부 추정
            if "mailCommon.do" in driver.current_url:
                print("[+] 메일함 URL 감지. 로그인 성공!")
                return driver
            if "twoFactorAuth" not in driver.current_url and "login" not in driver.current_url:
                print("[+] 인증 페이지를 벗어남. 로그인 성공으로 추정.")
                return driver
            return None
    
    except WebDriverException as e:
        print(f"[!] 브라우저 오류: {e}")
        return None
    except Exception as e:
        print(f"[!] 예상치 못한 오류: {e}")
        return None


def _print_page_debug(driver):
    """디버깅용: 현재 페이지의 주요 요소를 출력합니다."""
    print("\n[DEBUG] === 페이지 디버그 정보 ===")
    print(f"  URL: {driver.current_url}")
    print(f"  Title: {driver.title}")
    
    # input 요소 목록 출력
    try:
        inputs = driver.find_elements(By.TAG_NAME, "input")
        print(f"  발견된 input 요소: {len(inputs)}개")
        for idx, inp in enumerate(inputs[:10]):
            inp_id = inp.get_attribute("id") or "(없음)"
            inp_name = inp.get_attribute("name") or "(없음)"
            inp_type = inp.get_attribute("type") or "(없음)"
            inp_placeholder = inp.get_attribute("placeholder") or "(없음)"
            print(f"    [{idx}] id={inp_id}, name={inp_name}, type={inp_type}, placeholder={inp_placeholder}")
    except Exception:
        pass
    
    # button 요소 목록 출력
    try:
        buttons = driver.find_elements(By.TAG_NAME, "button")
        print(f"  발견된 button 요소: {len(buttons)}개")
        for idx, btn in enumerate(buttons[:5]):
            btn_id = btn.get_attribute("id") or "(없음)"
            btn_text = btn.text.strip() or "(빈 텍스트)"
            print(f"    [{idx}] id={btn_id}, text={btn_text}")
    except Exception:
        pass
    
    # iframe 목록 출력
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print(f"  ⚠️ iframe 발견: {len(iframes)}개 (iframe 내부 요소 접근 시 switch_to.frame 필요)")
            for idx, iframe in enumerate(iframes):
                iframe_id = iframe.get_attribute("id") or "(없음)"
                iframe_src = iframe.get_attribute("src") or "(없음)"
                print(f"    [{idx}] id={iframe_id}, src={iframe_src[:80]}")
    except Exception:
        pass
    
    print("[DEBUG] ==============================\n")


# ─────────────────────────────────────────────
# EML 첨부파일 다운로드
# ─────────────────────────────────────────────
def download_eml_attachments(driver):
    """
    로그인 후 받은편지함에서 모든 메일을 확인하고,
    첨부파일 중 .eml 확장자를 가진 파일을 모두 다운로드합니다.
    
    신한 메일 구조:
    - 메일 목록: table.mail_list > tr (id=Inbox_XXXX)
    - 안 읽은 메일: tr.read_no
    - 첨부파일 목록: #attachListWrap > li
    - EML 아이콘: .ic_file.ic_eml
    - 다운로드 클릭: span[evt-rol="download-attach"]
    """
    print("\n" + "="*50)
    print("  [자동화] 받은편지함 EML 첨부파일 다운로드")
    print("="*50)
    
    download_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eml")
    os.makedirs(download_dir, exist_ok=True)
    
    # 파일명에서 이모지 및 Windows에서 사용 불가능한 특수문자를 제거하는 함수
    import re as _re
    def sanitize_filename(name):
        # cp949에서 표현 불가능한 모든 이모지/특수 유니코드 문자 제거
        # BMP 이모지: Dingbats, Misc Symbols, Emoticons, Transport, 등
        emoji_pattern = _re.compile(
            "["
            "\U00002600-\U000027BF"  # Misc Symbols, Dingbats (☀☁☂...✂✈✉❤...)
            "\U0000FE00-\U0000FE0F"  # Variation Selectors
            "\U0000FE20-\U0000FE2F"  # Combining Half Marks
            "\U00002702-\U000027B0"  # Dingbats
            "\U00002B50-\U00002B55"  # Stars
            "\U0000200D"             # Zero Width Joiner
            "\U000023E9-\U000023FA"  # Media symbols
            "\U00002934-\U00002935"  # Arrows
            "\U000025AA-\U000025FE"  # Geometric shapes
            "\U00002300-\U000023FF"  # Misc Technical
            "\U00002190-\U000021FF"  # Arrows
            "\U00010000-\U0010FFFF"  # 모든 Supplementary plane (🚀🎉💯 등)
            "]+", flags=_re.UNICODE
        )
        name = emoji_pattern.sub('', name)
        # Windows 파일명 금지 문자 제거
        name = _re.sub(r'[<>:"/\\|?*]', '_', name)
        # 앞뒤 공백/점 제거
        name = name.strip(' .')
        return name
    
    time.sleep(3)  # 페이지 완전 로딩 대기
    
    # ── 페이지당 표시 건수를 80으로 변경 ──
    try:
        from selenium.webdriver.support.ui import Select
        pagebase_select = driver.find_element(By.CSS_SELECTOR, "#toolbar_list_pagebase")
        select_obj = Select(pagebase_select)
        current_value = select_obj.first_selected_option.get_attribute("value")
        if current_value != "80":
            select_obj.select_by_value("80")
            print("[*] 페이지당 표시 건수를 80으로 변경했습니다.")
            time.sleep(5)  # 목록 새로고침 대기
        else:
            print("[*] 페이지당 표시 건수: 이미 80")
    except Exception as e:
        print(f"[!] 페이지당 건수 변경 실패 (기본값 유지): {e}")
    
    # ── 전체 메일 수 확인 ──
    try:
        page_navi = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap")
        data_total = page_navi.get_attribute("data-total")
        data_pagebase = page_navi.get_attribute("data-pagebase")
        print(f"[*] 전체 메일 수: {data_total}, 페이지당: {data_pagebase}")
    except Exception:
        pass
    
    # ── 모든 페이지에서 메일 ID 수집 ──
    all_mail_ids = []
    current_page = 1
    
    while True:
        mail_rows = driver.find_elements(By.CSS_SELECTOR, "tr[id^='Inbox_']")
        page_count = len(mail_rows)
        print(f"\n[*] 페이지 {current_page}: {page_count}개 메일 발견")
        
        if page_count == 0 and current_page == 1:
            print("[-] 처리할 메일이 없습니다.")
            return
        
        # 현재 페이지의 메일 ID 수집
        for row in mail_rows:
            mid = row.get_attribute("id")
            if mid and mid not in all_mail_ids:
                all_mail_ids.append(mid)
        
        # 다음 페이지 버튼 확인
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.next.paginate_button")
            if "paginate_button_disabled" in (next_btn.get_attribute("class") or ""):
                print(f"[*] 마지막 페이지입니다. (총 {current_page} 페이지)")
                break
            else:
                driver.execute_script("arguments[0].click();", next_btn)
                print(f"[*] 다음 페이지({current_page + 1})로 이동 중...")
                current_page += 1
                time.sleep(3)  # 페이지 로딩 대기
        except Exception:
            print(f"[*] 다음 페이지 버튼 없음. (총 {current_page} 페이지)")
            break
    
    print(f"\n[*] 전체 메일 ID 수집 완료: {len(all_mail_ids)}개")
    
    # ── 첫 페이지로 복귀 (메일 클릭 후 목록 복귀 시 항상 1페이지로 돌아감) ──
    if current_page > 1:
        try:
            first_btn = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.first.paginate_button")
            driver.execute_script("arguments[0].click();", first_btn)
            time.sleep(3)
        except Exception:
            driver.get("https://mail.shinhan.com/mail/mail/mailCommon.do?state=1")
            time.sleep(5)
    
    mail_ids = all_mail_ids
    downloaded_count = 0
    skipped_count = 0
    
    history_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachfiles", "download_history.txt")
    os.makedirs(os.path.dirname(history_file), exist_ok=True)
    processed_ids = set()
    if os.path.exists(history_file):
        try:
            with open(history_file, "r", encoding="utf-8") as f:
                processed_ids = set(line.strip() for line in f if line.strip())
        except Exception:
            pass
            
    print(f"[*] 기존 다운로드(처리) 완료 메일 이력: {len(processed_ids)}건")
    
    for idx, mail_id in enumerate(mail_ids):
        print(f"\n[*] --- 메일 {idx+1}/{len(mail_ids)} (ID: {mail_id}) ---")
        
        if mail_id in processed_ids:
            print("[*] 이미 다운로드/처리 이력이 있는 메일입니다. (중복 방지 스킵)")
            skipped_count += 1
            continue
            
        # 메일 행 재탐색 (DOM이 변경될 수 있으므로)
        try:
            row = driver.find_element(By.CSS_SELECTOR, f"#{mail_id}")
        except Exception:
            # 현재 페이지에 해당 메일이 없음 → 다음 페이지로 이동 시도
            print(f"[*] 현재 페이지에 {mail_id} 없음. 다음 페이지로 이동 시도...")
            try:
                # 페이지당 건수를 다시 80으로 설정 (목록 복귀 시 리셋될 수 있음)
                try:
                    from selenium.webdriver.support.ui import Select
                    pb_sel = driver.find_element(By.CSS_SELECTOR, "#toolbar_list_pagebase")
                    sel_obj = Select(pb_sel)
                    if sel_obj.first_selected_option.get_attribute("value") != "80":
                        sel_obj.select_by_value("80")
                        time.sleep(5)
                except Exception:
                    pass
                
                next_btn = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.next.paginate_button")
                if "paginate_button_disabled" not in (next_btn.get_attribute("class") or ""):
                    driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(3)
                    # 다시 찾기
                    try:
                        row = driver.find_element(By.CSS_SELECTOR, f"#{mail_id}")
                    except Exception:
                        print(f"[!] 다음 페이지에서도 메일 행을 찾을 수 없습니다: {mail_id}")
                        continue
                else:
                    print(f"[!] 마지막 페이지입니다. 메일을 찾을 수 없습니다: {mail_id}")
                    continue
            except Exception:
                print(f"[!] 메일 행을 찾을 수 없습니다: {mail_id}")
                continue
        
        # 메일 제목 추출
        try:
            link = row.find_element(By.TAG_NAME, "a")
            title = link.text.strip()[:60]
            print(f"[*] 제목: {title}")
        except Exception:
            title = "(제목 없음)"
            print(f"[*] 제목을 가져올 수 없습니다.")
        
        # 메일 클릭 (본문으로 이동)
        try:
            link = row.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", link)
            time.sleep(3)  # 본문 로딩 대기
        except Exception as e:
            print(f"[!] 메일 클릭 실패: {e}")
            continue
        
        # 첨부파일 목록 확인
        eml_found = False
        try:
            attach_wrap = driver.find_elements(By.CSS_SELECTOR, "#attachListWrap li")
            if not attach_wrap:
                print("[-] 첨부파일 없음. 본문 URL 스크래핑 및 [은행신고-첨부파일X] 폴더 생성 시도...")
                try:
                    attach_base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachfiles")
                    safe_title = re.sub(r'[<>:"/\\|?*]', '_', title).strip(' .')[:80].strip()
                    if not safe_title:
                        safe_title = "Unknown_Subject"
                    
                    no_attach_dir = os.path.join(attach_base_dir, f"[은행신고-첨부파일X]{safe_title}")
                    os.makedirs(no_attach_dir, exist_ok=True)
                    print(f"[+] 폴더 생성 완료: {no_attach_dir}")

                    # ── div#mainContentWrap 영역에서 본문 URL 직접 수집 ──
                    try:
                        safe_domains = _load_safe_domains()
                        raw_links = driver.execute_script("""
                            var wrap = document.querySelector('#mainContentWrap');
                            if (!wrap) return [];
                            return Array.from(wrap.querySelectorAll('a[href]'))
                                        .map(function(a) { return a.href; })
                                        .filter(function(h) { return h.startsWith('http'); });
                        """)
                        # 안전 도메인, 이미지 URL 필터링 + 스마트 정규화 + 도메인별 샘플링(최대 5개)
                        unique_set = set()
                        dom_counts = {}

                        def _smart_normalize(url):
                            import base64
                            import binascii
                            import re
                            base = url.split('?')[0].split('#')[0].rstrip('/')
                            prefix = "/v2/click/"
                            if prefix in base:
                                try:
                                    idx = base.find(prefix)
                                    # Base64 알맹이만 추출 (세척 과정 유무가 핵심)
                                    after = base[idx+len(prefix):]
                                    cleaned = re.sub(r'[^A-Za-z0-9+/=]', '', after)
                                    match = re.search(r'aHR0[A-Za-z0-9+/=]+', cleaned)
                                    if match:
                                        token = match.group(0)
                                        p = len(token) % 4
                                        if p: token += '=' * (4 - p)
                                        decoded = base64.b64decode(token).decode('utf-8', errors='ignore')
                                        if decoded.startswith('http'):
                                            return decoded.split('?')[0].split('#')[0].rstrip('/')
                                except Exception: pass
                            
                            # HubSpot 등 너무 긴 트래킹 링크는 파라미터를 제거한 주소만 사용
                            if "hubspot" in base or "hs-sites" in base:
                                return base

                            return base

                        for u in raw_links:
                            if _is_safe_domain(u, safe_domains) or _is_image_url(u):
                                continue
                            try:
                                # 스마트 정규화 적용 (인코딩 내부 투시)
                                base_final = _smart_normalize(u)

                                # 추출된 최종 주소가 안전 도메인이면 스킵
                                if _is_safe_domain(base_final, safe_domains) or _is_image_url(base_final):
                                    continue
                                    
                                p = urlparse(base_final)
                                d = (p.hostname or '').lower()
                                if not d: continue

                                if d not in dom_counts: dom_counts[d] = 0
                                if dom_counts[d] >= 5: continue 

                                if base_final not in unique_set:
                                    unique_set.add(base_final)
                                    dom_counts[d] += 1
                            except Exception:
                                unique_set.add(u)

                        filtered = sorted(list(unique_set))
                        if filtered:
                            urls_file = os.path.join(no_attach_dir, "urls.txt")
                            with open(urls_file, "w", encoding="utf-8") as uf:
                                for u in filtered:
                                    uf.write(u + "\n")
                            print(f"[+] {len(filtered)}개 의심 URL 정규화 추출 → urls.txt 저장 (raw: {len(raw_links)}개)")
                        else:
                            print(f"[-] 필터링 후 의심 URL 없음 (raw: {len(raw_links)}개)")
                    except Exception as url_ex:
                        print(f"[!] 본문 URL 수집 실패: {url_ex}")

                except Exception as e:
                    print(f"[!] 폴더 생성 실패: {e}")

            else:
                print(f"[*] 첨부파일 {len(attach_wrap)}개 발견")
                for aidx, attach_li in enumerate(attach_wrap):
                    try:
                        # EML이든 아니든 무조건 다운로드 (확장자 기반 분리 로직 추가)
                        try:
                            name_span = attach_li.find_element(By.CSS_SELECTOR, 'span[evt-rol="download-attach"]')
                            raw_filename = name_span.text.strip()
                        except Exception:
                            raw_filename = f"unknown_{mail_id}_{aidx}.ext"
                            
                        filename = sanitize_filename(raw_filename)
                        
                        # 1. 파일이 존재하는지 (기존 파일) 검사 (다운로드 전 판단)
                        # 여기서는 임시로 원래 download_dir를 기준으로만 확인합니다.
                        # (추후 이동될 수 있는 위치의 존재 여부도 확인 필요)
                        
                        # 기존 파일 목록 스냅샷
                        before_files = set(os.listdir(download_dir))
                        
                        # 다운로드 클릭
                        try:
                            name_span = attach_li.find_element(By.CSS_SELECTOR, 'span[evt-rol="download-attach"]')
                            driver.execute_script("arguments[0].click();", name_span)
                            print(f"[*] 다운로드 요청 전송 완료: {filename}")
                        except Exception as e:
                            print(f"[!] 다운로드 클릭 실패: {e}")
                            continue
                        
                        # 다운로드 완료 대기 (최대 30초)
                        download_complete = False
                        downloaded_file_name_in_dir = None
                        
                        for wait in range(30):
                            time.sleep(1)
                            current_files = set(os.listdir(download_dir))
                            new_files = current_files - before_files
                            
                            is_downloading = any(
                                f.endswith(".crdownload") or f.endswith(".tmp") 
                                for f in current_files
                            )
                            
                            if new_files and not is_downloading:
                                for nf in new_files:
                                    downloaded_file_name_in_dir = nf
                                    download_complete = True
                                    break
                            
                            if download_complete:
                                break
                                
                        if not download_complete:
                            # 팝업 처리 (용량이 크거나 문제시 확인창)
                            try:
                                popup = driver.find_element(By.CSS_SELECTOR, "[class*='popup']")
                                if popup.is_displayed():
                                    confirm_btns = popup.find_elements(By.CSS_SELECTOR, "button, a, [class*='btn']")
                                    for cb in confirm_btns:
                                        if "확인" in (cb.text or "") or "저장" in (cb.text or ""):
                                            driver.execute_script("arguments[0].click();", cb)
                                            print("[*] 다운로드 팝업 확인 클릭")
                                            time.sleep(3)
                                            break
                            except Exception:
                                pass
                            
                            # 재확인
                            current_files = set(os.listdir(download_dir))
                            new_files = current_files - before_files
                            for nf in new_files:
                                downloaded_file_name_in_dir = nf
                                download_complete = True
                                break
                                
                        if not download_complete:
                            print(f"[!] 다운로드 시간 초과: {filename}")
                            continue

                        # 다운로드에 성공한 파일을 성격에 맞게 분류
                        if downloaded_file_name_in_dir:
                            final_ext = downloaded_file_name_in_dir.lower()
                            src_path = os.path.join(download_dir, downloaded_file_name_in_dir)
                            
                            # --- 해시 중복 체크 로직 추가 ---
                            hash_history_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachfiles", "downloaded_hash_history.txt")
                            try:
                                with open(src_path, 'rb') as f:
                                    file_hash = hashlib.sha256(f.read()).hexdigest()
                                
                                is_duplicate = False
                                if os.path.exists(hash_history_file):
                                    with open(hash_history_file, 'r', encoding='utf-8') as hf:
                                        if file_hash in hf.read():
                                            is_duplicate = True
                                
                                if is_duplicate:
                                    print(f"  [-] 중복 파일 감지(해시 일치). 다운로드 취소: {downloaded_file_name_in_dir}")
                                    try:
                                        os.remove(src_path)
                                    except Exception:
                                        pass
                                    continue
                                else:
                                    with open(hash_history_file, 'a', encoding='utf-8') as hf:
                                        hf.write(f"{file_hash}\n")
                            except Exception as e:
                                print(f"[!] 해시 계산 실패: {e}")
                            
                            if final_ext.endswith(".eml"):
                                downloaded_count += 1
                                print(f"[SUCCESS] EML 다운로드 완료: {downloaded_file_name_in_dir}")
                                eml_found = True
                                # .eml은 기본 download_dir에 그대로 둡니다.
                            else:
                                # 일반 첨부파일: attachfiles 폴더 하위에 [메일제목] 폴더를 생성하여 격리
                                attach_base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachfiles")
                                safe_title = sanitize_filename(title)
                                if not safe_title:
                                    safe_title = "Unknown_Subject"
                                    
                                target_sub_dir = os.path.join(attach_base_dir, f"[은행신고]{safe_title}")
                                os.makedirs(target_sub_dir, exist_ok=True)
                                
                                dest_path = os.path.join(target_sub_dir, downloaded_file_name_in_dir)
                                
                                # 이동 (덮어쓰기 허용방식)
                                try:
                                    import shutil
                                    if os.path.exists(dest_path):
                                        os.remove(dest_path)
                                    shutil.move(src_path, dest_path)
                                    print(f"[SUCCESS] 일반 파일 이동 완료: {safe_title} / {downloaded_file_name_in_dir}")
                                except Exception as e:
                                    print(f"[!] 일반 파일 이동 실패: {e}")
                    except Exception as e:
                        print(f"[!] 첨부파일 루프 내 예외 발생: {e}")
                        continue
                                    
                    except Exception as e:
                        print(f"[!] 첨부파일 처리 오류: {e}")
                        continue
        except Exception as e:
            print(f"[!] 첨부파일 목록 확인 실패: {e}")
        
        if not eml_found:
            skipped_count += 1
            
        # 처리 이력 저장 (다운로드 성공했거나, 첨부파일이 없어서 확인 완료된 경우 모두 포함)
        try:
            with open(history_file, "a", encoding="utf-8") as f:
                f.write(f"{mail_id}\n")
            processed_ids.add(mail_id)
        except Exception as e:
            print(f"[!] 이력 저장 실패: {e}")
        
        # 목록으로 돌아가기
        print("[*] 메일 목록으로 돌아가기...")
        try:
            # '목록' 버튼 찾기
            list_btn = None
            btns = driver.find_elements(By.CSS_SELECTOR, "a, button")
            for btn in btns:
                try:
                    t = btn.text.strip() if btn.text else ""
                    evt = btn.get_attribute("evt-rol") or ""
                    if "목록" in t or "list" in evt.lower():
                        list_btn = btn
                        break
                except Exception:
                    continue
            
            if list_btn:
                driver.execute_script("arguments[0].click();", list_btn)
                print("[+] 목록 버튼 클릭")
            else:
                driver.back()
                print("[*] 브라우저 Back 사용")
            
            time.sleep(3)  # 목록 로딩 대기
            
            # 목록 페이지 확인 (state=1)
            if "state=1" not in driver.current_url:
                print("[*] 목록 URL이 아닙니다. 새로고침...")
                driver.get("https://mail.shinhan.com/mail/mail/mailCommon.do?state=1")
                time.sleep(5)
        except Exception as e:
            print(f"[!] 목록 복귀 오류: {e}. 새로고침합니다.")
            driver.get("https://mail.shinhan.com/mail/mail/mailCommon.do?state=1")
            time.sleep(5)
    
    # 결과 요약
    print("\n" + "="*50)
    print(f"  [결과] 전체 메일: {len(mail_ids)}개")
    print(f"  [결과] 다운로드한 EML: {downloaded_count}개")
    print(f"  [결과] EML 없는 메일: {skipped_count}개")
    print(f"  [결과] 저장 위치: {download_dir}")
    print("="*50)


# ─────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────
def main():
    import argparse
    arg_parser = argparse.ArgumentParser(description="mail.shinhan.com 자동 로그인")
    arg_parser.add_argument("--auto-close", action="store_true",
                            help="다운로드 완료 후 브라우저를 자동 종료합니다 (파이프라인 모드)")
    args = arg_parser.parse_args()

    print("=" * 50)
    print("  mail.shinhan.com 자동 로그인")
    print("  Selenium + Gmail IMAP 2FA")
    print("=" * 50)
    print()
    
    config = load_config()
    
    # 설정 확인 출력
    print(f"[*] 대상 URL: {config.get('shinhan_mail', 'url')}")
    print(f"[*] 사용자 ID: {config.get('shinhan_mail', 'username')}")
    print(f"[*] Gmail: {config.get('gmail_imap', 'email')}")
    print(f"[*] 브라우저: {config.get('browser', 'browser_type', fallback='chrome')}")
    print(f"[*] Headless: {config.get('browser', 'headless', fallback='False')}")
    print()
    
    driver = login_shinhan_mail(config)
    
    if driver:
        print("\n[SUCCESS] 자동 로그인이 완료되었습니다!")
        
        # EML 첨부파일 다운로드
        try:
            download_eml_attachments(driver)
        except Exception as e:
            print(f"[!] EML 다운로드 중 오류: {e}")
        
        if args.auto_close:
            # 파이프라인 모드: 다운로드 완료 후 자동으로 브라우저 종료
            print("\n[*] --auto-close 모드: 브라우저를 자동 종료합니다.")
            try:
                driver.quit()
            except Exception:
                pass
            print("[*] 브라우저 종료.")
        else:
            # 수동 모드: 사용자 입력 대기
            print("\n[*] 브라우저를 유지합니다. 수동으로 닫아주세요.")
            print("[*] (자동 종료를 원하시면 Enter를 누르세요)")
            try:
                input()
                driver.quit()
                print("[*] 브라우저 종료.")
            except (KeyboardInterrupt, EOFError):
                pass
    else:
        print("\n[FAIL] 로그인에 실패했습니다. config.ini 설정을 확인해주세요.")
        print("   특히 CSS 셀렉터가 실제 페이지 구조와 맞는지 확인이 필요합니다.")
        print("   브라우저 개발자 도구(F12)로 요소의 ID, Name을 확인하세요.")
        if not args.auto_close:
            sys.exit(1)
        else:
            sys.exit(1)


if __name__ == "__main__":
    main()
