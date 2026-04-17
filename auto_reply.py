#!/usr/bin/env python3
"""
악성메일 자동 회신 프로그램 (v1.0)

흐름:
  1. 자동 로그인 (auto_login.py 재사용)
  2. 작업중 편지함 이동
  3. 회신할 건수 입력받기
  4. 메일 목록 수집 (입력 건수만큼)
  5. 각 메일에 대해:
     a. 메일 행 클릭 (본문 열람)
     b. 답장 버튼 클릭 (답장 작성 화면 진입)
     c. 보내기 버튼 클릭 (회신 발송)
     d. 목록으로 복귀
  6. 결과 요약 출력 및 종료

탐지된 셀렉터:
  - 답장: <div class="btn_submenu"> > <a class="btn_tool btn_tool_multi"> 텍스트='답장'
  - 보내기: <a class="btn_major_s" evt-rol="send-message"> 텍스트='보내기'
  - 취소:  evt-rol='toolbar-write-cancel'
"""

import os
import sys
import time
import re
from datetime import datetime

os.environ['PYTHONUTF8'] = '1'
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

try:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
except ImportError:
    print("[!] pip install selenium")
    sys.exit(1)

# ─────────────────────────────────────────────
# 임포트 경로 설정 (EXE 실행 대응)
# ─────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

if base_path not in sys.path:
    sys.path.insert(0, base_path)

from auto_login import load_config, login_shinhan_mail


# ─────────────────────────────────────────────
# 전역 변수
# ─────────────────────────────────────────────
MALMAIL_URL = ""
DRY_RUN = False  # True이면 보내기 버튼 클릭을 건너뜀 (테스트용)

EXCLUDE_ID_RE = re.compile(
    r'^(allSelectTr|dateDesc_|dateAsc_|toolbar_|pageNavi)',
    re.IGNORECASE
)


# ─────────────────────────────────────────────
# 작업중 편지함 이동
# ─────────────────────────────────────────────
def navigate_to_malmail_folder(driver):
    """작업중 편지함으로 이동합니다."""
    print("\n[*] 작업중 편지함으로 이동 중...")

    # iframe 탐색
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.frame(iframe)
            if _click_malmail_element(driver):
                driver.switch_to.default_content()
                try:
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr[id]"))
                    )
                except TimeoutException:
                    pass
                time.sleep(0.5)
                return True
            driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()

    # 기본 컨텍스트
    driver.switch_to.default_content()
    if _click_malmail_element(driver):
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr[id]"))
            )
        except TimeoutException:
            pass
        time.sleep(0.5)
        return True

    print("[!] 작업중 폴더 찾기 실패. 폴더 목록:")
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, "li[evt-rol='folder'] a"):
            print(f"  - '{el.text.strip()}'")
    except Exception:
        pass
    return False


def _click_malmail_element(driver):
    for kw in ["작업중"]:
        try:
            for el in driver.find_elements(By.XPATH, f"//*[contains(text(),'{kw}')]"):
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    print(f"[+] '{kw}' 클릭 완료")
                    return True
        except Exception:
            continue
    return False


# ─────────────────────────────────────────────
# 유틸리티 함수
# ─────────────────────────────────────────────
def _find_row_by_id(driver, mail_id):
    """JavaScript querySelector로 특수문자 포함 ID를 안전하게 탐색합니다."""
    try:
        js = """
        var mid = arguments[0];
        try {
            return document.querySelector('tr[id="' + mid + '"]');
        } catch(e) {
            var rows = document.querySelectorAll('table tbody tr[id]');
            for (var i = 0; i < rows.length; i++) {
                if (rows[i].id === mid) return rows[i];
            }
            return null;
        }
        """
        return driver.execute_script(js, mail_id)
    except Exception:
        return None


def _set_page_size(driver, value="80"):
    """페이지당 표시 건수를 설정합니다."""
    try:
        from selenium.webdriver.support.ui import Select
        sel = Select(driver.find_element(By.CSS_SELECTOR, "#toolbar_list_pagebase"))
        if sel.first_selected_option.get_attribute("value") != value:
            sel.select_by_value(value)
            print(f"[*] 페이지당 {value}건으로 변경")
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr[id]"))
                )
            except TimeoutException:
                pass
            time.sleep(1)
    except Exception as e:
        print(f"[!] 페이지당 건수 변경 실패: {e}")


def _handle_confirm_popup(driver):
    """확인 팝업(alert 또는 HTML 팝업) 처리"""
    time.sleep(1)
    try:
        # alert 처리
        try:
            alert = driver.switch_to.alert
            print(f"  [*] Alert 팝업 감지: '{alert.text}'")
            alert.accept()
            print("  [*] Alert 확인 클릭")
            time.sleep(1)
            return True
        except Exception:
            pass

        # HTML 팝업의 확인/예 버튼 처리
        confirm_keywords = ["확인", "예", "OK", "Yes"]
        for kw in confirm_keywords:
            try:
                for el in driver.find_elements(By.XPATH, f"//*[contains(text(),'{kw}')]"):
                    if el.is_displayed():
                        tag = el.tag_name
                        if tag in ["button", "a", "span", "input"]:
                            driver.execute_script("arguments[0].click();", el)
                            print(f"  [*] 확인 팝업 '{kw}' 클릭")
                            time.sleep(1)
                            return True
            except Exception:
                continue
    except Exception:
        pass
    return False


# ─────────────────────────────────────────────
# 메일 ID 수집 (페이지네이션)
# ─────────────────────────────────────────────
def collect_all_mail_ids(driver, target_limit=None):
    """메일 목록에서 ID를 수집합니다."""
    print("\n[*] 메일 목록 수집 중...")
    _set_page_size(driver, "80")

    all_ids = []
    page = 1

    while True:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr[id]"))
            )
        except TimeoutException:
            pass
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr[id]")
        mail_ids = [
            r.get_attribute("id") for r in rows
            if r.get_attribute("id") and not EXCLUDE_ID_RE.match(r.get_attribute("id"))
        ]
        print(f"[*] 페이지 {page}: {len(mail_ids)}개 메일")
        if not mail_ids and page == 1:
            print("[-] 메일 없음")
            return []
        for mid in mail_ids:
            if mid not in all_ids:
                all_ids.append(mid)
                if target_limit and len(all_ids) >= target_limit:
                    print(f"[*] 목표 건수({target_limit}건) 달성. 목록 수집 조기 종료.")
                    return all_ids[:target_limit]

        try:
            nb = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.next.paginate_button")
            if "paginate_button_disabled" in (nb.get_attribute("class") or ""):
                break
            driver.execute_script("arguments[0].click();", nb)
            page += 1
            try:
                WebDriverWait(driver, 5).until(EC.staleness_of(rows[0]))
            except Exception:
                pass
            time.sleep(0.5)
        except Exception:
            break

    print(f"[*] 총 {len(all_ids)}개 수집 완료")
    return all_ids


# ─────────────────────────────────────────────
# 메일 열람 (본문 진입)
# ─────────────────────────────────────────────
def open_mail(driver, mail_id):
    """메일 행을 클릭하여 본문으로 진입합니다."""
    row = _find_row_by_id(driver, mail_id)
    if row is None:
        print(f"  [!] 메일 행 못 찾음: {mail_id}")
        return False

    # 제목 출력
    try:
        for lk in row.find_elements(By.TAG_NAME, "a"):
            txt = (lk.text or "").strip()
            if txt and len(txt) > 2 and "첨부파일" not in txt:
                # 말머리 제거
                prefix_pattern = re.compile(r'^\[(신고메일|Report email|举报邮件)\]\s*', re.IGNORECASE)
                cleaned = prefix_pattern.sub('', txt[:120])
                print(f"  제목: {cleaned[:60]}")
                break
    except Exception:
        pass

    # 메일 클릭
    clicked = False
    try:
        click_targets = row.find_elements(By.CSS_SELECTOR, "a.subject, .mail_subject, td.subject a, span.subject")
        if not click_targets:
            click_targets = row.find_elements(By.TAG_NAME, "td")

        for target in click_targets:
            if target.is_displayed() and target.text.strip():
                driver.execute_script("arguments[0].click();", target)
                clicked = True
                break

        if not clicked:
            driver.execute_script("arguments[0].click();", row)
            clicked = True
    except Exception as e:
        print(f"  [!] 메일 클릭 실패: {e}")
        return False

    if not clicked:
        print("  [!] 클릭 요소 없음")
        return False

    # 본문 로드 대기
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "span[evt-rol='read-nested-pop'], .btn_fn4, .mail_read_wrap, .mail_header, div.btn_submenu"))
        )
    except TimeoutException:
        pass
    time.sleep(2)

    return True


# ─────────────────────────────────────────────
# 답장 버튼 클릭
# ─────────────────────────────────────────────
def click_reply_button(driver):
    """
    메일 열람 화면에서 답장 버튼을 클릭하여 답장 작성 화면에 진입합니다.
    탐지된 셀렉터: <div class="btn_submenu"> 내부 <a class="btn_tool btn_tool_multi"> 텍스트='답장'
    """
    reply_clicked = False

    # 방법 1: btn_submenu 내부의 답장 텍스트를 가진 a 태그
    try:
        submenu_divs = driver.find_elements(By.CSS_SELECTOR, "div.btn_submenu")
        for div in submenu_divs:
            try:
                text = (div.text or "").strip()
                if text == "답장":
                    link = div.find_element(By.TAG_NAME, "a")
                    if link and link.is_displayed():
                        driver.execute_script("arguments[0].click();", link)
                        reply_clicked = True
                        print("  [+] 답장 버튼 클릭 (btn_submenu > a)")
                        break
            except Exception:
                continue
    except Exception:
        pass

    # 방법 2: 텍스트 '답장'을 가진 a.btn_tool 직접 탐색
    if not reply_clicked:
        try:
            links = driver.find_elements(By.CSS_SELECTOR, "a.btn_tool")
            for link in links:
                text = (link.text or "").strip()
                if text == "답장" and link.is_displayed():
                    driver.execute_script("arguments[0].click();", link)
                    reply_clicked = True
                    print("  [+] 답장 버튼 클릭 (a.btn_tool)")
                    break
        except Exception:
            pass

    # 방법 3: XPath fallback
    if not reply_clicked:
        try:
            elements = driver.find_elements(By.XPATH,
                "//a[text()='답장'] | //span[text()='답장'] | //button[text()='답장']")
            for el in elements:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    reply_clicked = True
                    print("  [+] 답장 버튼 클릭 (XPath)")
                    break
        except Exception:
            pass

    if reply_clicked:
        # 답장 작성 화면 로드 대기
        try:
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "a.btn_major_s[evt-rol='send-message']"))
            )
        except TimeoutException:
            pass
        time.sleep(2)
    else:
        print("  [!] 답장 버튼을 찾을 수 없습니다.")

    return reply_clicked


# ─────────────────────────────────────────────
# 보내기 버튼 클릭
# ─────────────────────────────────────────────
def click_send_button(driver):
    """
    답장 작성 화면에서 보내기 버튼을 클릭합니다.
    
    프로세스:
      1. 답장 작성 화면에서 보내기 버튼 클릭 (evt-rol='send-message')
      2. 팝업이 나타나면:
         a. 전체체크 버튼 클릭 (evt-rol='send-check-all')
         b. 팝업 내 보내기 버튼 클릭 (footer.btn_layer_wrap > a.btn_major_s)
    
    DRY_RUN이 True이면 클릭을 건너뜁니다.
    """
    global DRY_RUN

    if DRY_RUN:
        print("  [DRY_RUN] 보내기 버튼 클릭 건너뜀 (테스트 모드)")
        return "DRY_RUN"

    # ── STEP 1: 답장 작성 화면의 보내기 버튼 클릭 ──
    send_clicked = False

    # 방법 1: evt-rol="send-message" 셀렉터 (가장 정확)
    try:
        btns = driver.find_elements(By.CSS_SELECTOR, "a.btn_major_s[evt-rol='send-message']")
        for btn in btns:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                send_clicked = True
                print("  [+] 1단계: 보내기 버튼 클릭 (evt-rol='send-message')")
                break
    except Exception:
        pass

    # 방법 2: 텍스트 '보내기'를 가진 a.btn_major_s
    if not send_clicked:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, "a.btn_major_s")
            for btn in btns:
                text = (btn.text or "").strip()
                if text == "보내기" and btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    send_clicked = True
                    print("  [+] 1단계: 보내기 버튼 클릭 (a.btn_major_s)")
                    break
        except Exception:
            pass

    if not send_clicked:
        print("  [!] 보내기 버튼을 찾을 수 없습니다.")
        return "FAIL"

    # ── STEP 2: 팝업 대기 후 전체체크 + 팝업 보내기 ──
    time.sleep(3)

    # 팝업이 나타났는지 확인 (전체체크 버튼 존재 여부)
    popup_handled = False
    try:
        check_all_btns = driver.find_elements(By.CSS_SELECTOR, "[evt-rol='send-check-all']")
        for btn in check_all_btns:
            if btn.is_displayed():
                # 전체체크 버튼 클릭
                driver.execute_script("arguments[0].click();", btn)
                print("  [+] 2단계: 전체체크 버튼 클릭 (evt-rol='send-check-all')")
                time.sleep(1)

                # 팝업 내 보내기 버튼 클릭 (footer.btn_layer_wrap 안의 a.btn_major_s)
                popup_send_clicked = False
                try:
                    # footer.btn_layer_wrap 내부의 보내기 버튼
                    footer_btns = driver.find_elements(By.CSS_SELECTOR, "footer.btn_layer_wrap a.btn_major_s")
                    for fb in footer_btns:
                        if fb.is_displayed():
                            text = (fb.text or "").strip()
                            if text == "보내기":
                                driver.execute_script("arguments[0].click();", fb)
                                popup_send_clicked = True
                                print("  [+] 3단계: 팝업 보내기 버튼 클릭 (footer > a.btn_major_s)")
                                break
                except Exception:
                    pass

                # fallback: footer 내부가 아닌 일반 보내기 버튼 (팝업 영역 내)
                if not popup_send_clicked:
                    try:
                        all_send_btns = driver.find_elements(By.CSS_SELECTOR, "a.btn_major_s")
                        for sb in all_send_btns:
                            text = (sb.text or "").strip()
                            # evt-rol이 없는 보내기 버튼 = 팝업 내 보내기
                            if text == "보내기" and sb.is_displayed() and not sb.get_attribute("evt-rol"):
                                driver.execute_script("arguments[0].click();", sb)
                                popup_send_clicked = True
                                print("  [+] 3단계: 팝업 보내기 버튼 클릭 (a.btn_major_s, 팝업)")
                                break
                    except Exception:
                        pass

                if popup_send_clicked:
                    popup_handled = True
                else:
                    print("  [!] 팝업 내 보내기 버튼을 찾지 못했습니다.")
                    return "FAIL"
                break
    except Exception:
        pass

    # 팝업이 없는 경우 (확인 팝업만 있거나 바로 전송되는 경우)
    if not popup_handled:
        print("  [*] 전체체크 팝업 미감지 → 일반 확인 팝업 처리")
        _handle_confirm_popup(driver)

    # 전송 완료 대기
    time.sleep(3)
    # 혹시 추가 확인 팝업이 뜨면 처리
    _handle_confirm_popup(driver)

    return "SUCCESS"


# ─────────────────────────────────────────────
# 답장 취소 (DRY_RUN 모드 또는 오류 시)
# ─────────────────────────────────────────────
def cancel_reply(driver):
    """답장 작성 화면에서 안전하게 빠져나옵니다."""
    cancelled = False

    # 방법 1: evt-rol 기반 취소/새로쓰기 버튼
    try:
        cancel_rols = ["toolbar-write-cancel", "cancel", "write-cancel", "cancel-write"]
        for rol in cancel_rols:
            els = driver.find_elements(By.CSS_SELECTOR, f"[evt-rol*='{rol}']")
            for el in els:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    print(f"  [*] evt-rol='{rol}' 버튼으로 답장 취소")
                    time.sleep(2)
                    _handle_confirm_popup(driver)
                    cancelled = True
                    break
            if cancelled:
                break
    except Exception:
        pass

    # 방법 2: '취소' 텍스트 버튼
    if not cancelled:
        try:
            for el in driver.find_elements(By.XPATH, "//*[contains(text(),'취소')]"):
                if el.is_displayed():
                    text = (el.text or "").strip()
                    if text in ["취소", "쓰기취소"]:
                        driver.execute_script("arguments[0].click();", el)
                        print(f"  [*] '{text}' 버튼으로 답장 취소")
                        time.sleep(2)
                        _handle_confirm_popup(driver)
                        cancelled = True
                        break
        except Exception:
            pass

    # 방법 3: 브라우저 뒤로가기
    if not cancelled:
        try:
            driver.back()
            print("  [*] 브라우저 back()으로 답장 취소")
            time.sleep(3)
            _handle_confirm_popup(driver)
            cancelled = True
        except Exception:
            pass

    return cancelled


# ─────────────────────────────────────────────
# 목록 복귀
# ─────────────────────────────────────────────
def go_back_to_list(driver):
    """
    메일 본문 또는 답장 완료 후 작업중 목록으로 복귀합니다.
    """
    global MALMAIL_URL
    navigated = False

    # 1차: 좌측 사이드바에서 작업중 폴더 클릭
    for kw in ["작업중"]:
        try:
            elems = driver.find_elements(By.XPATH, f"//*[contains(text(),'{kw}')]")
            for el in elems:
                try:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].click();", el)
                        print(f"  [*] 사이드바 '{kw}' 폴더 클릭으로 복귀")
                        navigated = True
                        time.sleep(4)
                        break
                except Exception:
                    continue
            if navigated:
                break
        except Exception:
            continue

    # 2차: 폴더 클릭 실패 시 URL 직접 이동
    if not navigated:
        if MALMAIL_URL:
            print("  [*] 폴더 클릭 실패 → URL 직접 이동")
            driver.get(MALMAIL_URL)
            time.sleep(4)
        else:
            driver.back()
            print("  [*] 브라우저 Back")
            time.sleep(3)


# ─────────────────────────────────────────────
# 단일 메일 회신 처리
# ─────────────────────────────────────────────
def reply_to_mail(driver, mail_id, idx, total):
    """
    하나의 메일에 대해 열람 → 답장 → 보내기를 수행합니다.
    
    Returns:
        str: "SUCCESS", "DRY_RUN", "FAIL_OPEN", "FAIL_REPLY", "FAIL_SEND"
    """
    print(f"\n{'─' * 55}")
    print(f"[*] 메일 {idx}/{total}  (ID: {mail_id})")

    # ── 1. 메일 행이 현재 페이지에 없으면 탐색 ──
    if _find_row_by_id(driver, mail_id) is None:
        print("  [*] 현재 페이지에 행 없음 → 다음 페이지 이동 탐색")
        found = False
        for _ in range(30):
            try:
                nb = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.next.paginate_button")
                if "paginate_button_disabled" in (nb.get_attribute("class") or ""):
                    break
                driver.execute_script("arguments[0].click();", nb)
                time.sleep(3)
                if _find_row_by_id(driver, mail_id):
                    found = True
                    print("  [*] 다음 페이지에서 행 발견")
                    break
            except Exception:
                break

        if not found:
            print("  [*] 탐색 실패, 악성메일함 초기화 후 재탐색")
            driver.get(MALMAIL_URL)
            time.sleep(5)
            _set_page_size(driver, "80")
            time.sleep(2)
            for _ in range(30):
                if _find_row_by_id(driver, mail_id):
                    found = True
                    print("  [*] 재로딩 후 행 발견")
                    break
                try:
                    nb = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.next.paginate_button")
                    if "paginate_button_disabled" in (nb.get_attribute("class") or ""):
                        break
                    driver.execute_script("arguments[0].click();", nb)
                    time.sleep(3)
                except Exception:
                    break

    # ── 2. 메일 열람 ──
    if not open_mail(driver, mail_id):
        return "FAIL_OPEN"

    # ── 3. 답장 버튼 클릭 ──
    if not click_reply_button(driver):
        print("  [!] 답장 실패 → 목록으로 복귀")
        go_back_to_list(driver)
        time.sleep(1)
        return "FAIL_REPLY"

    # ── 4. 보내기 버튼 클릭 ──
    result = click_send_button(driver)

    if result == "DRY_RUN":
        # 드라이런 모드: 답장 취소 후 복귀
        cancel_reply(driver)
        time.sleep(1)
        go_back_to_list(driver)
        time.sleep(1)
        return "DRY_RUN"
    elif result == "SUCCESS":
        # 전송 완료 후 목록 복귀
        time.sleep(2)
        go_back_to_list(driver)
        time.sleep(1)
        return "SUCCESS"
    else:
        # 보내기 실패: 답장 취소 후 복귀
        cancel_reply(driver)
        time.sleep(1)
        go_back_to_list(driver)
        time.sleep(1)
        return "FAIL_SEND"


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def main():
    global DRY_RUN, MALMAIL_URL

    print("=" * 60)
    print("  📧 악성메일 자동 회신 프로그램 v1.0")
    print("=" * 60)

    # ── 1. 회신 건수 입력 ──
    if len(sys.argv) > 1:
        limit_input = sys.argv[1].strip()
        print(f"  [*] CLI 입력 감지: {limit_input}")
    else:
        limit_input = input("  회신할 건수를 입력해주세요 (전체 회신은 Enter): ").strip()

    target_limit = None
    if limit_input.isdigit() and int(limit_input) > 0:
        target_limit = int(limit_input)
        print(f"  [*] 목표 회신 건수: {target_limit}건")
    else:
        print("  [*] 목표 회신 건수: 전체")

    # ── DRY_RUN 모드 확인 ──
    if "--dry-run" in sys.argv:
        DRY_RUN = True
        print("  ⚠️ DRY_RUN 모드: 보내기 버튼은 클릭하지 않습니다.")

    print("=" * 60)

    # ── 2. 로그인 ──
    config = load_config()
    driver = login_shinhan_mail(config)
    if not driver:
        print("[FAIL] 로그인 실패")
        sys.exit(1)

    driver.implicitly_wait(0)
    print("[SUCCESS] 로그인 완료!\n")

    results = {"SUCCESS": 0, "DRY_RUN": 0, "FAIL_OPEN": 0, "FAIL_REPLY": 0, "FAIL_SEND": 0}

    try:
        # ── 3. 작업중 편지함 이동 ──
        if not navigate_to_malmail_folder(driver):
            print("[!] 작업중 폴더 이동 실패")
            input("Enter 후 종료...")
            driver.quit()
            sys.exit(1)

        time.sleep(3)
        MALMAIL_URL = driver.current_url
        print(f"[*] 작업중 폴더 URL: {MALMAIL_URL}")

        # ── 4. 메일 ID 수집 ──
        mail_ids = collect_all_mail_ids(driver, target_limit)

        if target_limit and len(mail_ids) > target_limit:
            mail_ids = mail_ids[:target_limit]

        if not mail_ids:
            print("[!] 메일 없음")
            driver.quit()
            sys.exit(0)

        # 첫 페이지 복귀
        try:
            fb = driver.find_element(By.CSS_SELECTOR, "#pageNaviWrap a.first.paginate_button")
            if fb.is_displayed() and "paginate_button_disabled" not in (fb.get_attribute("class") or ""):
                driver.execute_script("arguments[0].click();", fb)
                time.sleep(1.5)
        except Exception:
            pass

        total = len(mail_ids)
        print(f"\n[*] 총 {total}개 메일 회신 처리 시작\n")

        # ── 5. 메일별 회신 루프 ──
        for idx, mail_id in enumerate(mail_ids, 1):
            result = reply_to_mail(driver, mail_id, idx, total)
            results[result] = results.get(result, 0) + 1
            print(f"  결과: {result}")

        # ── 6. 결과 요약 ──
        print("\n" + "=" * 60)
        print(f"  📧 회신 완료 요약")
        print("=" * 60)
        print(f"  총 대상: {total}건")
        print(f"  ✅ 성공:       {results['SUCCESS']}건")
        if results['DRY_RUN'] > 0:
            print(f"  🔍 DRY_RUN:    {results['DRY_RUN']}건 (보내기 건너뜀)")
        if results['FAIL_OPEN'] > 0:
            print(f"  ❌ 열람 실패:  {results['FAIL_OPEN']}건")
        if results['FAIL_REPLY'] > 0:
            print(f"  ❌ 답장 실패:  {results['FAIL_REPLY']}건")
        if results['FAIL_SEND'] > 0:
            print(f"  ❌ 전송 실패:  {results['FAIL_SEND']}건")
        print("=" * 60)

    except Exception as e:
        print(f"\n[!] 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        print("[*] 브라우저 종료")


if __name__ == "__main__":
    main()
