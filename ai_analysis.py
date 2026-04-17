#!/usr/bin/env python3

# Windows cp949 인코딩 문제 방지 (스크립트 최상단에서 설정)
import os
os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"
# Gemini API 키 자동 연동 (settings.json에서 읽어오기)
import json

settings_path = os.path.expanduser("~/.gemini/settings.json")
try:
    with open(settings_path, "r", encoding="utf-8") as f:
        settings = json.load(f)
        if "apiKey" in settings:
            os.environ["GEMINI_API_KEY"] = settings["apiKey"]
except Exception as e:
    print(f"[!] Cannot read API key from {settings_path}: {e}")

import argparse
import subprocess
import sys
import datetime
import shutil
import asyncio
import glob
import re

try:
    from deep_translator import GoogleTranslator
except ImportError:
    GoogleTranslator = None
    print("[!] deep-translator module not found. Auto-translation to Korean disabled.")

# 표준 출력/에러 인코딩 오류(cp949 변환 불가 문제) 방지 설정 (asyncio 충돌 방지를 위해 reconfigure 사용)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')


# 분석 결과가 저장될 디렉토리 경로
output_path = "./ai_analysis_report/" 

def get_current_cycle_start():
    """현재 시간 기준, 17:00 리셋 주기의 시작 일시를 반환합니다."""
    now = datetime.datetime.now()
    if now.hour < 17:
        # 17시 이전이면, 어제 17시가 주기 시작
        cycle_start = now - datetime.timedelta(days=1)
    else:
        # 17시 이후이면, 오늘 17시가 주기 시작
        cycle_start = now
    return cycle_start.replace(hour=17, minute=0, second=0, microsecond=0)

def log_api_request(filename, status):
    """API 호출 기록을 파일에 남기며, 17시 리셋 기반 누적 카운트를 함께 기록합니다."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(script_dir, "gemini_api_usage.log")
    count_file = os.path.join(script_dir, ".gemini_api_count.txt")
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    current_cycle = get_current_cycle_start()
    cycle_str = current_cycle.strftime("%Y-%m-%d %H:00")
    
    # 1. 이전 카운트 읽기 및 주기 체크
    count = 0
    saved_cycle = ""
    if os.path.exists(count_file):
        try:
            with open(count_file, "r", encoding="utf-8") as cf:
                content = cf.read().strip()
                if "|" in content:
                    saved_cycle, count_str = content.split("|")
                    count = int(count_str)
        except Exception:
            pass
            
    # 주기가 바뀌었으면 카운트 리셋
    if saved_cycle != cycle_str:
        count = 0
        
    count += 1 # 이번 호출 카운트 증가
    
    # 2. 카운트 파일 업데이트
    try:
        with open(count_file, "w", encoding="utf-8") as cf:
            cf.write(f"{cycle_str}|{count}")
    except Exception:
        pass

    # 3. 로그 작성
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[Count: {count}] [{timestamp}] API Request | File: {filename} | Status: {status}\n")
    except Exception as e:
        print(f"[!] Failed to log API request: {e}")

def get_daily_usage_count():
    """현재 17:00 주기 내에서 성공한 호출 수를 파일에서 가져옵니다."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    count_file = os.path.join(script_dir, ".gemini_api_count.txt")
    
    if not os.path.exists(count_file):
        return 0
        
    current_cycle = get_current_cycle_start()
    cycle_str = current_cycle.strftime("%Y-%m-%d %H:00")
    
    try:
        with open(count_file, "r", encoding="utf-8") as cf:
            content = cf.read().strip()
            if "|" in content:
                saved_cycle, count_str = content.split("|")
                if saved_cycle == cycle_str:
                    return int(count_str)
        return 0
    except Exception:
        return 0
# ─────────────────────────────────────────────
# API 키 동적 로테이션 매니저 (Smart Rotation)
# ─────────────────────────────────────────────
import time as _time

class ApiKeyManager:
    """
    API 키의 상태(OK/BLOCKED)와 마지막 사용 시각을 추적하여,
    사용 가능한 키를 동적으로 선택하고 오류를 전용 로그에 기록합니다.
    """
    RPM_PER_KEY = 5  # Gemini 무료 티어: 분당 5회
    MIN_INTERVAL = 12.5  # 60초 / 5RPM = 12초 (안전 마진 포함 12.5초)

    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.settings_path = os.path.expanduser("~/.gemini/settings.json")
        self.api_keys_path = os.path.join(self.script_dir, "api_keys.txt")
        self.error_log_path = os.path.join(self.script_dir, "gemini_api_key_errors.log")
        self.keys_exhausted = False
        
        # 키 로드
        self.keys = self._load_keys()
        # 각 키의 상태: {key: {"status": "OK"|"BLOCKED", "last_used": float, "blocked_at": float|None}}
        self.key_states = {}
        for k in self.keys:
            self.key_states[k] = {"status": "OK", "last_used": 0.0, "blocked_at": None}
        
        print(f"[*] ApiKeyManager 초기화: {len(self.keys)}개 키 로드됨")

    def _load_keys(self):
        """api_keys.txt에서 키 목록을 로드합니다."""
        if not os.path.exists(self.api_keys_path):
            return []
        try:
            with open(self.api_keys_path, 'r', encoding='utf-8') as f:
                return [line.strip() for line in f if line.strip() and not line.startswith('#')]
        except Exception:
            return []

    def _apply_key(self, key):
        """선택된 키를 settings.json과 환경변수에 적용합니다."""
        try:
            data = {}
            if os.path.exists(self.settings_path):
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            data["apiKey"] = key
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            os.environ["GEMINI_API_KEY"] = key
        except Exception as e:
            print(f"[!] 키 적용 실패: {e}")

    def get_available_key(self):
        """
        사용 가능한(OK 상태) 키 중에서, 마지막 사용으로부터 가장 오래된 키를 선택합니다.
        모든 키가 BLOCKED이면 None을 반환하고 keys_exhausted를 True로 설정합니다.
        """
        available = [
            (k, s) for k, s in self.key_states.items() if s["status"] == "OK"
        ]
        if not available:
            self.keys_exhausted = True
            print("[!] 모든 API 키가 차단(BLOCKED) 상태입니다.")
            return None

        # 마지막 사용 시각이 가장 오래된(가장 쉬고 있는) 키 선택
        available.sort(key=lambda x: x[1]["last_used"])
        best_key = available[0][0]
        
        self._apply_key(best_key)
        print(f"[*] Using key: {best_key[:12]}...")
        return best_key

    def mark_used(self, key):
        """성공적으로 사용한 키의 마지막 사용 시각을 갱신합니다."""
        if key in self.key_states:
            self.key_states[key]["last_used"] = _time.time()

    def mark_blocked(self, key, reason="429 Quota Exceeded"):
        """
        429/Quota 에러가 발생한 키를 BLOCKED 상태로 전환하고 오류 로그를 기록합니다.
        """
        if key not in self.key_states:
            return
        self.key_states[key]["status"] = "BLOCKED"
        self.key_states[key]["blocked_at"] = _time.time()
        
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_msg = f"[{timestamp}] KEY: {key[:12]}... | STATUS: BLOCKED | REASON: {reason}"
        print(f"[!] {log_msg}")
        
        # 전용 오류 로그 파일에 기록
        try:
            with open(self.error_log_path, 'a', encoding='utf-8') as f:
                f.write(log_msg + "\n")
        except Exception:
            pass

        # 다음 사용 가능한 키가 있는지 확인하고 로그
        available = [k for k, s in self.key_states.items() if s["status"] == "OK"]
        if available:
            failover_key = available[0]
            try:
                with open(self.error_log_path, 'a', encoding='utf-8') as f:
                    f.write(f"[{timestamp}] KEY: {failover_key[:12]}... | STATUS: OK → ACTIVE (failover)\n")
            except Exception:
                pass
        else:
            self.keys_exhausted = True

    def get_ok_key_count(self):
        """현재 OK 상태인 키의 개수를 반환합니다."""
        return sum(1 for s in self.key_states.values() if s["status"] == "OK")

    def get_optimal_delay(self):
        """
        사용 가능한 키 수에 따라 최적 대기 시간을 계산합니다.
        키가 많을수록 대기 시간이 짧아집니다.
        """
        ok_count = self.get_ok_key_count()
        if ok_count <= 0:
            return 13  # fallback
        # 각 키가 분당 5회 가능 → 총 분당 (ok_count * 5)회 가능
        # 안전 마진 20% 추가
        delay = (60.0 / (ok_count * self.RPM_PER_KEY)) * 1.2
        # 최소 2초, 최대 13초로 클램핑
        return max(2.0, min(delay, 13.0))

    def get_wait_for_key(self, key):
        """
        특정 키를 사용하려면 얼마나 기다려야 하는지 계산합니다.
        마지막 사용 시각으로부터 MIN_INTERVAL이 경과하지 않았으면 남은 시간을 반환합니다.
        """
        if key not in self.key_states:
            return 0
        elapsed = _time.time() - self.key_states[key]["last_used"]
        if elapsed >= self.MIN_INTERVAL:
            return 0
        return self.MIN_INTERVAL - elapsed


# 전역 매니저 인스턴스 생성
key_manager = ApiKeyManager()

# 하위 호환성을 위한 플래그 (기존 코드에서 참조하는 부분 대응)
api_keys_exhausted = False

def translate_if_english(text):
    """지나치게 많은 영어가 포함된 AI 응답을 5000자 제한을 피해 청크 단위로 분할하여 한글로 번역합니다."""
    if not GoogleTranslator or not text.strip():
        return text
        
    korean_chars = len(re.findall(r'[가-힣]', text))
    english_chars = len(re.findall(r'[a-zA-Z]', text))
    
    # 영어가 한글보다 훨씬 많은(2배 이상) 경우에만 스크립트가 영어로 출력되었다고 판단하고 번역 진행
    if english_chars > (korean_chars * 2) and english_chars > 100:
        print(f"[*] Gemini output appears to be in English ({english_chars} vs {korean_chars} Korean). Translating to Korean...")
        try:
            translator = GoogleTranslator(source='auto', target='ko')
            
            # API 5000자 제한을 우회하기 위해 문단 단위로 분리
            chunks = text.split('\n\n')
            translated_chunks = []
            
            for chunk in chunks:
                if not chunk.strip():
                    translated_chunks.append("")
                    continue
                
                # 매우 긴 텍스트 분할 (안전장치)
                if len(chunk) > 4000:
                    lines = chunk.split('\n')
                    for line in lines:
                        if line.strip():
                            if len(line) > 4000:
                                translated_chunks.append(translator.translate(line[:4000]))
                                translated_chunks.append(translator.translate(line[4000:]))
                            else:
                                translated_chunks.append(translator.translate(line.strip()))
                else:
                    translated_chunks.append(translator.translate(chunk))
                    
            print("[*] Translation successful.")
            return "\n\n".join(translated_chunks)
        except Exception as e:
            print(f"[!] Background translation failed: {e}")
            return text
            
    return text

async def run_command_async(cmd_list, input_data=None, max_retries=5, retry_delay=20):
    """
    서브프로세스 명령어를 비동기적으로 실행합니다. (동적 키 로테이션 + 재시도 로직)
    반환값: (returncode, stdout, stderr)
    """
    global api_keys_exhausted
    current_key = None
    
    for attempt in range(max_retries + 1):
        try:
            # Gemini 명령어인 경우, 사용 가능한 키를 선택하고 필요 시 대기
            is_gemini = "gemini" in cmd_list[0].lower()
            if is_gemini:
                current_key = key_manager.get_available_key()
                if current_key is None:
                    api_keys_exhausted = True
                    return -1, "", "All API keys exhausted"
                
                # 해당 키의 RPM 제한 대기
                wait_needed = key_manager.get_wait_for_key(current_key)
                if wait_needed > 0:
                    print(f"[*] 키 RPM 대기: {wait_needed:.1f}s (KEY: {current_key[:12]}...)")
                    await asyncio.sleep(wait_needed)
            
            # Windows에서 하위 프로세스도 UTF-8을 사용하도록 환경변수 설정
            env = os.environ.copy()
            env["PYTHONUTF8"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"

            process = await asyncio.create_subprocess_exec(
                *cmd_list,
                stdin=asyncio.subprocess.PIPE if input_data else None,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                env=env
            )
            
            stdout, stderr = await process.communicate(input=input_data.encode('utf-8') if input_data else None)
            stdout_str = stdout.decode('utf-8', errors='replace').strip()
            stderr_str = stderr.decode('utf-8', errors='replace').strip()
            
            # Gemini Quota Exceeded (429) 감지 및 동적 키 전환
            full_output = stdout_str + stderr_str
            if is_gemini and ("429" in full_output or "Quota exceeded" in full_output or "exhausted" in full_output.lower()):
                if attempt < max_retries and current_key:
                    # 현재 키를 BLOCKED 처리
                    reason = "429 Quota Exceeded" if "429" in full_output else "Quota/Key Exhausted"
                    key_manager.mark_blocked(current_key, reason=reason)
                    
                    # 다른 사용 가능한 키가 있으면 즉시 재시도
                    next_key = key_manager.get_available_key()
                    if next_key:
                        print(f"[*] Failover: 다른 키로 즉시 재시도 (Attempt {attempt + 1}/{max_retries})")
                        current_key = next_key
                        continue
                    else:
                        api_keys_exhausted = True
                        key_manager.keys_exhausted = True
                        print("[!] 모든 API 키의 한도가 소진되었습니다.")
                        return -1, "", "All API keys exhausted"
            
            # 성공 시 키 사용 시각 갱신
            if is_gemini and current_key:
                key_manager.mark_used(current_key)
            
            return process.returncode, stdout_str, stderr_str

        except Exception as e:
            if attempt < max_retries:
                print(f"[!] Error: {e}. Retrying in {retry_delay}s...")
                await asyncio.sleep(retry_delay)
                continue
            return -1, "", str(e)
    
    api_keys_exhausted = True
    key_manager.keys_exhausted = True
    print("[!] Max retries exceeded. 모든 API 키의 한도가 소진되었습니다.")
    return -1, "", "Max retries exceeded - API keys exhausted"

async def analyze_file_async(target_file, output_dir_base):
    # 나중에 안전하게 이동하기 위해 절대 경로가 필요합니다.
    target_abs_path = os.path.abspath(target_file)
    filename = os.path.basename(target_abs_path)
    
    if not os.path.exists(target_abs_path):
        print(f"[!] File not found: {target_file}")
        return


    # Define the prompt
    # 동적으로 prompt/file분석.md 파일의 내용을 읽어서 결합합니다.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    prompt_file_path = os.path.join(script_dir, "prompt", "file분석.md")
    
    custom_prompt_content = ""
    if os.path.exists(prompt_file_path):
        try:
            with open(prompt_file_path, "r", encoding="utf-8") as f:
                custom_prompt_content = f.read()
        except Exception as e:
            print(f"[!] Warning: Cannot read prompt file from {prompt_file_path}: {e}")
    else:
        print(f"[!] Warning: Prompt file not found at {prompt_file_path}, using fallback.")
        custom_prompt_content = "다음의 형식을 지켜 보고서를 작성해 주세요."
        
    # --- 추가 프롬프트 자동 병합 로직 ---
    # 해시 분석 가이드라인 추가
    hash_prompt_path = os.path.join(script_dir, "prompt", "해시분석.md")
    if os.path.exists(hash_prompt_path):
        with open(hash_prompt_path, "r", encoding="utf-8") as f:
            custom_prompt_content += "\n\n[해시분석.md 내용]\n" + f.read()

    # 확장자별 분석 가이드라인 추가
    _, ext = os.path.splitext(target_abs_path)
    ext = ext.lower()
    ext_prompt_content = ""
    
    if ext == '.pdf':
        ext_prompt_path = os.path.join(script_dir, "prompt", "pdf분석.md")
    elif ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.jfif']:
        ext_prompt_path = os.path.join(script_dir, "prompt", "이미지파일분석.md")
    elif ext in ['.xls', '.xlsx', '.xlsm']:
        ext_prompt_path = os.path.join(script_dir, "prompt", "xlsx파일분석.md")
    else:
        ext_prompt_path = None
        
    if ext_prompt_path and os.path.exists(ext_prompt_path):
        with open(ext_prompt_path, "r", encoding="utf-8") as f:
             custom_prompt_content += f"\n\n[{os.path.basename(ext_prompt_path)} 내용]\n" + f.read()
    # -----------------------------------
        
    prompt = f"""
[CRITICAL INSTRUCTION: STRICT TEMPLATE ENFORCEMENT]
You are an expert malware analyst and an automated reporting bot.
You MUST output your ENTIRE response in KOREAN (한국어). Do NOT use English except for specific technical terms.

YOUR ONLY PURPOSE IS TO FILL IN THE [REQUIRED TEMPLATE] BELOW.
1. You MUST copy the EXACT markdown headers (`#`, `##`) from the template.
2. DO NOT change the names or numbering of the headers in the template.
3. DO NOT write a conversational summary or a bulleted list instead of the template.
4. If a file is benign or small, STILL USE THE FULL TEMPLATE and simply write "특이사항 없음" or "정상 파일" in the respective sections.
5. DO NOT output anything outside of the template structure.

[REQUIRED TEMPLATE]
{custom_prompt_content}

[EXAMPLE OF EXPECTED OUTPUT REASONING]
If the log shows it's a PNG image, your output should look exactly like this:
---
# IMG_1234.png 파일 분석 결과

## 1. 분석 개요
- 분석 파일 : IMG_1234.png
- 분석 도구 : Hash, ExifTool, Strings
- 분석 일시 : 2026-02-26 14:00:00

## 2. 해시분석 결과
- SHA256: d0595d853... (바이러스토탈 제출 기록 없음)

...and so on for the rest of the template headers exactly as requested.

[ANALYSIS LOG DATA]
"""
    
    python_exe = sys.executable
    # file_analysis.py가 이 스크립트와 같은 디렉토리에 있다고 가정하거나,
    # 이전의 하드코딩된 경로를 사용합니다. 
    # 스크립트 위치를 기반으로 경로를 찾는 것이 더 안정적입니다.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_analysis_script = os.path.join(script_dir, "file_analysis.py")
    if not os.path.exists(file_analysis_script):
        print(f"[!] file_analysis.py not found at: {file_analysis_script}")
        return

    print(f"[*] [START] Analyzing: {filename}")
    
    # API 키가 모두 소진된 경우 즉시 스킵
    if api_keys_exhausted or key_manager.keys_exhausted:
        print(f"  [-] Skipped (API keys exhausted): {filename}")
        return
    
    # 중복 분석 방지: 이미 분석된 결과 보고서가 존재하는지 날짜 앞자리와 무관하게 패턴(glob)으로 확인
    base_name_no_ext = os.path.splitext(filename)[0].strip()
    target_dir = os.path.dirname(target_abs_path)
    
    # glob를 이용해 과거(어제 등)에 생성된 파일들도 인식하도록 와일드카드 사용
    safe_base_name = glob.escape(base_name_no_ext)
    safe_target_dir = glob.escape(target_dir)
    md_pattern = os.path.join(safe_target_dir, f"*_{safe_base_name}_analysis.md")
    report_pattern = os.path.join(safe_target_dir, f"*_{safe_base_name}_ai_analysis_report.md")
    
    # AI 보고서가 존재할 때만 스킵 (이전에 파일 분석만 되고 AI 분석이 실패/누락된 경우 재시도하기 위해)
    if glob.glob(report_pattern):
        print(f"  [-] Skipped (Already analyzed): {filename}")
        return

    # ── 폴더 내 분석 대상 파일 수 체크 ──────────────────────────────────
    LARGE_ATTACHMENT_THRESHOLD = 3
    SUPPORTED_EXTS = {'.pdf', '.xls', '.xlsx', '.xlsm', '.doc', '.docx', '.docm',
                      '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.jfif',
                      '.exe', '.dll', '.sys', '.ocx',
                      '.ppt', '.pptx', '.pptm', '.potx', '.pps', '.ppsx'}

    sibling_files = [
        f for f in os.listdir(target_dir)
        if os.path.isfile(os.path.join(target_dir, f))
        and os.path.splitext(f)[1].lower() in SUPPORTED_EXTS
    ]
    is_large_folder = len(sibling_files) >= LARGE_ATTACHMENT_THRESHOLD

    # EML 없이 직접 넣은 경우에도 폴더명 앞에 [첨부파일 3개 이상] 붙이기
    if is_large_folder:
        folder_basename = os.path.basename(target_dir)
        if not folder_basename.startswith("[첨부파일 3개 이상]"):
            parent_dir = os.path.dirname(target_dir)
            new_folder_name = f"[첨부파일 3개 이상]{folder_basename}"
            new_target_dir = os.path.join(parent_dir, new_folder_name)
            try:
                if not os.path.exists(new_target_dir):
                    os.rename(target_dir, new_target_dir)
                    print(f"  [*] 폴더명 변경: {folder_basename} → {new_folder_name}")
                    # target_abs_path도 새 경로로 갱신
                    target_abs_path = os.path.join(new_target_dir, filename)
                    target_dir = new_target_dir
            except Exception as e:
                print(f"  [!] 폴더 이름 변경 실패: {e}")
    # ────────────────────────────────────────────────────────────────────

    found_mds = glob.glob(md_pattern)
    expected_md_path = ""
    analysis_stdout = ""

    if found_mds:
        # 이미 로컬 분석 결과가 있다면 재분석(file_analysis.py)하지 않고 재사용하여 시간을 크게 절약합니다.
        expected_md_path = found_mds[-1]
        try:
            with open(expected_md_path, 'r', encoding='utf-8', errors='replace') as f:
                analysis_stdout = f.read()
        except Exception as e:
            print(f"[!] [{filename}] Failed to read generated analysis report: {e}")
            return

        # 파일 수 3개 이상이면 AI 리포트 없이 종료
        if is_large_folder:
            print(f"  [-] Skipped AI report (Large folder: {len(sibling_files)} files in {os.path.basename(target_dir)})")
            return
    else:
        # 1. file_analysis.py 모듈 직접 호출
        target_dir = os.path.dirname(target_abs_path)
        sys.path.insert(0, script_dir)
        try:
            import file_analysis
            # 분석 결과 dict 받기
            result_dict = file_analysis.analyze_file_as_dict(target_abs_path)
            
            if result_dict.get("status") == "error":
                print(f"[!] [{filename}] file_analysis.py module failed: {result_dict.get('message')}")
                return
            
            analysis_stdout = result_dict.get("raw_analysis_log", "")
            
            if not analysis_stdout:
                print(f"[!] [{filename}] file_analysis.py did not generate expected report output.")
                return
                
            # 로그를 .md 파일로도 백업 저장 (기존 동작 유지)
            date_str_local = datetime.datetime.now().strftime("%y%m%d")
            log_filename = f"{date_str_local}_{base_name_no_ext}_analysis.md"
            log_full_path = os.path.join(target_dir, log_filename)
            try:
                with open(log_full_path, "w", encoding='utf-8') as lf:
                    lf.write(analysis_stdout)
            except Exception as e:
                pass

            # 파일 수 3개 이상이면 AI 리포트 없이 종료 (file_analysis만 실행됨)
            if is_large_folder:
                print(f"  [-] Skipped AI report (Large folder: {len(sibling_files)} files in {os.path.basename(target_dir)})")
                return
                
        except Exception as e:
            print(f"[!] [{filename}] Failed to run file_analysis module: {e}")
            return
        finally:
            if script_dir in sys.path:
                sys.path.remove(script_dir)
        
    # PDF 등 텍스트가 너무 길어서 Gemini API 터지는 문제 방지 (글자 수 강제 제한)
    MAX_CHAR_LIMIT = 10000
    if len(analysis_stdout) > MAX_CHAR_LIMIT:
        print(f"[*] [{filename}] 텍스트가 너무 깁니다 ({len(analysis_stdout)}자). 요약 서버 전송을 위해 자릅니다.")
        half_limit = MAX_CHAR_LIMIT // 2
        analysis_stdout = analysis_stdout[:half_limit] + "\n\n... [내용이 너무 길어 중략됨] ...\n\n" + analysis_stdout[-half_limit:]

    # 2. gemini 실행 (Windows에서는 .cmd 래퍼 사용)
    # .cmd를 우선적으로 검색 (PowerShell 정책 우회)
    gemini_cmd = shutil.which("gemini.cmd")
    if not gemini_cmd:
        gemini_cmd = shutil.which("gemini")
        
    if not gemini_cmd:
        print(f"[!] [{filename}] gemini command not found in PATH.")
        return

    cmd_ai = [gemini_cmd, "--model", "gemini-2.5-flash"]
    # 프롬프트를 인자로 전달하면 Windows CMD에서 멀티라인(줄바꿈) 포맷이 파괴되어 
    # 마크다운 템플릿 구조가 무너지는 치명적 버그가 있습니다. 
    # 따라서 안내 프롬프트와 분석 데이터를 통째로 합쳐서 STDIN으로 밀어넣습니다.
    combined_input = prompt + "\n\n" + analysis_stdout
    ret_code_ai, ai_stdout, ai_stderr = await run_command_async(cmd_ai, input_data=combined_input)

    if ret_code_ai != 0:
        print(f"[!] [{filename}] gemini-cli failed: {ai_stderr}")
        log_api_request(filename, f"FAILED ({ai_stderr.replace(chr(10), ' ')[:50]}...)")
        final_result = f"[Gemini API 분석 실패 - {ai_stderr.strip()}]\n\n"
    else:
        log_api_request(filename, "SUCCESS")
        # 영어로 나온 결과물에 대해 오프라인 번역 모듈 자동 가동
        translated_stdout = translate_if_english(ai_stdout)
        final_result = translated_stdout + "\n\n"

    # 원본 파일 정보 제거 요청에 따라 분석원문 추가 로직 삭제
    
    # 3. 통합 보고서 저장 (대상 파일과 같은 폴더에 저장)
    date_str = datetime.datetime.now().strftime("%y%m%d")
    output_filename = f"{date_str}_{base_name_no_ext}_ai_analysis_report.md"
    report_path = os.path.join(target_dir, output_filename)
    
    try:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(final_result)
        print(f"[+] [{filename}] AI Report saved to: {report_path}")
    except IOError as e:
        print(f"[!] [{filename}] Failed to write AI report: {e}")


async def analyze_urls_async(urls_file, output_dir_base):
    """
    urls.txt 파일을 기반으로 VT URL 평판 조회 + AI 리포트를 생성합니다.
    """
    urls_abs_path = os.path.abspath(urls_file)
    target_dir = os.path.dirname(urls_abs_path)
    folder_name = os.path.basename(target_dir)
    
    if not os.path.exists(urls_abs_path):
        print(f"[!] URL file not found: {urls_file}")
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_analysis_script = os.path.join(script_dir, "file_analysis.py")
    python_exe = sys.executable

    print(f"[*] [START] URL Analysis: {folder_name}")
    
    # API 키가 모두 소진된 경우 즉시 스킵
    if api_keys_exhausted or key_manager.keys_exhausted:
        print(f"  [-] Skipped (API keys exhausted): {folder_name}")
        return

    # 중복 분석 방지
    date_str = datetime.datetime.now().strftime("%y%m%d")
    report_pattern = os.path.join(glob.escape(target_dir), "*_url_ai_analysis_report.md")
    if glob.glob(report_pattern):
        print(f"  [-] Skipped (URL already analyzed): {folder_name}")
        return

    # 1. file_analysis.py -urls 모듈 직접 호출
    md_pattern = os.path.join(glob.escape(target_dir), "*_url_analysis.md")
    found_mds = glob.glob(md_pattern)
    analysis_stdout = ""
    
    if found_mds:
        expected_md_path = found_mds[-1]
        try:
            with open(expected_md_path, 'r', encoding='utf-8', errors='replace') as f:
                analysis_stdout = f.read()
        except Exception as e:
            print(f"[!] [{folder_name}] Failed to read URL analysis report: {e}")
            return
    else:
        try:
            # -urls 플래그를 사용하여 독립 서브프로세스로 안전하게 실행
            cmd = [python_exe, file_analysis_script, "-urls", urls_abs_path]
            ret_code, stdout, stderr = await run_command_async(cmd)
            
            if ret_code != 0:
                print(f"[!] [{folder_name}] file_analysis.py sub-process failed: {stderr}")
                return
                
            analysis_stdout = stdout
            
            if not analysis_stdout:
                print(f"[!] [{folder_name}] file_analysis.py did not generate URL analysis report.")
                return
                
            # 로그를 .md 파일로 예약 백업
            date_str_local = datetime.datetime.now().strftime("%y%m%d")
            log_filename = f"{date_str_local}_url_analysis.md"
            log_full_path = os.path.join(target_dir, log_filename)
            try:
                with open(log_full_path, "w", encoding='utf-8') as lf:
                    lf.write(analysis_stdout)
            except Exception as e:
                pass
                
        except Exception as e:
            print(f"[!] [{folder_name}] Failed to run file_analysis sub-process for URLs: {e}")
            return

    # URL 분석용 프롬프트 로드
    url_prompt_path = os.path.join(script_dir, "prompt", "url분석.md")
    custom_prompt_content = ""
    if os.path.exists(url_prompt_path):
        with open(url_prompt_path, "r", encoding="utf-8") as f:
            custom_prompt_content = f.read()
    else:
        custom_prompt_content = "다음 URL 분석 결과를 기반으로 보고서를 작성해 주세요."

    prompt = f"""
[CRITICAL INSTRUCTION: STRICT TEMPLATE ENFORCEMENT]
You are an expert malware analyst and an automated reporting bot.
You MUST output your ENTIRE response in KOREAN (한국어). Do NOT use English except for specific technical terms.

YOUR ONLY PURPOSE IS TO FILL IN THE [REQUIRED TEMPLATE] BELOW.
1. You MUST copy the EXACT markdown headers (`#`, `##`) from the template.
2. DO NOT change the names or numbering of the headers in the template.
3. If all URLs are benign, STILL USE THE FULL TEMPLATE and write "특이사항 없음" or "정상 URL" in the respective sections.
4. DO NOT output anything outside of the template structure.

[REQUIRED TEMPLATE]
{custom_prompt_content}

[URL ANALYSIS LOG DATA]
"""

    combined_input = prompt + "\n\n" + analysis_stdout
    
    gemini_cmd = shutil.which("gemini.cmd")
    if not gemini_cmd:
        gemini_cmd = shutil.which("gemini")
    if not gemini_cmd:
        print(f"[!] [{folder_name}] gemini command not found in PATH.")
        return

    cmd_ai = [gemini_cmd, "--model", "gemini-2.5-flash"]
    ret_code_ai, ai_stdout, ai_stderr = await run_command_async(cmd_ai, input_data=combined_input)

    if ret_code_ai != 0:
        print(f"[!] [{folder_name}] gemini-cli failed: {ai_stderr}")
        log_api_request(f"URL:{folder_name}", f"FAILED ({ai_stderr.replace(chr(10), ' ')[:50]}...)")
        final_result = f"[Gemini API URL 분석 실패 - {ai_stderr.strip()}]\n\n"
    else:
        log_api_request(f"URL:{folder_name}", "SUCCESS")
        translated_stdout = translate_if_english(ai_stdout)
        final_result = translated_stdout + "\n\n"

    # AI 리포트 저장
    output_filename = f"{date_str}_{folder_name}_url_ai_analysis_report.md"
    # 파일명 길이 제한
    if len(output_filename) > 80:
        output_filename = output_filename[:76] + ".md"
    report_path = os.path.join(target_dir, output_filename)
    
    try:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(final_result)
        print(f"[+] [{folder_name}] URL AI Report saved to: {report_path}")
    except IOError as e:
        print(f"[!] [{folder_name}] Failed to write URL AI report: {e}")



async def main_async():
    parser = argparse.ArgumentParser(description="AI Analysis Wrapper (Batch & Async)")
    parser.add_argument("-file", dest="filename", help="Target file for analysis (Optional). If omitted, scans directory.")
    parser.add_argument("-dir", dest="target_dir", help="Target directory to scan for batch analysis (Optional). Defaults to current directory.")
    parser.add_argument("-out", dest="output_dir", default=output_path, help="Directory to save analysis result")
    
    args = parser.parse_args()
    
    tasks = []
    url_tasks = []
    
    if args.filename:
        # 단일 파일 모드
        tasks.append(analyze_file_async(args.filename, args.output_dir))
    else:
        # 일괄 처리 모드 (지정된 디렉토리 또는 기본 attachfiles 디렉토리)
        if args.target_dir:
            scan_dir = os.path.abspath(args.target_dir)
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            scan_dir = os.path.join(script_dir, "attachfiles")
            
            # Step 0: EML 첨부파일 자동 추출
            eml_dir = os.path.join(script_dir, "eml")
            if os.path.isdir(eml_dir):
                eml_files = [f for f in os.listdir(eml_dir) if f.lower().endswith('.eml')]
                if eml_files:
                    print(f"[*] Found {len(eml_files)} EML file(s). Extracting attachments...")
                    try:
                        from extract_attachments import extract_attachments
                        total = 0
                        for eml_file in sorted(eml_files):
                            eml_path = os.path.join(eml_dir, eml_file)
                            count, folder, _, _, _ = extract_attachments(eml_path, scan_dir)
                            total += count
                        if total > 0:
                            print(f"[*] Extracted {total} attachment(s) from EML files.\n")
                        else:
                            print(f"[*] No new attachments found in EML files.\n")
                    except Exception as e:
                        print(f"[!] EML extraction failed: {e}\n")
        
        if not os.path.isdir(scan_dir):
            print(f"[!] Directory not found: {scan_dir}")
            return
        
        print(f"[*] Scanning directory (recursive): {scan_dir}")
        
        ignored_files = ["ai_analysis.py", "file_analysis.py", "extract_attachments.py", ".DS_Store"]
        ignored_extensions = [".md", ".py", ".pyc", ".txt", ".log", ".ini", ".eml"]
        ignored_dirs = {"analyzed", "ai_analysis_report", "instruction_output", "__pycache__"}
        
        for root, dirs, files in os.walk(scan_dir):
            # 제외 디렉토리는 탐색하지 않음
            dirs[:] = [d for d in dirs if d not in ignored_dirs]
            
            for f in files:
                full_path = os.path.join(root, f)
                
                # 스크립트 자체나 이미 알려진 제외 파일들은 건너뜀
                if f in ignored_files or f.startswith('.'):
                    continue
                    
                # 확장자를 기준으로 결과 파일이나 스크립트는 건너뜀
                _, ext = os.path.splitext(f)
                if ext.lower() in ignored_extensions:
                    continue

                tasks.append(analyze_file_async(full_path, args.output_dir))
        
        # urls.txt 파일도 스캔하여 URL 분석 태스크 추가
        for root, dirs, files in os.walk(scan_dir):
            dirs[:] = [d for d in dirs if d not in ignored_dirs]
            for f in files:
                if f == "urls.txt":
                    full_path = os.path.join(root, f)
                    url_tasks.append(analyze_urls_async(full_path, args.output_dir))

    if not tasks and not url_tasks:
        print("[*] No files or URLs to analyze.")
        return
        


    print(f"[*] Starting batch analysis for {len(tasks)} file(s) + {len(url_tasks)} URL set(s)...")
    all_tasks = tasks + url_tasks
    total_items = len(all_tasks)
    
    # 동적 대기 시간 계산 (사용 가능한 키 수 기반)
    initial_delay = key_manager.get_optimal_delay()
    ok_keys = key_manager.get_ok_key_count()
    expected_time_sec = int(total_items * initial_delay)
    minutes = expected_time_sec // 60
    seconds = expected_time_sec % 60
    time_str = f"{minutes}분 {seconds}초" if minutes > 0 else f"{seconds}초"
    
    print("\n" + "!" * 60)
    print(f" [분석 요약] 총 {total_items}건 (파일 {len(tasks)}개, URL {len(url_tasks)}개)")
    print(f" [사용 가능 키] {ok_keys}개 (동적 대기: {initial_delay:.1f}초/건)")
    print(f" [예상 소요 시간] 약 {time_str}")
    print("!" * 60 + "\n")
    
    print(f"\n[*] Found {total_items} item(s). Smart rotation active ({ok_keys} keys)...\n")
    
    for i, task in enumerate(all_tasks):
        await task
        
        # API 키가 모두 소진된 경우 나머지 파일 처리를 중단
        if api_keys_exhausted or key_manager.keys_exhausted:
            remaining = total_items - i - 1
            if remaining > 0:
                print(f"\n[!] API 키 한도가 모두 소진되어 나머지 {remaining}개 항목 분석을 중단합니다.")
                print("[!] API 한도가 초기화된 후 다시 실행하면 미완료 항목부터 이어서 분석합니다.")
            break
        
        if i < total_items - 1:
            # 매 요청마다 사용 가능한 키 수에 따라 동적 대기 시간 재계산
            dynamic_delay = key_manager.get_optimal_delay()
            print(f"[*] Waiting {dynamic_delay:.1f}s (OK keys: {key_manager.get_ok_key_count()})...")
            await asyncio.sleep(dynamic_delay)
            
    print("\n[*] All analysis tasks completed.")

def main():
    try:
        if sys.platform == 'win32':
             # 윈도우 환경에서 비동기 서브프로세스 실행을 위한 이벤트 루프 정책 설정
             # 파이썬 3.8+ 윈도우에서는 ProactorEventLoop가 기본이지만 명시적으로 설정
             asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        asyncio.run(main_async())
    except KeyboardInterrupt:
        print("\n[!] Analysis interrupted by user.")
    except Exception as e:
        print(f"[!] Unexpected error: {e}")

if __name__ == "__main__":
    main()