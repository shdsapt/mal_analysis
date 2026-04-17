#!/usr/bin/env python3
"""
Automated File Analysis Tool (Windows Compatible)
PDF, Excel, Image, EXE, PPT 파일의 정적 분석을 수행합니다.
Linux 셸 명령어 의존성을 제거하고 Python 내장 라이브러리로 교체하였습니다.
"""

import argparse
import subprocess
import json
import os
import sys
import re
import hashlib
import datetime
import io
import base64
import time

# 중첩된 코루틴 루프 처리 (VirusTotal `already running` 에러 방지)
try:
    import nest_asyncio
    nest_asyncio.apply()
except ImportError:
    pass

# 윈도우 한글(cp949) 출력 에러 방지 (sys.stdout, sys.stderr 덮어쓰기)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 이번 실행 세션 동안 분석된 URL을 기억하여 중복 조회를 방지합니다.
GLOBAL_ANALYZED_URLS = set()

try:
    import pefile
except ImportError:
    pefile = None

try:
    import magic
except ImportError:
    magic = None

try:
    import pdfid.pdfid as pdfid_module
except ImportError:
    pdfid_module = None

try:
    import pikepdf
except ImportError:
    pikepdf = None

try:
    import vt
except ImportError:
    vt = None

# 실행 결과가 저장될 디렉토리 경로 (기본값: 대상 파일과 같은 폴더)

# VirusTotal API 키 (환경변수 우선, 없으면 기본값 사용)
VT_API_KEY = os.environ.get("VT_API_KEY", "e096491a2920f219e960abe5732c27dd3e0b7bcd7aaba9de262ebd70f4767f96")

# 일일 VT API 조회 최대 횟수 (VT 무료 API 500건 중 450건으로 제한)
MAX_DAILY_VT_LOOKUPS = 450
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 헬퍼 함수 (Linux 셸 명령어 대체)
# ============================================================

def extract_strings(filepath, min_length=4, chunk_size=1*1024*1024, timeout=30):
    """
    파일에서 printable ASCII 문자열을 추출합니다. (Linux 'strings' 명령어 대체)
    min_length 이상의 연속된 printable 문자열을 반환합니다.
    
    대용량 파일의 행(Hang) 방지를 위해:
    - chunk_size (기본 1MB) 단위로 잘라서 순차 스캔
    - timeout (기본 30초) 초과 시 스캔 중단 후 그때까지의 결과 반환
    """
    import time
    pattern = re.compile(rb'[\x20-\x7E]{%d,}' % min_length)
    results = []
    start_time = time.time()
    timed_out = False
    try:
        file_size = os.path.getsize(filepath)
        with open(filepath, 'rb') as f:
            while True:
                # 타임아웃 체크
                if time.time() - start_time > timeout:
                    timed_out = True
                    break
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                results.extend(
                    s.decode('ascii', errors='ignore') for s in pattern.findall(chunk)
                )
        if timed_out:
            elapsed = time.time() - start_time
            print(f"[!] extract_strings: {timeout}초 타임아웃 도달 (파일 크기: {file_size:,} bytes). 스캔된 부분까지의 결과를 사용합니다.")
    except Exception as e:
        print(f"[!] extract_strings failed: {e}")
    return results


def grep_patterns(strings_list, pattern_str, case_insensitive=True):
    """
    문자열 리스트에서 정규식 패턴에 매칭되는 항목을 필터링합니다. (Linux 'grep -E' 대체)
    """
    flags = re.IGNORECASE if case_insensitive else 0
    try:
        compiled = re.compile(pattern_str, flags)
        return [s for s in strings_list if compiled.search(s)]
    except re.error as e:
        print(f"[!] grep_patterns regex error: {e}")
        return []


def calculate_sha256(filepath):
    """
    파일의 SHA256 해시를 계산합니다. (Linux 'sha256sum' 대체)
    """
    sha256 = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()
    except Exception as e:
        print(f"[!] SHA256 calculation failed: {e}")
        return None


def get_file_type(filepath):
    """
    파일의 MIME 타입을 확인합니다. (Linux 'file' 명령어 대체)
    python-magic 라이브러리를 사용하며, 없을 경우 확장자 기반 판별로 fallback합니다.
    """
    if magic:
        try:
            # magic.from_file가 한글 경로에서 Illegal byte sequence 오류를 내는 경우 우회 (버퍼를 직접 읽음)
            with open(filepath, 'rb') as f:
                header_data = f.read(2048)
            result = magic.from_buffer(header_data)
            return result
        except Exception as e:
            # libmagic C 라이브러리가 한글 윈도우 사용자명 경로 파싱 중 에러를 발생시키는 경우가 잦으므로,
            # 에러 메시지를 숨기고 조용히 확장자 기반 fallback(아래 코드)으로 넘깁니다.
            pass

    # Fallback: 확장자 기반 판별
    ext = os.path.splitext(filepath)[1].lower()
    type_map = {
        '.pdf': 'PDF document',
        '.xls': 'Microsoft Excel (Legacy Binary)',
        '.xlsx': 'Microsoft Excel (OpenXML / ZIP archive)',
        '.xlsm': 'Microsoft Excel Macro-Enabled (OpenXML / ZIP archive)',
        '.png': 'PNG image',
        '.jpg': 'JPEG image',
        '.jpeg': 'JPEG image',
        '.gif': 'GIF image',
        '.bmp': 'BMP image',
        '.exe': 'PE32 executable',
        '.dll': 'PE32 dynamic-link library',
        '.ppt': 'Microsoft PowerPoint (Legacy Binary)',
        '.pptx': 'Microsoft PowerPoint (OpenXML / ZIP archive)',
    }
    return type_map.get(ext, f'Unknown file type (extension: {ext})')


def run_external_command(command_list, description, timeout=45):
    """
    외부 명령어(exiftool, vt, olevba 등)를 실행합니다.
    shell=False로 실행하여 보안성을 높입니다.
    기본 타임아웃은 45초로 설정하여 대용량 파일 분석 시 무한 대기를 방지합니다.
    """
    print(f"\n{'='*60}")
    print(f"[+] Running: {description}")
    print(f"    Command: {' '.join(command_list)}")
    print(f"{'='*60}\n")

    try:
        import tempfile
        import os
        
        # OS 파이프 버퍼 데드락을 방지하면서 출력을 문자열로 캡처하기 위해 임시 파일을 사용합니다.
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding='utf-8') as temp_out:
            temp_out_path = temp_out.name

        with open(temp_out_path, 'w+', encoding='utf-8', errors='replace') as f_out:
            subprocess.run(
                command_list,
                stdout=f_out,
                stderr=subprocess.STDOUT,
                timeout=timeout,
                text=True
            )
            
        with open(temp_out_path, 'r', encoding='utf-8', errors='replace') as f_in:
            output_text = f_in.read()
            
        try:
            os.remove(temp_out_path)
        except Exception:
            pass

        if output_text:
            print(output_text)

        return output_text
    except FileNotFoundError:
        print(f"[!] Command not found: {command_list[0]}")
        print(f"    Ensure '{command_list[0]}' is installed and in your PATH.")
        return ""
    except subprocess.TimeoutExpired:
        print(f"[!] Command timed out after 120 seconds.")
        return ""
    except Exception as e:
        print(f"[!] Execution Failed: {e}")
        return ""


# ============================================================
# Logger 클래스
# ============================================================

class Logger(object):
    def __init__(self, filename):
        self.terminal = sys.stdout
        # 로그 파일 저장 시 cp949 충돌을 막기 위해 utf-8 지정 및 에러 replace
        self.log = open(filename, "a", encoding='utf-8', errors='replace')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()


# ============================================================
# 분석 함수
# ============================================================

def analyze_hash(target_file):
    """
    파일의 해시를 계산하고 VirusTotal API로 조회합니다.
    vt-py 라이브러리를 사용합니다.
    """
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 해시 및 바이러스토탈 분석")
    print(f"{'='*60}\n")

    # Step 1: SHA256 해시 계산 (Python 내장)
    file_hash = calculate_sha256(target_file)
    if not file_hash:
        print("[!] Hash calculation failed.")
        return

    print(f"[*] SHA256 해시값: {file_hash}")

    # 일일 한도 검사 추가
    daily_count = _get_daily_vt_count()
    if daily_count >= MAX_DAILY_VT_LOOKUPS:
        print(f"[!] 일일 VT API 조회 한도({MAX_DAILY_VT_LOOKUPS}건) 도달. 해시 조회를 생략합니다.")
        return

    # Step 2: VirusTotal API 조회 (vt-py)
    if not vt:
        print("[*] vt-py module not installed. Skipping VirusTotal lookup.")
        print("    Install with: pip install vt-py")
        return

    if not VT_API_KEY:
        print("[*] VT_API_KEY not set. Skipping VirusTotal lookup.")
        return

    try:
        client = vt.Client(VT_API_KEY)
        try:
            file_info = client.get_object(f"/files/{file_hash}")
            _increment_daily_vt_count()

            # 기본 정보
            print(f"\n[+] 바이러스토탈 검사 결과:")
            stats = file_info.last_analysis_stats
            total = sum(stats.values())
            malicious = stats.get('malicious', 0)
            suspicious = stats.get('suspicious', 0)
            undetected = stats.get('undetected', 0)

            print(f"    Detection: {malicious}/{total} 엔진에서 악성으로 탐지됨")
            print(f"    의심: {suspicious}, 미탐지: {undetected}")

            # 파일 정보
            if hasattr(file_info, 'meaningful_name'):
                print(f"    파일명: {file_info.meaningful_name}")
            if hasattr(file_info, 'type_description'):
                print(f"    파일 확장자: {file_info.type_description}")
            if hasattr(file_info, 'size'):
                print(f"    파일 크기: {file_info.size:,} bytes")
            if hasattr(file_info, 'first_submission_date'):
                print(f"    최초 발견: {file_info.first_submission_date}")
            if hasattr(file_info, 'last_analysis_date'):
                print(f"    최근 분석: {file_info.last_analysis_date}")

            # 악성 탐지 엔진 목록
            if malicious > 0:
                print(f"\n[!] 악성 탐지 내역:")
                results = file_info.last_analysis_results
                for engine, result in results.items():
                    if result['category'] == 'malicious':
                        print(f"    [{engine}] {result.get('result', 'N/A')}")

            # 태그 정보
            if hasattr(file_info, 'tags') and file_info.tags:
                print(f"\n[*] 태그: {', '.join(file_info.tags)}")

        except vt.error.APIError as e:
            if 'NotFoundError' in str(e):
                print(f"[*] 바이러스토탈 데이터베이스에서 해시를 찾을 수 없습니다.")
                print(f"    이 파일은 바이러스토탈에 제출된 적이 없습니다.")
            else:
                print(f"[!] VirusTotal API Error: {e}")
        finally:
            client.close()

    except Exception as e:
        print(f"[!] VirusTotal query failed: {e}")


def analyze_pdf(target_file):
    """
    PDF 파일 분석을 수행합니다.
    pdfid(Python API), pikepdf, Python strings 추출을 사용합니다.
    """
    # Step 1: pdfid 분석
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: PDFID 분석")
    print(f"{'='*60}\n")

    if pdfid_module:
        try:
            # pdfid를 직접 실행 (CLI 래퍼 사용). -l 옵션은 와일드카드 매칭을 무시하게 하여 [] 폴더명 에러를 방지합니다.
            run_external_command(
                [sys.executable, "-m", "pdfid.pdfid", "-l", target_file],
                "PDFID Analysis (Python module)"
            )
        except Exception as e:
            print(f"[!] pdfid failed: {e}")
    else:
        print("[!] pdfid module not installed. Install with: pip install pdfid")

    # Step 2: pikepdf 분석 (peepdf 대체)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: PDF 구조 분석 (pikepdf)")
    print(f"{'='*60}\n")

    if pikepdf:
        try:
            pdf = pikepdf.Pdf.open(target_file)
            print(f"[*] PDF 버전: {pdf.pdf_version}")
            print(f"[*] 페이지 수: {len(pdf.pages)}")
            print(f"[*] 암호화 여부: {pdf.is_encrypted}")

            # 메타데이터 출력 (인코딩 안전 처리)
            if pdf.docinfo:
                print(f"\n[*] 문서 정보:")
                for key, value in pdf.docinfo.items():
                    try:
                        val_str = str(value)
                        # cp949에서 인코딩할 수 없는 문자를 안전하게 처리
                        val_str = val_str.encode('utf-8', errors='replace').decode('utf-8', errors='replace')
                        print(f"    {key}: {val_str}")
                    except Exception:
                        print(f"    {key}: (encoding error - unable to display)")

            # PDF 객체 트리 요약
            print(f"\n[*] PDF 객체 요약:")
            print(f"    총 객체 수: {len(pdf.objects)}")

            # JavaScript 또는 의심스러운 액션 검색
            suspicious_keys = ['/JS', '/JavaScript', '/OpenAction', '/AA',
                             '/Launch', '/EmbeddedFile', '/URI', '/SubmitForm']
            found_suspicious = []

            import time
            scan_start_time = time.time()
            scan_timeout = 30
            timed_out = False

            for obj_id in pdf.objects:
                if time.time() - scan_start_time > scan_timeout:
                    timed_out = True
                    break
                try:
                    obj = pdf.objects[obj_id]
                    if isinstance(obj, pikepdf.Dictionary):
                        for key in suspicious_keys:
                            if key in obj:
                                found_suspicious.append(
                                    f"Object {obj_id}: {key} found"
                                )
                except Exception:
                    continue

            if timed_out:
                print(f"\n[!] pikepdf 객체 스캔 시간이 {scan_timeout}초 타임아웃을 초과했습니다. 일부 스캔 건너뜀.")

            if found_suspicious:
                print(f"\n[!] 의심스러운 PDF 요소 발견:")
                for item in found_suspicious:
                    print(f"    {item}")
            else:
                print(f"\n[*] 감시 목록에서 의심스러운 PDF 요소가 발견되지 않았습니다.")

            pdf.close()
        except pikepdf.PasswordError:
            print("[!] PDF is password-protected. Cannot fully analyze without password.")
            print("[*] Basic info (from pdfid above) is still available.")
        except Exception as e:
            print(f"[!] pikepdf analysis failed: {e}")
    else:
        print("[!] pikepdf module not installed. Install with: pip install pikepdf")

    # Step 3: 의심스러운 문자열 추출 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 의심스러운 문자열 추출 (PDF)")
    print(f"{'='*60}\n")

    all_strings = extract_strings(target_file)
    pattern = r'https?://|www\.|\.exe|\.js|\.zip|\.bat|\.ps1|powershell|cmd\.exe'
    matches = grep_patterns(all_strings, pattern)

    if matches:
        print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
        for m in matches:
            print(f"    {m}")
    else:
        print("[*] 의심스러운 문자열이 발견되지 않았습니다.")


def analyze_xls(target_file, deep_analysis=False):
    """
    엑셀(XLS, XLSX) 파일 분석을 수행합니다.
    """
    import zipfile
    import shutil

    # Step 1: 파일 타입 확인 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 파일 구조 및 타입 확인")
    print(f"{'='*60}\n")

    file_type_str = get_file_type(target_file)
    print(f"[*] 파일 확장자: {file_type_str}")

    # Step 2: olevba 실행 (매크로 분석) - Windows에서도 동작
    run_external_command(
        ["olevba", target_file],
        "OLEVBA Macro Analysis"
    )

    # Step 3: Deep Analysis (ZIP 기반 OpenXML)
    if zipfile.is_zipfile(target_file):
        if deep_analysis:
            print(f"\n{'='*60}")
            print(f"[+] 실행 중: 정밀 분석 (압축 해제 및 내부 검사)")
            print(f"{'='*60}\n")

            temp_dir = f"temp_unzip_{os.path.basename(target_file)}"
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            try:
                print(f"[*] '{target_file}'을(를) '{temp_dir}'에 압축 해제 중...")
                with zipfile.ZipFile(target_file, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)

                # 3-1. 매크로(Macro) 존재 여부 확인
                vba_bin_path = os.path.join(temp_dir, "xl", "vbaProject.bin")
                if os.path.exists(vba_bin_path):
                    print("\n[!] 'xl/vbaProject.bin' 발견됨: 이 파일에는 내부 VBA 매크로가 포함되어 있습니다.")
                else:
                    print("\n[*] 'xl/vbaProject.bin' 미발견: 표준 내부 매크로 바이너리가 없습니다.")

                # 3-2. 위험한 키워드 검색 (Python 구현 - grep 대체)
                print("\n[*] 압축 해제된 내부 파일들에서 의심 키워드 검색 중:")
                grep_pattern = r'cmd|powershell|\.exe|http://|https://|external|\.bat|\.vbs|\.ps1'
                total_found = 0

                for root, dirs, files in os.walk(temp_dir):
                    for fname in files:
                        fpath = os.path.join(root, fname)
                        try:
                            with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                                for line_num, line in enumerate(f, 1):
                                    if re.search(grep_pattern, line, re.IGNORECASE):
                                        rel_path = os.path.relpath(fpath, temp_dir)
                                        print(f"    {rel_path}:{line_num}: {line.strip()[:200]}")
                                        total_found += 1
                        except Exception:
                            continue

                if total_found == 0:
                    print("    내부 파일들에서 의심스러운 추출 키워드가 없습니다.")
                else:
                    print(f"\n    [*] 총 일치 항목: {total_found}")

            except Exception as e:
                print(f"[!] Deep Analysis Failed: {e}")
            finally:
                if os.path.exists(temp_dir):
                    print(f"[*] 임시 디렉토리 정리 중: {temp_dir}")
                    shutil.rmtree(temp_dir)
        else:
            print("\n[*] ZIP 형식 문서(최신 Office) 확인. (정밀 압축해제 분석은 생략됨)")
            print("    '--deep' 옵션을 주면 강제 압축해제 정밀 분석을 수행합니다.")

            # Light mode: strings 체크
            print(f"\n{'='*60}")
            print(f"[+] 실행 중: 의심스러운 문자열 추출 (XLS - 경량 모드)")
            print(f"{'='*60}\n")

            all_strings = extract_strings(target_file)
            pattern = r'http|ftp|tcp|udp|powershell|cmd\.exe|\.vbs|\.exe|\.bat'
            matches = grep_patterns(all_strings, pattern)

            if matches:
                print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
                for m in matches:
                    print(f"    {m}")
            else:
                print("[*] 의심스러운 문자열이 발견되지 않았습니다.")
    else:
        # ZIP 형식이 아닌 경우 (Legacy .xls)
        print("\n[*] 파일이 ZIP 기반 아카이브가 아닙니다 (구형 .xls로 추정). 구조 분해 분석을 진행하지 않습니다.")

        print(f"\n{'='*60}")
        print(f"[+] 실행 중: 의심스러운 문자열 추출 (구형 XLS)")
        print(f"{'='*60}\n")

        all_strings = extract_strings(target_file)
        pattern = r'http|ftp|tcp|udp|powershell|cmd\.exe|\.vbs|\.exe|\.bat'
        matches = grep_patterns(all_strings, pattern)

        if matches:
            print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
            for m in matches:
                print(f"    {m}")
        else:
            print("[*] 의심스러운 문자열이 발견되지 않았습니다.")


def analyze_img(target_file):
    """
    이미지 파일 분석을 수행합니다.
    """
    # Step 1: 파일 타입 확인 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 파일 구조 및 타입 확인")
    print(f"{'='*60}\n")

    file_type_str = get_file_type(target_file)
    print(f"[*] 파일 확장자: {file_type_str}")

    # Step 2: exiftool 실행 (Windows 바이너리 사용)
    # Windows 환경의 exiftool은 유니코드 경로(한글, 베트남어 등)를 전달받을 때 글자가 깨지는 버그가 있습니다.
    # 이를 우회하기 위해 ASCII 경로(임시 폴더)로 파일을 복사한 후 분석하고 결과를 치환합니다.
    import tempfile
    import shutil
    print(f"\n{'='*60}")
    print(f"[+] Running: ExifTool Metadata Analysis")
    print(f"    Command: exiftool {target_file}")
    print(f"{'='*60}\n")
    try:
        tmp_dir = tempfile.mkdtemp()
        _, ext = os.path.splitext(target_file)
        tmp_file = os.path.join(tmp_dir, f"temp_exif{ext}")
        shutil.copy2(target_file, tmp_file)
        
        result = subprocess.run(["exiftool", tmp_file], capture_output=True, text=True, errors='replace', timeout=120)
        
        # 출력 결과에서 임시 파일 경로를 원래 파일 경로로 자연스럽게 치환
        out = result.stdout.replace(tmp_file, target_file).replace(tmp_file.replace("\\", "/"), target_file.replace("\\", "/"))
        out = out.replace("temp_exif" + ext, os.path.basename(target_file))
        if out: print(out)
        
        err = result.stderr.replace(tmp_file, target_file).replace(tmp_file.replace("\\", "/"), target_file.replace("\\", "/"))
        err = err.replace("temp_exif" + ext, os.path.basename(target_file))
        if err: print(f"[!] Output (stderr):\n{err}")
        
    except Exception as e:
        print(f"[!] Exiftool failed: {e}")
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass

    # Step 3: 의심스러운 문자열 추출 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 의심스러운 문자열 추출 (이미지)")
    print(f"{'='*60}\n")

    all_strings = extract_strings(target_file)
    pattern = r'html|script|\.zip|PK\x03\x04|<svg|<iframe|javascript'
    matches = grep_patterns(all_strings, pattern)

    if matches:
        print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
        for m in matches:
            print(f"    {m}")
    else:
        print("[*] 의심스러운 문자열이 발견되지 않았습니다.")


def analyze_exe(target_file):
    """
    EXE 파일 분석을 수행합니다.
    pefile(순수 Python)을 사용하여 헤더, 섹션, 임포트 정보를 분석합니다.
    """
    if not pefile:
        print("[!] pefile module not installed. Install with: pip install pefile")
        return

    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 윈도우 실행파일(PE) 분석 (pefile)")
    print(f"{'='*60}\n")

    try:
        pe = pefile.PE(target_file)

        # 1. Basic Information
        print(f"[*] 기본 정보:")
        print(f"    Machine: {hex(pe.FILE_HEADER.Machine)}")
        print(f"    TimeDateStamp: {pe.FILE_HEADER.TimeDateStamp} ({datetime.datetime.fromtimestamp(pe.FILE_HEADER.TimeDateStamp)})")
        print(f"    Subsystem: {hex(pe.OPTIONAL_HEADER.Subsystem)}")
        print(f"    EntryPoint: {hex(pe.OPTIONAL_HEADER.AddressOfEntryPoint)}")
        print(f"    ImageBase: {hex(pe.OPTIONAL_HEADER.ImageBase)}")

        # 2. Sections & Entropy
        print(f"\n[*] 섹션(Sections):")
        print(f"    {'Name':<10} {'Virtual Addr':<15} {'Raw Size':<10} {'Entropy':<10}")
        print(f"    {'-'*10} {'-'*15} {'-'*10} {'-'*10}")

        for section in pe.sections:
            name = section.Name.decode('utf-8', errors='ignore').strip('\x00')
            try:
                entropy = section.get_entropy()
            except AttributeError:
                entropy = 0.0

            print(f"    {name:<10} {hex(section.VirtualAddress):<15} {section.SizeOfRawData:<10} {entropy:.4f}")
            if entropy > 7.0:
                print(f"    [!] 높은 엔트로피(암호화/패킹 의심) 감지됨: {name} (possible packing/encryption)")

        # 3. Suspicious Imports
        print(f"\n[*] 의심스러운 임포트(API) 확인:")
        suspicious_apis = [
            'VirtualAlloc', 'VirtualProtect', 'CreateRemoteThread', 'WriteProcessMemory',
            'InternetOpen', 'URLDownloadToFile', 'ShellExecute', 'RegOpenKey',
            'GetProcAddress', 'LoadLibrary', 'NtCreateThread', 'RtlCreateUserThread',
            'WinExec', 'CreateProcess', 'OpenProcess'
        ]

        if hasattr(pe, 'DIRECTORY_ENTRY_IMPORT'):
            found_suspicious = False
            for entry in pe.DIRECTORY_ENTRY_IMPORT:
                dll_name = entry.dll.decode('utf-8', errors='ignore')
                for imp in entry.imports:
                    if imp.name:
                        func_name = imp.name.decode('utf-8', errors='ignore')
                        if any(s_api in func_name for s_api in suspicious_apis):
                            print(f"    [!] 의심스러운 API 발견: {func_name} ({dll_name})")
                            found_suspicious = True
            if not found_suspicious:
                print("    감시 목록에 있는 악성 행위 의심 API가 발견되지 않았습니다.")
        else:
            print("    DLL 임포트 테이블이 없습니다 (패킹되었거나 원시 코드일 확률이 높음).")

        pe.close()

    except Exception as e:
        print(f"[!] PE Analysis Failed: {e}")

    # 4. Strings Analysis (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 의심스러운 문자열 추출 (EXE)")
    print(f"{'='*60}\n")

    all_strings = extract_strings(target_file)
    pattern = r'https?://|\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|\.pdb|\.exe|\.dll|\.bat|\.ps1|powershell|cmd\.exe'
    matches = grep_patterns(all_strings, pattern)

    if matches:
        print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
        for m in matches:
            print(f"    {m}")
    else:
        print("[*] 의심스러운 문자열이 발견되지 않았습니다.")


def analyze_ppt(target_file):
    """
    PPT/PPTX 파일 분석을 수행합니다.
    """
    # Step 1: 파일 타입 확인 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 파일 구조 및 타입 확인")
    print(f"{'='*60}\n")

    file_type_str = get_file_type(target_file)
    print(f"[*] 파일 확장자: {file_type_str}")

    # Step 2: olevba 실행 (매크로 분석)
    run_external_command(
        ["olevba", target_file],
        "OLEVBA Macro Analysis"
    )

    # Step 3: 의심스러운 문자열 추출 (Python 내장)
    print(f"\n{'='*60}")
    print(f"[+] 실행 중: 의심스러운 문자열 추출 (PPT)")
    print(f"{'='*60}\n")

    all_strings = extract_strings(target_file)
    pattern = r'http|ftp|tcp|udp|powershell|cmd\.exe|\.vbs|\.exe|\.bat'
    matches = grep_patterns(all_strings, pattern)

    if matches:
        print(f"[*] {len(matches)}개의 의심스러운 문자열 발견:")
        for m in matches:
            print(f"    {m}")
    else:
        print("[*] 의심스러운 문자열이 발견되지 않았습니다.")
# ============================================================
# URL 평판 조회 함수
# ============================================================

def _get_daily_vt_count_file():
    """오늘 날짜의 VT API 조회 카운터 파일 경로를 반환합니다."""
    date_str = datetime.datetime.now().strftime("%y%m%d")
    attachfiles_dir = os.path.join(SCRIPT_DIR, "attachfiles")
    os.makedirs(attachfiles_dir, exist_ok=True)
    return os.path.join(attachfiles_dir, f".vt_api_count_{date_str}.txt")


def _get_daily_vt_count():
    """오늘 VT API 조회 횟수를 읽어옵니다."""
    count_file = _get_daily_vt_count_file()
    if os.path.exists(count_file):
        try:
            with open(count_file, "r") as f:
                return int(f.read().strip())
        except (ValueError, IOError):
            return 0
    return 0


def _increment_daily_vt_count():
    """오늘 VT API 조회 횟수를 1 증가시킵니다."""
    count = _get_daily_vt_count() + 1
    count_file = _get_daily_vt_count_file()
    with open(count_file, "w") as f:
        f.write(str(count))
    return int(count)


def analyze_url_reputation(url):
    """
    단일 URL에 대해 VirusTotal API로 평판 조회를 수행합니다.
    - 1차: GET으로 기존 이력 조회
    - 이력 없으면 2차: POST로 URL 신규 제출 → 분석 완료 대기 → 결과 조회
    """
    print(f"\n{'='*60}")
    print(f"[+] URL 평판 조회: {url}")
    print(f"{'='*60}\n")

    if not vt:
        print("[*] vt-py 모듈이 설치되지 않았습니다. VT URL 조회를 건너뜁니다.")
        print("    설치: pip install vt-py")
        return

    if not VT_API_KEY:
        print("[*] VT_API_KEY가 설정되지 않았습니다. VT URL 조회를 건너뜁니다.")
        return

    def _print_url_result(url_info):
        """VT URL 결과 출력 헬퍼."""
        try:
            stats = url_info.last_analysis_stats
            total = sum(stats.values())
            malicious = stats.get('malicious', 0)
            suspicious = stats.get('suspicious', 0)
            harmless = stats.get('harmless', 0)
            undetected = stats.get('undetected', 0)

            print(f"[+] 바이러스토탈 URL 검사 결과:")
            print(f"    Detection: {malicious}/{total} 엔진에서 악성으로 탐지됨")
            print(f"    의심: {suspicious}, 안전: {harmless}, 미탐지: {undetected}")

            if hasattr(url_info, 'url'):
                print(f"    대상 URL: {url_info.url}")
            if hasattr(url_info, 'last_final_url'):
                print(f"    최종 URL (리다이렉트): {url_info.last_final_url}")
            if hasattr(url_info, 'last_analysis_date'):
                print(f"    최근 분석: {url_info.last_analysis_date}")
            if hasattr(url_info, 'categories') and url_info.categories:
                cats = ', '.join(f"{k}: {v}" for k, v in url_info.categories.items())
                print(f"    카테고리: {cats}")

            if malicious > 0:
                print(f"\n[!] 악성 탐지 내역:")
                results = url_info.last_analysis_results
                for engine, result in results.items():
                    if result['category'] == 'malicious':
                        print(f"    [{engine}] {result.get('result', 'N/A')}")

            if hasattr(url_info, 'tags') and url_info.tags:
                print(f"\n[*] 태그: {', '.join(url_info.tags)}")

        except Exception as e:
            print(f"[!] 결과 출력 오류: {e}")

    try:
        url_id = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        client = vt.Client(VT_API_KEY)
        try:
            # ── 1차: 기존 이력 GET 조회 ──
            try:
                url_info = client.get_object(f"/urls/{url_id}")
                stats = url_info.last_analysis_stats
                total_detections = sum(stats.values()) if stats else 0

                if total_detections == 0:
                    # [SJ 병합] GET 이력 있으나 결과 미완료 → 5초 간격 최대 3회 재시도
                    print(f"[*] VT 기존 이력 있으나 결과 미완료 (통계 0). 5초 후 재시도...")
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries:
                        time.sleep(5)
                        try:
                            url_info = client.get_object(f"/urls/{url_id}")
                            if sum(url_info.last_analysis_stats.values()) > 0:
                                _print_url_result(url_info)
                                s = url_info.last_analysis_stats
                                return {"malicious": s.get('malicious', 0), "total": sum(s.values()), "is_cached": True}
                        except Exception:
                            pass
                        retry_count += 1
                        print(f"    [-] 재조회 중... ({retry_count}/{max_retries})")
                    # 재시도 실패 시 그냥 출력하고 진행
                    print(f"[!] {max_retries}회 재시도 후에도 완전한 결과를 가져오지 못했습니다.")
                    _print_url_result(url_info)
                    return {"malicious": stats.get('malicious', 0), "total": total_detections, "is_cached": True}
                else:
                    _print_url_result(url_info)
                    return {"malicious": stats.get('malicious', 0), "total": total_detections, "is_cached": True}

            except vt.error.APIError as e:
                if 'NotFoundError' not in str(e):
                    print(f"[!] VirusTotal API Error: {e}")
                else:
                    # ── 2차: DB 이력 없음 → POST로 신규 스캔 제출 ──
                    print(f"[*] VT DB 이력 없음. 신규 스캔 제출 중...")
                    try:
                        import requests as _requests
                        import json as _json

                        headers = {
                            "accept": "application/json",
                            "content-type": "application/x-www-form-urlencoded",
                            "x-apikey": VT_API_KEY
                        }
                        resp = _requests.post(
                            "https://www.virustotal.com/api/v3/urls",
                            headers=headers,
                            data=f"url={url}",
                            timeout=30
                        )
                        resp.raise_for_status()
                        analysis_id = resp.json()["data"]["id"]
                        print(f"[*] 스캔 제출 완료. 분석 ID: {analysis_id}")

                        # ── 3차: 분석 완료 + 유효한 결과 수신 대기 ──
                        VT_SCAN_WAIT = 8           # 폴링 간격(초)
                        VT_SCAN_HARD_TIMEOUT = 300 # 최대 총 대기 시간(초)
                        elapsed = 0
                        analysis_done = False
                        a_data = None

                        while elapsed < VT_SCAN_HARD_TIMEOUT:
                            print(f"[*] 분석 완료 대기 중... ({elapsed}/{VT_SCAN_HARD_TIMEOUT}초)")
                            time.sleep(VT_SCAN_WAIT)
                            elapsed += VT_SCAN_WAIT

                            analysis_resp = _requests.get(
                                f"https://www.virustotal.com/api/v3/analyses/{analysis_id}",
                                headers={"accept": "application/json", "x-apikey": VT_API_KEY},
                                timeout=30
                            )
                            analysis_resp.raise_for_status()
                            a_data = analysis_resp.json()
                            status = a_data.get("data", {}).get("attributes", {}).get("status", "")
                            stats  = a_data.get("data", {}).get("attributes", {}).get("stats", {})
                            total  = sum(stats.values())

                            if status == "completed" and total > 0:
                                # 유효한 결과 수신 완료
                                analysis_done = True
                                break
                            elif status == "completed" and total == 0:
                                # 완료됐지만 엔진 집계 미반영 → 계속 대기
                                print(f"[*] 분석 상태 completed이나 결과 미집계 (total=0). 재시도 중...")
                            else:
                                # queued / in-progress 상태
                                print(f"[*] 현재 분석 상태: {status}. 대기 계속...")

                        if analysis_done:
                            # ── 4차: 분석 완료 → URL 상세 결과 재조회 ──
                            try:
                                url_info = client.get_object(f"/urls/{url_id}")
                                _print_url_result(url_info)
                                # 결과 리턴
                                stats = url_info.last_analysis_stats
                                return {"malicious": stats.get('malicious', 0), "total": sum(stats.values()), "is_cached": False}
                            except Exception as fetch_e:
                                # get_object 실패 시 폴링에서 받은 a_data로 Fallback 출력
                                if a_data and "attributes" in a_data.get("data", {}):
                                    print(f"[*] 상세 정보 재조회 실패({fetch_e}). 분석 데이터로 대체 출력합니다.")
                                    attr  = a_data["data"]["attributes"]
                                    stats = attr.get("stats", {})
                                    total = sum(stats.values())
                                    mal   = stats.get('malicious', 0)
                                    sus   = stats.get('suspicious', 0)
                                    hml   = stats.get('harmless', 0)
                                    und   = stats.get('undetected', 0)
                                    print(f"[+] 바이러스토탈 URL 검사 결과 (Analysis ID 기점):")
                                    print(f"    Detection: {mal}/{total} 엔진에서 악성으로 탐지됨")
                                    print(f"    의심: {sus}, 안전: {hml}, 미탐지: {und}")
                                    print(f"    대상 URL: {url}")
                                    return {"malicious": mal, "total": total, "is_cached": False}
                                else:
                                    print(f"[!] 스캔 완료 후 결과 조회 실패: {fetch_e}")
                        else:
                            print(f"[!] {VT_SCAN_HARD_TIMEOUT}초 내에 유효한 결과를 받지 못했습니다. 다음 URL로 넘어갑니다.")

                    except Exception as post_e:
                        print(f"[!] VT URL 신규 스캔 제출 실패: {post_e}")

        finally:
            client.close()

    except Exception as e:
        print(f"[!] VirusTotal URL 조회 실패: {e}")
    
    return None



def analyze_urls_from_file(urls_file):
    """
    urls.txt 파일을 읽어 각 URL에 대해 VT 평판 조회를 수행합니다.
    - VT 무료 API 4 RPM 제한을 준수하기 위해 16초 간격으로 호출
    - 일일 최대 제한(MAX_DAILY_VT_LOOKUPS) 준수
    """
    if not os.path.exists(urls_file):
        print(f"[!] URL 파일을 찾을 수 없습니다: {urls_file}")
        return

    with open(urls_file, "r", encoding="utf-8") as f:
        urls = [line.strip() for line in f if line.strip()]

    if not urls:
        print("[*] 분석할 URL이 없습니다.")
        return

    print(f"[*] 총 {len(urls)}개 URL에 대해 VirusTotal 평판 조회를 시작합니다.")
    daily_count = _get_daily_vt_count()
    print(f"[*] 오늘 VT API 조회 누적: {daily_count}/{MAX_DAILY_VT_LOOKUPS}")

    VT_URL_DELAY = 16  # 4 RPM = 15초 + 안전 마진 1초

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(script_dir, "attachfiles", "vt_api_usage.log")
    os.makedirs(os.path.dirname(log_file), exist_ok=True)

    for i, url in enumerate(urls):
        # 이번 실행 중에 이미 분석한 URL인지 확인 (전역 필터)
        if url in GLOBAL_ANALYZED_URLS:
            print(f"  [-] 중복 URL 발견 (이번 세션 이미 분석됨): {url}")
            continue

        # 일일 한도 검사
        daily_count = _get_daily_vt_count()
        if daily_count >= MAX_DAILY_VT_LOOKUPS:
            remaining = len(urls) - i
            print(f"\n[!] 일일 VT API 조회 한도({MAX_DAILY_VT_LOOKUPS}건) 도달. 나머지 {remaining}개 URL 건너뜀.")
            break

        res = analyze_url_reputation(url)
        _increment_daily_vt_count()
        GLOBAL_ANALYZED_URLS.add(url) # 분석 완료 리스트에 추가

        # 진행 상황 파일 로그 기록 (결과 포함)
        try:
            with open(log_file, "a", encoding="utf-8") as lf:
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 결과 요약 태그 생성
                tag = "[UNKNOWN]"
                stats_str = "N/A"
                if res:
                    mal = res.get('malicious', 0)
                    total = res.get('total', 0)
                    stats_str = f"{mal}/{total}"
                    if mal >= 3: tag = "[MALICIOUS]"
                    elif mal > 0: tag = "[SUSPICIOUS]"
                    else: tag = "[CLEAN]"
                
                msg_suffix = f"(대기 {VT_URL_DELAY}초 후 진행 예정)" if i < len(urls) - 1 else "(마지막 URL)"
                lf.write(f"[{now_str}] [{i+1}/{len(urls)}] {tag} ({stats_str}) {url} {msg_suffix}\n")
        except Exception as e:
            print(f"[!] VT 로그 기록 실패: {e}")

        # [SJ 병합] 다음 URL 대기 (캐시 분기: 신규 스캔만 16초 대기)
        is_cached = res.get('is_cached', False) if res else False
        if not is_cached and i < len(urls) - 1:
            print(f"[*] 신규 스캔 발생: VT API rate limit 대기 {VT_URL_DELAY}초...")
            time.sleep(VT_URL_DELAY)
        elif is_cached and i < len(urls) - 1:
            print(f"[*] 기존 이력 확인됨: 대기 없이 다음으로 진행합니다.")

    final_count = _get_daily_vt_count()
    print(f"\n[*] URL 평판 조회 완료. 오늘 누적: {final_count}/{MAX_DAILY_VT_LOOKUPS}")


# ============================================================
# 메인 함수
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="Automated File Analysis Tool (Windows Compatible)")

    # 상호 배타적인 인자 그룹 생성
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-pdf", dest="pdf_filename", help="Path to the suspicious PDF file")
    group.add_argument("-xls", dest="xls_filename", help="Path to the suspicious Excel file")
    group.add_argument("-ppt", dest="ppt_filename", help="Path to the suspicious PowerPoint file")
    group.add_argument("-img", dest="img_filename", help="Path to the suspicious Image file")
    group.add_argument("-exe", dest="exe_filename", help="Path to the suspicious EXE file")
    group.add_argument("-file", dest="generic_filename", help="Path to any suspicious file (auto-detect)")
    group.add_argument("-urls", dest="urls_filename", help="Path to urls.txt for URL reputation analysis")

    parser.add_argument("-out", dest="output_dir", default=None,
                        help="Directory to save analysis result (default: same as target file)")
    parser.add_argument("--deep", "-d", dest="deep_analysis", action="store_true",
                        help="Enable deep analysis (Unzip & Recursive search for Excel files)")

    args = parser.parse_args()

    target_file = None
    file_type = None

    # URL 분석 모드
    if args.urls_filename:
        urls_file = os.path.abspath(args.urls_filename)
        if not os.path.exists(urls_file):
            print(f"[!] URL file not found: {urls_file}")
            sys.exit(1)

        # Setup Logging
        save_dir = args.output_dir if args.output_dir else os.path.dirname(urls_file)
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        try:
            date_str = datetime.datetime.now().strftime("%y%m%d")
            log_filename = f"{date_str}_url_analysis.md"
            log_full_path = os.path.join(save_dir, log_filename)
            sys.stdout = Logger(log_full_path)
            print(f"[*] URL 평판 조회 분석 시작 일시: {datetime.datetime.now()}")
            print(f"[*] 분석 결과물 저장 경로: {log_full_path}")
        except Exception as e:
            print(f"[!] Logging setup failed: {e}")

        analyze_urls_from_file(urls_file)
        return

    if args.pdf_filename:
        target_file = args.pdf_filename
        file_type = 'pdf'
    elif args.xls_filename:
        target_file = args.xls_filename
        file_type = 'xls'
    elif args.img_filename:
        target_file = args.img_filename
        file_type = 'img'
    elif args.exe_filename:
        target_file = args.exe_filename
        file_type = 'exe'
    elif args.ppt_filename:
        target_file = args.ppt_filename
        file_type = 'ppt'
    elif args.generic_filename:
        target_file = args.generic_filename
        # Auto-detect file type based on extension
        _, ext = os.path.splitext(target_file)
        ext = ext.lower()

        if ext in ['.pdf']:
            file_type = 'pdf'
        elif ext in ['.xls', '.xlsx', '.xlsm', '.doc', '.docx', '.docm']:
            file_type = 'xls'
        elif ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.jfif']:
            file_type = 'img'
        elif ext in ['.exe', '.dll', '.sys', '.ocx']:
            file_type = 'exe'
        elif ext in ['.ppt', '.pptx', '.pptm', '.potx', '.pps', '.ppsx']:
            file_type = 'ppt'
        else:
            print(f"[!] Unsupported file extension: {ext}")
            sys.exit(1)

    if not os.path.exists(target_file):
        print(f"[!] File not found: {target_file}")
        sys.exit(1)

    # Setup Logging
    save_dir = args.output_dir if args.output_dir else os.path.dirname(os.path.abspath(target_file))
    if not os.path.exists(save_dir):
        try:
            os.makedirs(save_dir)
        except OSError as e:
            print(f"[!] Could not create output directory: {e}")
            sys.exit(1)

    try:
        date_str = datetime.datetime.now().strftime("%y%m%d")
        base_name = os.path.splitext(os.path.basename(target_file))[0]
        log_filename = f"{date_str}_{base_name}_analysis.md"
        log_full_path = os.path.join(save_dir, log_filename)

        # Redirect stdout to Logger
        sys.stdout = Logger(log_full_path)

        print(f"[*] 로컬 정적 분석 시작 일시: {datetime.datetime.now()}")
        print(f"[*] 로컬 분석 결과물 저장 경로: {log_full_path}")
    except Exception as e:
        print(f"[!] Logging setup failed: {e}")

    print(f"[*] 분석 대상 파일: {target_file} ({file_type})")

    # Common Step: Hash Analysis
    analyze_hash(target_file)

    if file_type == 'pdf':
        analyze_pdf(target_file)
    elif file_type == 'xls':
        analyze_xls(target_file, args.deep_analysis)
    elif file_type == 'img':
        analyze_img(target_file)
    elif file_type == 'exe':
        analyze_exe(target_file)
    elif file_type == 'ppt':
        analyze_ppt(target_file)


def analyze_file_as_dict(target_file, file_type=None, deep_analysis=False):
    """
    Subprocess 로그 파싱 대신 파이썬 모듈로서 직접 호출되어 
    분석 결과를 Dictionary 형태로 반환하는 래퍼 함수입니다.
    """
    import io, sys, os
    
    if file_type is None:
        _, ext = os.path.splitext(target_file)
        ext = ext.lower()
        if ext in ['.pdf']:
            file_type = 'pdf'
        elif ext in ['.xls', '.xlsx', '.xlsm', '.doc', '.docx', '.docm']:
            file_type = 'xls'
        elif ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.jfif']:
            file_type = 'img'
        elif ext in ['.exe', '.dll', '.sys', '.ocx']:
            file_type = 'exe'
        elif ext in ['.ppt', '.pptx', '.pptm', '.potx', '.pps', '.ppsx']:
            file_type = 'ppt'
        else:
            return {"status": "error", "message": f"Unsupported file extension: {ext}"}

    old_stdout = sys.stdout
    old_stderr = sys.stderr
    captured_out = io.StringIO()
    # 터미널 출력을 메모리 버퍼로 리다이렉션
    sys.stdout = captured_out
    sys.stderr = captured_out
    
    try:
        analyze_hash(target_file)
        if file_type == 'pdf':
            analyze_pdf(target_file)
        elif file_type == 'xls':
            analyze_xls(target_file, deep_analysis)
        elif file_type == 'img':
            analyze_img(target_file)
        elif file_type == 'exe':
            analyze_exe(target_file)
        elif file_type == 'ppt':
            analyze_ppt(target_file)
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        
    return {
        "status": "success",
        "file_path": target_file,
        "file_type": file_type,
        "raw_analysis_log": captured_out.getvalue()
    }


def analyze_urls_as_dict(urls_file):
    """
    URL 분석 결과를 Dictionary 형태로 반환하는 래퍼 함수입니다.
    """
    import io, sys
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    captured_out = io.StringIO()
    sys.stdout = captured_out
    sys.stderr = captured_out
    
    try:
        analyze_urls_from_file(urls_file)
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        
    return {
        "status": "success",
        "urls_file": urls_file,
        "raw_analysis_log": captured_out.getvalue()
    }


if __name__ == "__main__":
    main()