#!/usr/bin/env python3
"""
통합 파이프라인 실행 스크립트
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1단계: auto_login.py  -> 신한메일 로그인 + EML 다운로드
2단계: ai_analysis.py -> 첨부파일 추출 + 로컬분석 + AI 보고서 생성
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import subprocess
import sys
import os
import shutil
import datetime

# Windows cp949 인코딩 문제 방지
os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

# 스크립트가 위치한 디렉토리를 기준으로 경로 설정
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE = sys.executable

AUTO_LOGIN_SCRIPT = os.path.join(SCRIPT_DIR, "auto_login.py")
EXTRACT_SCRIPT = os.path.join(SCRIPT_DIR, "extract_attachments.py")
AI_ANALYSIS_SCRIPT = os.path.join(SCRIPT_DIR, "ai_analysis.py")


def print_banner():
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print()
    print("=" * 55)
    print("  [Mail] EML 수집 + [AI] 파일 분석 통합 파이프라인")
    print(f"  실행 시각: {now}")
    print("=" * 55)
    print()


def run_step(step_num, step_name, script_path, extra_args=None):
    """하위 스크립트를 실행하고 결과를 반환합니다."""
    print(f"\n{'-' * 55}")
    print(f"  [{step_num}단계] {step_name}")
    print(f"  스크립트: {os.path.basename(script_path)}")
    print(f"{'-' * 55}\n")

    if not os.path.exists(script_path):
        print(f"[!] 스크립트를 찾을 수 없습니다: {script_path}")
        return False

    cmd = [PYTHON_EXE, "-u", script_path]
    if extra_args:
        cmd.extend(extra_args)

    try:
        result = subprocess.run(
            cmd,
            cwd=SCRIPT_DIR,
            env=os.environ.copy()
        )

        if result.returncode == 0:
            print(f"\n[OK] [{step_num}단계] {step_name} 완료!")
            return True
        else:
            print(f"\n[!] [{step_num}단계] {step_name} 실패 (Exit code: {result.returncode})")
            return False

    except KeyboardInterrupt:
        print(f"\n[!] 사용자가 {step_name}을 중단했습니다.")
        return False
    except Exception as e:
        print(f"\n[!] {step_name} 실행 중 오류 발생: {e}")
        return False


def main():
    print_banner()

    # ━━━━ 0단계: 기존 데이터 초기화 여부 확인 ━━━━
    eml_dir = os.path.join(SCRIPT_DIR, "eml")
    attachfiles_dir = os.path.join(SCRIPT_DIR, "attachfiles")
    history_files = [
        os.path.join(SCRIPT_DIR, "attachfiles", "download_history.txt"),
        os.path.join(SCRIPT_DIR, "attachfiles", "downloaded_hash_history.txt"),
        os.path.join(SCRIPT_DIR, "attachfiles", "extracted_hash_history.txt"),
    ]
    
    print("[!] 알림: 시간날때 totalurls.txt 검토하여 safe_domains.txt에 적용해주세요. 적용 후, totalurls.txt 삭제해주세요.")
    print("[?] 기존 메일과 첨부파일을 삭제하시겠습니까?")
    print("    Y: eml 폴더, attachfiles 폴더, 다운로드/분석 이력 파일을 모두 초기화합니다.")
    print("    N: 기존 데이터를 유지하고 새로운 데이터만 추가합니다.")
    reset_answer = input(">> 선택 (Y/N): ").strip().lower()
    
    if reset_answer == 'y':
        # eml 폴더 내부 파일 삭제
        if os.path.exists(eml_dir):
            print(f"[*] EML 폴더 초기화 중... ({eml_dir})")
            for f in os.listdir(eml_dir):
                fpath = os.path.join(eml_dir, f)
                try:
                    if os.path.isfile(fpath):
                        os.remove(fpath)
                    elif os.path.isdir(fpath):
                        shutil.rmtree(fpath)
                except Exception:
                    pass
        
        # attachfiles 폴더 내부 삭제
        if os.path.exists(attachfiles_dir):
            print(f"[*] 첨부파일 폴더 초기화 중... ({attachfiles_dir})")
            for f in os.listdir(attachfiles_dir):
                # API 사용량 카운터 파일은 삭제 방지
                if f.startswith(".vt_api_count_"):
                    continue

                fpath = os.path.join(attachfiles_dir, f)
                try:
                    if os.path.isfile(fpath):
                        os.remove(fpath)
                    elif os.path.isdir(fpath):
                        shutil.rmtree(fpath)
                except Exception:
                    pass
        
        # 이력 파일 삭제
        for hf in history_files:
            if os.path.exists(hf):
                try:
                    os.remove(hf)
                    print(f"[*] 이력 파일 삭제: {os.path.basename(hf)}")
                except Exception:
                    pass
        
        print("[*] 초기화 완료!\n")
    else:
        print("[*] 기존 데이터를 유지합니다.\n")
    
    # ━━━━ 1단계: 메일 로그인 + EML 다운로드 ━━━━
    step1_ok = run_step(
        step_num=1,
        step_name="신한메일 로그인 & EML 다운로드",
        script_path=AUTO_LOGIN_SCRIPT,
        extra_args=["--auto-close"]
    )

    if not step1_ok:
        print("\n[!] 1단계가 실패했습니다.")
        answer = input("[?] 그래도 2단계(첨부파일 추출)를 계속 진행하시겠습니까? (y/N): ").strip().lower()
        if answer != 'y':
            print("[*] 파이프라인을 종료합니다.")
            return

    # ━━━━ 2단계: 첨부파일 추출 + URL 추출 ━━━━
    step2_ok = run_step(
        step_num=2,
        step_name="첨부파일 추출 & URL 추출",
        script_path=EXTRACT_SCRIPT
    )

    # ━━━━ 2.5단계: totalurls.txt에 URL 누적 수집 ━━━━
    print(f"\n{'-' * 55}")
    print(f"  [2.5단계] totalurls.txt URL 누적 수집")
    print(f"{'-' * 55}\n")

    totalurls_path = os.path.join(SCRIPT_DIR, "totalurls.txt")

    # 기존 totalurls.txt에 이미 있는 URL 로드 (중복 방지)
    existing_urls = set()
    if os.path.exists(totalurls_path):
        with open(totalurls_path, "r", encoding="utf-8") as f:
            existing_urls = set(line.strip() for line in f if line.strip())

    prev_count = len(existing_urls)

    # attachfiles 하위 모든 urls.txt 수집
    for root, dirs, files in os.walk(attachfiles_dir):
        if "urls.txt" in files:
            urls_file = os.path.join(root, "urls.txt")
            try:
                with open(urls_file, "r", encoding="utf-8") as uf:
                    for line in uf:
                        url = line.strip()
                        if url:
                            existing_urls.add(url)
            except Exception as e:
                print(f"  [!] {urls_file} 읽기 실패: {e}")

    new_count = len(existing_urls) - prev_count

    # 전체 목록을 정렬하여 저장
    with open(totalurls_path, "w", encoding="utf-8") as f:
        for url in sorted(existing_urls):
            f.write(url + "\n")

    print(f"[OK] totalurls.txt 업데이트: 신규 {new_count}건 추가 (총 {len(existing_urls)}건)")
    print(f"     경로: {totalurls_path}")

    # ━━━━ 3단계: 첨부파일 분석 + URL 분석 + AI 보고서 생성 ━━━━
    step3_ok = run_step(
        step_num=3,
        step_name="파일 분석 & URL 분석 & AI 보고서 생성",
        script_path=AI_ANALYSIS_SCRIPT
    )

    # ━━━━ 최종 결과 ━━━━
    print()
    print("=" * 55)
    all_ok = step1_ok and step2_ok and step3_ok
    if all_ok:
        print("  [SUCCESS] 전체 파이프라인이 성공적으로 완료되었습니다!")
    else:
        print("  [WARNING] 일부 단계에서 문제가 발생했습니다.")
        if not step1_ok:
            print("    - 1단계 (메일 다운로드): 실패")
        if not step2_ok:
            print("    - 2단계 (첨부파일/URL 추출): 실패")
        if not step3_ok:
            print("    - 3단계 (AI 분석): 실패")
    print("=" * 55)
    print()


if __name__ == "__main__":
    main()
