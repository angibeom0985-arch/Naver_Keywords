#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
파일 및 폴더 보호 스크립트
생성된 키워드 추출 결과 파일들과 폴더들을 보호하여
관리자 권한으로만 삭제할 수 있도록 설정합니다.
"""

import os
import sys
import stat
import subprocess
import ctypes
from pathlib import Path

def is_admin():
    """현재 프로세스가 관리자 권한으로 실행되었는지 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """관리자 권한으로 스크립트 재실행"""
    if is_admin():
        return True
    else:
        # 관리자 권한으로 재실행
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            return False
        except:
            print("❌ 관리자 권한으로 실행할 수 없습니다.")
            return False

def protect_file_with_attrib(file_path):
    """Windows attrib 명령을 사용하여 파일 보호"""
    try:
        # 시스템 파일 + 숨김 파일 + 읽기 전용 속성 설정
        cmd = f'attrib +S +H +R "{file_path}"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f"✅ 파일 보호 완료: {file_path}")
            return True
        else:
            print(f"⚠️ 파일 보호 실패: {file_path} - {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ 파일 보호 오류: {file_path} - {str(e)}")
        return False

def protect_folder_with_attrib(folder_path):
    """Windows attrib 명령을 사용하여 폴더 보호"""
    try:
        # 시스템 폴더 + 읽기 전용 속성 설정
        cmd = f'attrib +S +R "{folder_path}"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f"✅ 폴더 보호 완료: {folder_path}")
            return True
        else:
            print(f"⚠️ 폴더 보호 실패: {folder_path} - {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ 폴더 보호 오류: {folder_path} - {str(e)}")
        return False

def set_file_permissions(file_path):
    """파일 권한 설정 - 읽기 전용으로 설정"""
    try:
        # 읽기 전용 속성 설정
        os.chmod(file_path, stat.S_IRUSR | stat.S_IRGRP | stat.S_IROTH)
        print(f"✅ 파일 권한 설정 완료: {file_path}")
        return True
    except Exception as e:
        print(f"⚠️ 파일 권한 설정 실패: {file_path} - {str(e)}")
        return False

def protect_keyword_results():
    """키워드 추출 결과 폴더와 파일들 보호"""
    
    keyword_results_dir = "keyword_results"
    
    # 키워드 결과 폴더 존재 확인
    if not os.path.exists(keyword_results_dir):
        print(f"⚠️ {keyword_results_dir} 폴더가 존재하지 않습니다.")
        return False
    
    protected_count = 0
    
    print(f"🛡️ {keyword_results_dir} 폴더 보호 중...")
    
    # 폴더 자체 보호
    if protect_folder_with_attrib(keyword_results_dir):
        protected_count += 1
    
    # 폴더 내 모든 파일 보호
    try:
        for root, dirs, files in os.walk(keyword_results_dir):
            # 하위 폴더들 보호
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                if protect_folder_with_attrib(dir_path):
                    protected_count += 1
            
            # 파일들 보호
            for file_name in files:
                file_path = os.path.join(root, file_name)
                if file_name.endswith(('.xlsx', '.csv', '.json', '.txt')):
                    # 중요한 파일들은 이중 보호
                    if protect_file_with_attrib(file_path):
                        protected_count += 1
                    set_file_permissions(file_path)
                else:
                    # 기타 파일들
                    if protect_file_with_attrib(file_path):
                        protected_count += 1
    
    except Exception as e:
        print(f"❌ 폴더 순회 중 오류: {str(e)}")
        return False
    
    print(f"✅ 총 {protected_count}개의 항목이 보호되었습니다.")
    return True

def protect_dist_folder():
    """dist 폴더 보호 (exe 파일들이 있는 경우)"""
    
    dist_dir = "dist"
    
    if not os.path.exists(dist_dir):
        print(f"⚠️ {dist_dir} 폴더가 존재하지 않습니다.")
        return False
    
    protected_count = 0
    
    print(f"🛡️ {dist_dir} 폴더 보호 중...")
    
    # dist 폴더 자체 보호
    if protect_folder_with_attrib(dist_dir):
        protected_count += 1
    
    # dist 폴더 내 파일들 보호
    try:
        for root, dirs, files in os.walk(dist_dir):
            # 하위 폴더들 보호
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                if protect_folder_with_attrib(dir_path):
                    protected_count += 1
            
            # 파일들 보호
            for file_name in files:
                file_path = os.path.join(root, file_name)
                if file_name.endswith('.exe'):
                    # EXE 파일들은 특별 보호
                    if protect_file_with_attrib(file_path):
                        protected_count += 1
                    print(f"🔒 EXE 파일 보호: {file_path}")
                elif file_name.endswith(('.xlsx', '.csv', '.json', '.txt')):
                    # 결과 파일들 보호
                    if protect_file_with_attrib(file_path):
                        protected_count += 1
                else:
                    # 기타 파일들
                    if protect_file_with_attrib(file_path):
                        protected_count += 1
    
    except Exception as e:
        print(f"❌ dist 폴더 처리 중 오류: {str(e)}")
        return False
    
    print(f"✅ dist 폴더에서 총 {protected_count}개의 항목이 보호되었습니다.")
    return True

def create_unprotect_script():
    """보호 해제 스크립트 생성 (관리자용)"""
    
    unprotect_script = """@echo off
echo ====================================================
echo       🔓 파일 보호 해제 스크립트 (관리자 전용)
echo ====================================================
echo.
echo ⚠️  이 스크립트는 보호된 파일들을 해제합니다.
echo 계속하시겠습니까? (y/n)
set /p choice=선택: 
if /i not "%choice%"=="y" (
    echo 취소되었습니다.
    pause
    exit /b
)

echo.
echo 🔓 keyword_results 폴더 보호 해제 중...
if exist "keyword_results" (
    attrib -S -H -R "keyword_results" /S /D
    echo ✅ keyword_results 폴더 보호 해제 완료
) else (
    echo ⚠️ keyword_results 폴더가 존재하지 않습니다.
)

echo.
echo 🔓 dist 폴더 보호 해제 중...
if exist "dist" (
    attrib -S -H -R "dist" /S /D
    echo ✅ dist 폴더 보호 해제 완료
) else (
    echo ⚠️ dist 폴더가 존재하지 않습니다.
)

echo.
echo ✅ 모든 파일 보호가 해제되었습니다.
echo 이제 파일들을 삭제하거나 수정할 수 있습니다.
echo.
pause
"""
    
    try:
        with open("unprotect_files.bat", "w", encoding="utf-8") as f:
            f.write(unprotect_script)
        
        # 배치 파일도 보호
        protect_file_with_attrib("unprotect_files.bat")
        
        print("✅ 보호 해제 스크립트가 생성되었습니다: unprotect_files.bat")
        print("   (관리자 권한으로 실행해야 합니다)")
        return True
        
    except Exception as e:
        print(f"❌ 보호 해제 스크립트 생성 실패: {str(e)}")
        return False

def main():
    """메인 실행 함수"""
    print("🛡️ 파일 보호 스크립트 시작")
    print("=" * 50)
    
    # 관리자 권한 확인
    if not is_admin():
        print("⚠️ 이 스크립트는 관리자 권한이 필요합니다.")
        print("관리자 권한으로 재실행을 시도합니다...")
        
        if not run_as_admin():
            input("❌ 관리자 권한으로 실행할 수 없습니다. 엔터 키를 눌러 종료하세요...")
            return False
        else:
            return True  # 재실행됨
    
    print("✅ 관리자 권한으로 실행 중")
    print()
    
    success_count = 0
    
    # 1. 키워드 결과 폴더 보호
    if protect_keyword_results():
        success_count += 1
    
    # 2. dist 폴더 보호 (있는 경우)
    if protect_dist_folder():
        success_count += 1
    
    # 3. 보호 해제 스크립트 생성
    if create_unprotect_script():
        success_count += 1
    
    print("\n" + "=" * 50)
    if success_count > 0:
        print("🎉 파일 보호가 완료되었습니다!")
        print()
        print("📋 보호된 내용:")
        print("   • keyword_results 폴더와 모든 하위 파일")
        print("   • dist 폴더와 모든 하위 파일 (있는 경우)")
        print("   • 시스템 파일 속성으로 보호")
        print("   • 읽기 전용 속성으로 보호")
        print()
        print("🔓 보호 해제 방법:")
        print("   1. unprotect_files.bat를 관리자 권한으로 실행")
        print("   2. 또는 명령 프롬프트(관리자)에서:")
        print("      attrib -S -H -R keyword_results /S /D")
        print("      attrib -S -H -R dist /S /D")
    else:
        print("❌ 파일 보호에 실패했습니다.")
    
    input("\n엔터 키를 눌러 종료하세요...")
    return success_count > 0

if __name__ == "__main__":
    main() 